import assert from 'node:assert/strict';
import { createHash } from 'node:crypto';
import fs from 'node:fs/promises';
import os from 'node:os';
import test from 'node:test';
import path from 'node:path';
import { ServerDownloadManager, relativePathPartsFromGraphPath, resolveInsideRoot, sanitizeUserFolder } from '../src/server/serverDownloadEngine.ts';
import { ServerStateStore } from '../src/server/serverStateStore.ts';
import type { RemoteItemMetadata } from '../src/download/types.ts';

test('resolves server download paths inside the configured root', () => {
  const root = path.resolve('/tmp/syncpoint-root');
  assert.equal(
    resolveInsideRoot(root, 'folder/report.txt'),
    path.join(root, 'folder', 'report.txt'),
  );
});

test('rejects server download paths escaping the configured root', () => {
  const root = path.resolve('/tmp/syncpoint-root');
  assert.throws(() => resolveInsideRoot(root, '../escape.txt'), /escapes/);
  assert.throws(() => resolveInsideRoot(root, '/etc/passwd'), /escapes/);
});

test('sanitizes user emails for isolated server folders', () => {
  assert.equal(sanitizeUserFolder('Owner@Example.COM'), 'owner_example.com');
  assert.equal(sanitizeUserFolder('bad/user@example.com'), 'bad_user_example.com');
  assert.equal(sanitizeUserFolder('***'), 'user');
});

test('extracts OneDrive relative paths for nested selections', () => {
  assert.deepEqual(
    relativePathPartsFromGraphPath('/drive/root:/Imagini/Emergency'),
    ['Imagini', 'Emergency'],
  );
  assert.deepEqual(
    relativePathPartsFromGraphPath('/drive/root:/Imagini/Emergency/photo.jpg', 'photo.jpg'),
    ['Imagini', 'Emergency'],
  );
});

function sha1(value: string) {
  return createHash('sha1').update(value).digest('hex').toUpperCase();
}

async function waitForJob(
  manager: ServerDownloadManager,
  userEmail: string,
  id: string,
) {
  for (let attempt = 0; attempt < 50; attempt += 1) {
    const job = manager.getJob(userEmail, id);
    if (job && ['completed', 'failed', 'cancelled'].includes(job.status)) return job;
    await new Promise(resolve => setTimeout(resolve, 20));
  }
  throw new Error('Timed out waiting for server job.');
}

function installFetchMock(item: RemoteItemMetadata, content: string) {
  const originalFetch = globalThis.fetch;
  globalThis.fetch = async (input: RequestInfo | URL) => {
    const url = String(input);
    if (url.includes('/me/drive/items/item-1')) {
      return new Response(JSON.stringify({
        id: item.itemId,
        name: item.name,
        size: item.size,
        eTag: item.eTag,
        cTag: item.cTag,
        lastModifiedDateTime: item.lastModifiedDateTime,
        parentReference: { driveId: item.driveId, path: '/drive/root:/Documents' },
        file: { hashes: item.hashes },
        '@microsoft.graph.downloadUrl': 'https://download.example/report.txt',
      }), { status: 200, headers: { 'content-type': 'application/json' } });
    }
    if (url === 'https://download.example/report.txt') {
      return new Response(content, { status: 200 });
    }
    return new Response('not found', { status: 404 });
  };
  return () => {
    globalThis.fetch = originalFetch;
  };
}

test('server skip_existing verifies same-size files before skipping', async () => {
  const root = await fs.mkdtemp(path.join(os.tmpdir(), 'onedrive-archiver-root-'));
  const data = await fs.mkdtemp(path.join(os.tmpdir(), 'onedrive-archiver-data-'));
  try {
    const store = new ServerStateStore(data);
    const manager = new ServerDownloadManager(root, store);
    const user = { email: 'owner@example.com' };
    const targetFile = path.join(root, 'users', 'owner_example.com', 'Documents', 'report.txt');
    await fs.mkdir(path.dirname(targetFile), { recursive: true });
    await fs.writeFile(targetFile, 'good');
    const restoreFetch = installFetchMock({
      driveId: 'me',
      itemId: 'item-1',
      name: 'report.txt',
      remotePath: 'Documents/report.txt',
      size: 4,
      hashes: { sha1Hash: sha1('good') },
    }, 'good');

    try {
      const job = manager.start(user, 'start', [{ id: 'item-1', name: 'report.txt', type: 'file', sourcePath: '/drive/root:/Documents/report.txt' }], {
        conflictStrategy: 'skip_existing',
        verifyAfterDownload: true,
      }, async () => 'token');
      const completed = await waitForJob(manager, user.email, job.id);
      assert.equal(completed.status, 'completed');
      assert.match(completed.log.join('\n'), /Skipped existing verified file/);
      assert.equal(store.listArchiveRecords(user.email).length, 1);
    } finally {
      restoreFetch();
    }
  } finally {
    await fs.rm(root, { recursive: true, force: true });
    await fs.rm(data, { recursive: true, force: true });
  }
});

test('server skip_existing redownloads same-size files that fail verification', async () => {
  const root = await fs.mkdtemp(path.join(os.tmpdir(), 'onedrive-archiver-root-'));
  const data = await fs.mkdtemp(path.join(os.tmpdir(), 'onedrive-archiver-data-'));
  try {
    const store = new ServerStateStore(data);
    const manager = new ServerDownloadManager(root, store);
    const user = { email: 'owner@example.com' };
    const targetFile = path.join(root, 'users', 'owner_example.com', 'Documents', 'report.txt');
    await fs.mkdir(path.dirname(targetFile), { recursive: true });
    await fs.writeFile(targetFile, 'bad!');
    const item: RemoteItemMetadata = {
      driveId: 'me',
      itemId: 'item-1',
      name: 'report.txt',
      remotePath: 'Documents/report.txt',
      size: 4,
      hashes: { sha1Hash: sha1('good') },
    };
    const restoreFetch = installFetchMock(item, 'good');

    try {
      const job = manager.start(user, 'start', [{ id: 'item-1', name: 'report.txt', type: 'file', sourcePath: '/drive/root:/Documents/report.txt' }], {
        conflictStrategy: 'skip_existing',
        verifyAfterDownload: true,
      }, async () => 'token');
      const completed = await waitForJob(manager, user.email, job.id);
      assert.equal(completed.status, 'completed');
      assert.equal(await fs.readFile(targetFile, 'utf8'), 'good');
      assert.match(completed.log.join('\n'), /Existing file failed verification; redownloading/);
    } finally {
      restoreFetch();
    }
  } finally {
    await fs.rm(root, { recursive: true, force: true });
    await fs.rm(data, { recursive: true, force: true });
  }
});

test('server repair without selections uses saved archive records', async () => {
  const root = await fs.mkdtemp(path.join(os.tmpdir(), 'onedrive-archiver-root-'));
  const data = await fs.mkdtemp(path.join(os.tmpdir(), 'onedrive-archiver-data-'));
  try {
    const store = new ServerStateStore(data);
    const manager = new ServerDownloadManager(root, store);
    const user = { email: 'owner@example.com' };
    const item: RemoteItemMetadata = {
      driveId: 'me',
      itemId: 'item-1',
      name: 'report.txt',
      remotePath: 'Documents/report.txt',
      size: 4,
      hashes: { sha1Hash: sha1('good') },
    };
    store.upsertArchiveRecord({
      userEmail: user.email,
      localPath: path.join(root, 'users', 'owner_example.com', item.remotePath),
      status: 'completed',
      item,
    });
    const restoreFetch = installFetchMock(item, 'good');

    try {
      const job = manager.start(user, 'repair', [], {
        conflictStrategy: 'overwrite',
        verifyAfterDownload: true,
      }, async () => 'token');
      const completed = await waitForJob(manager, user.email, job.id);
      assert.equal(completed.status, 'completed');
      assert.equal(await fs.readFile(path.join(root, 'users', 'owner_example.com', 'Documents', 'report.txt'), 'utf8'), 'good');
      assert.match(completed.log.join('\n'), /Loaded 1 saved server archive record/);
    } finally {
      restoreFetch();
    }
  } finally {
    await fs.rm(root, { recursive: true, force: true });
    await fs.rm(data, { recursive: true, force: true });
  }
});
