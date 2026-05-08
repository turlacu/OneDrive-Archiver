import assert from 'node:assert/strict';
import test from 'node:test';
import { createManifest, manifestItemFromJob } from '../src/download/runManifest.ts';
import { defaultDownloadSettings, type DownloadJob } from '../src/download/types.ts';

function job(patch: Partial<DownloadJob> = {}): DownloadJob {
  return {
    id: 'drive:item:folder/report.txt',
    driveId: 'drive',
    itemId: 'item',
    name: 'report.txt',
    remotePath: 'drive/root:/folder/report.txt',
    localPath: 'folder/report.txt',
    partialPath: 'folder/report.txt.partial',
    size: 12,
    eTag: 'etag',
    cTag: 'ctag',
    hashes: { sha1Hash: 'sha1' },
    status: 'completed',
    priority: 0,
    retryCount: 0,
    createdAt: new Date(0).toISOString(),
    updatedAt: new Date(1).toISOString(),
    downloadedBytes: 12,
    chunks: [],
    ...patch,
  };
}

test('creates manifest items from completed jobs', () => {
  const item = manifestItemFromJob('run', job(), 'completed', {
    verification: 'SHA-1 verification passed.',
    finalName: 'report.txt',
  });

  assert.equal(item.runId, 'run');
  assert.equal(item.status, 'completed');
  assert.equal(item.remotePath, 'drive/root:/folder/report.txt');
  assert.equal(item.localPath, 'folder/report.txt');
  assert.equal(item.sha1Hash, 'sha1');
  assert.equal(item.verification, 'SHA-1 verification passed.');
  assert.equal(item.finalName, 'report.txt');
});

test('creates run manifest summary with archive metadata', () => {
  const item = manifestItemFromJob('run', job(), 'completed');
  const manifest = createManifest(
    'run',
    'normal',
    'Archive',
    new Date(0).toISOString(),
    defaultDownloadSettings,
    {
      downloaded: 1,
      resumed: 0,
      verified: 1,
      skipped: 0,
      failed: 0,
      changedRemotely: 0,
      conflicts: 0,
      insufficientDiskSpace: 0,
    },
    [item],
    12,
    1,
    0,
  );

  assert.equal(manifest.source, 'OneDrive');
  assert.equal(manifest.targetFolderName, 'Archive');
  assert.equal(manifest.summary.totalFiles, 1);
  assert.equal(manifest.summary.totalBytes, 12);
  assert.equal(manifest.items[0].itemId, 'item');
});
