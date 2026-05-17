import assert from 'node:assert/strict';
import fs from 'node:fs/promises';
import os from 'node:os';
import path from 'node:path';
import test from 'node:test';
import { ServerStateStore, type PersistedServerJob } from '../src/server/serverStateStore.ts';

const snapshot = {
  totalBytes: 0,
  downloadedBytes: 0,
  completedFiles: 0,
  failedFiles: 0,
  queuedFiles: 0,
  stagePercent: 0,
  speedBytesPerSecond: 0,
  status: 'queued' as const,
  summary: {
    downloaded: 0,
    resumed: 0,
    verified: 0,
    skipped: 0,
    failed: 0,
    changedRemotely: 0,
    conflicts: 0,
    insufficientDiskSpace: 0,
  },
};

async function withStore(run: (store: ServerStateStore) => void) {
  const directory = await fs.mkdtemp(path.join(os.tmpdir(), 'onedrive-archiver-store-'));
  try {
    run(new ServerStateStore(directory));
  } finally {
    await fs.rm(directory, { recursive: true, force: true });
  }
}

test('persists server jobs and filters by user', async () => {
  await withStore(store => {
    const baseJob: PersistedServerJob = {
      id: 'job-1',
      userEmail: 'owner@example.com',
      mode: 'start',
      status: 'queued',
      createdAt: '2026-01-01T00:00:00.000Z',
      updatedAt: '2026-01-01T00:00:00.000Z',
      targetRoot: '/downloads/users/owner@example.com',
      selections: [{ id: 'root', name: 'OneDrive', type: 'folder' }],
      settings: { conflictStrategy: 'skip_existing' },
      log: ['queued'],
      snapshot,
    };

    store.upsertJob(baseJob);
    store.upsertJob({ ...baseJob, id: 'job-2', userEmail: 'other@example.com' });

    assert.equal(store.listJobs('owner@example.com').length, 1);
    assert.equal(store.getJob('owner@example.com', 'job-1')?.id, 'job-1');
    assert.equal(store.getJob('owner@example.com', 'job-2'), undefined);
  });
});

test('marks active jobs interrupted and persists delta tokens', async () => {
  await withStore(store => {
    const job: PersistedServerJob = {
      id: 'job-1',
      userEmail: 'owner@example.com',
      mode: 'incremental',
      status: 'downloading',
      createdAt: '2026-01-01T00:00:00.000Z',
      updatedAt: '2026-01-01T00:00:00.000Z',
      targetRoot: '/downloads/users/owner@example.com',
      selections: [],
      settings: {},
      log: [],
      snapshot: { ...snapshot, status: 'downloading' },
    };

    store.upsertJob(job);
    store.markActiveJobsInterrupted();
    assert.equal(store.listResumableJobs('owner@example.com')[0]?.status, 'interrupted');

    store.saveDeltaToken('owner@example.com', 'delta-link');
    assert.equal(store.getDeltaToken('owner@example.com'), 'delta-link');
  });
});

test('persists server archive records for repair', async () => {
  await withStore(store => {
    store.upsertArchiveRecord({
      userEmail: 'owner@example.com',
      localPath: '/downloads/users/owner_example.com/Documents/report.txt',
      status: 'completed',
      verificationMessage: 'SHA-1 verification passed.',
      item: {
        driveId: 'me',
        itemId: 'item-1',
        name: 'report.txt',
        remotePath: 'Documents/report.txt',
        size: 4,
        hashes: { sha1Hash: 'A9993E364706816ABA3E25717850C26C9CD0D89D' },
      },
    });

    const records = store.listArchiveRecords('owner@example.com');
    assert.equal(records.length, 1);
    assert.equal(records[0].item.remotePath, 'Documents/report.txt');
    assert.equal(records[0].verificationMessage, 'SHA-1 verification passed.');
    assert.equal(store.listArchiveRecords('other@example.com').length, 0);
  });
});
