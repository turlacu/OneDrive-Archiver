import assert from 'node:assert/strict';
import test from 'node:test';
import { FileCommitter } from '../src/download/fileCommitter.ts';
import type { DownloadJob } from '../src/download/types.ts';

function memoryWritable(closeFile: (file: File) => void) {
  const chunks: Uint8Array[] = [];
  return new WritableStream({
    async write(data: BufferSource | Blob | string) {
    if (data instanceof Blob) {
      chunks.push(new Uint8Array(await data.arrayBuffer()));
    } else if (typeof data === 'string') {
      chunks.push(new TextEncoder().encode(data));
    } else if (ArrayBuffer.isView(data)) {
      chunks.push(new Uint8Array(data.buffer.slice(data.byteOffset, data.byteOffset + data.byteLength)));
    } else if (data instanceof ArrayBuffer) {
      chunks.push(new Uint8Array(data));
    }
    },
    close() {
      closeFile(new File(chunks, 'written.bin'));
    },
  });
}

class MemoryFileHandle {
  constructor(private file: File) {}

  async getFile() {
    return this.file;
  }

  async createWritable() {
    return memoryWritable(file => {
      this.file = file;
    }) as unknown as FileSystemWritableFileStream;
  }
}

class MemoryDirectory {
  files = new Map<string, MemoryFileHandle>();

  constructor(initial: Record<string, File> = {}) {
    for (const [name, file] of Object.entries(initial)) {
      this.files.set(name, new MemoryFileHandle(file));
    }
  }

  async getFileHandle(name: string, options?: { create?: boolean }) {
    const existing = this.files.get(name);
    if (existing) return existing;
    if (options?.create) {
      const handle = new MemoryFileHandle(new File([], name));
      this.files.set(name, handle);
      return handle;
    }
    throw new Error('not found');
  }

  async removeEntry(name: string) {
    this.files.delete(name);
  }
}

function job(patch: Partial<DownloadJob> = {}): DownloadJob {
  return {
    id: 'job',
    driveId: 'drive',
    itemId: 'item',
    name: 'report.txt',
    remotePath: 'report.txt',
    localPath: 'report.txt',
    partialPath: 'report.txt.partial',
    size: 4,
    hashes: {},
    status: 'queued',
    priority: 0,
    retryCount: 0,
    createdAt: new Date().toISOString(),
    updatedAt: new Date().toISOString(),
    downloadedBytes: 0,
    chunks: [],
    ...patch,
  };
}

test('keep_newest skips when remote timestamp is unavailable', async () => {
  const directory = new MemoryDirectory({
    'report.txt': new File(['old'], 'report.txt', { lastModified: 10 }),
  }) as unknown as FileSystemDirectoryHandle;
  const committer = new FileCommitter();
  const result = await committer.commit(directory, new File(['next'], 'partial'), job({ lastModifiedDateTime: undefined }), 'keep_newest');

  assert.equal(result.committed, false);
  assert.match(result.reason || '', /timestamp was unavailable/);
});

test('keep_newest overwrites when remote timestamp is newer', async () => {
  const directory = new MemoryDirectory({
    'report.txt': new File(['old'], 'report.txt', { lastModified: 10 }),
  }) as unknown as FileSystemDirectoryHandle;
  const committer = new FileCommitter();
  const result = await committer.commit(directory, new File(['next'], 'partial'), job({ lastModifiedDateTime: new Date(20).toISOString() }), 'keep_newest');

  assert.equal(result.committed, true);
});
