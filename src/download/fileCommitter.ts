import { getUniqueFileName, normalizePathSegment } from './pathTools';
import type { ConflictStrategy, DownloadJob } from './types';

export class FileCommitter {
  async getOrCreateDirectory(root: FileSystemDirectoryHandle, parts: string[]) {
    let directory = root;
    for (const part of parts) {
      directory = await directory.getDirectoryHandle(normalizePathSegment(part), { create: true });
    }
    return directory;
  }

  async getPartialHandle(directory: FileSystemDirectoryHandle, name: string) {
    return directory.getFileHandle(`${normalizePathSegment(name)}.partial`, { create: true });
  }

  async getExistingFile(directory: FileSystemDirectoryHandle, name: string) {
    try {
      const handle = await directory.getFileHandle(normalizePathSegment(name));
      return handle.getFile();
    } catch {
      return undefined;
    }
  }

  async getExistingFileByPath(root: FileSystemDirectoryHandle, localPath: string) {
    const parts = localPath.split('/').filter(Boolean);
    const name = parts.pop();
    if (!name) return undefined;
    let directory = root;
    for (const part of parts) {
      try {
        directory = await directory.getDirectoryHandle(normalizePathSegment(part));
      } catch {
        return undefined;
      }
    }
    return this.getExistingFile(directory, name);
  }

  async getPartialFile(directory: FileSystemDirectoryHandle, name: string) {
    try {
      const handle = await directory.getFileHandle(`${normalizePathSegment(name)}.partial`);
      return handle.getFile();
    } catch {
      return undefined;
    }
  }

  async resolveFinalName(
    directory: FileSystemDirectoryHandle,
    job: DownloadJob,
    strategy: ConflictStrategy,
  ) {
    return getUniqueFileName(directory, job.name, strategy);
  }

  async commit(
    directory: FileSystemDirectoryHandle,
    partialFile: File,
    job: DownloadJob,
    strategy: ConflictStrategy,
  ) {
    const finalName = job.name;
    const keepNewestDecision = strategy === 'keep_newest'
      ? await this.resolveKeepNewest(directory, job)
      : undefined;
    if (keepNewestDecision?.skip) {
      return { committed: false, finalName: normalizePathSegment(finalName), reason: keepNewestDecision.reason };
    }

    const resolvedName = keepNewestDecision?.finalName || await getUniqueFileName(directory, finalName, strategy);
    if (!resolvedName) {
      return { committed: false, finalName: normalizePathSegment(finalName), reason: 'Existing file kept by conflict strategy.' };
    }

    const tempName = `${resolvedName}.commit-${Date.now()}-${Math.random().toString(36).slice(2)}`;
    const tempHandle = await directory.getFileHandle(tempName, { create: true });
    try {
      const tempWritable = await tempHandle.createWritable();
      await partialFile.stream().pipeTo(tempWritable);
      const tempFile = await tempHandle.getFile();
      if (tempFile.size !== partialFile.size) {
        throw new Error(`Commit copy failed: expected ${partialFile.size} bytes, wrote ${tempFile.size} bytes.`);
      }

      const finalHandle = await directory.getFileHandle(resolvedName, { create: true });
      try {
        const writable = await finalHandle.createWritable();
        await tempFile.stream().pipeTo(writable);
      } catch (error) {
        try {
          await directory.removeEntry(resolvedName);
        } catch {
          // If cleanup fails, keep the original error so the job stays failed.
        }
        throw error;
      }
    } finally {
      try {
        await directory.removeEntry(tempName);
      } catch {
        // Best-effort cleanup for failed/interrupted commits.
      }
    }
    return { committed: true, finalName: resolvedName };
  }

  private async resolveKeepNewest(directory: FileSystemDirectoryHandle, job: DownloadJob) {
    const normalized = normalizePathSegment(job.name);
    let existing: File | undefined;
    try {
      existing = await (await directory.getFileHandle(normalized)).getFile();
    } catch {
      return { finalName: normalized };
    }

    const remoteModified = job.lastModifiedDateTime ? Date.parse(job.lastModifiedDateTime) : NaN;
    const localModified = existing.lastModified;
    if (!Number.isFinite(remoteModified) || !Number.isFinite(localModified)) {
      return {
        skip: true,
        reason: 'Existing file kept by keep_newest because a reliable local or remote timestamp was unavailable.',
      };
    }

    if (localModified >= remoteModified) {
      return { skip: true, reason: 'Existing file kept by keep_newest because it is newer or same age.' };
    }

    return { finalName: normalized };
  }

  async clearPartial(directory: FileSystemDirectoryHandle, name: string) {
    try {
      await directory.removeEntry(`${normalizePathSegment(name)}.partial`);
    } catch {
      // Best-effort cleanup; stale partial tracking remains in IndexedDB.
    }
  }

  async clearPartialByPath(root: FileSystemDirectoryHandle, localPath: string) {
    const parts = localPath.split('/').filter(Boolean);
    const name = parts.pop();
    if (!name) return;
    let directory = root;
    for (const part of parts) {
      try {
        directory = await directory.getDirectoryHandle(normalizePathSegment(part));
      } catch {
        return;
      }
    }
    await this.clearPartial(directory, name);
  }
}
