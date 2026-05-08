import type {
  DownloadJob,
  DownloadRunMode,
  DownloadSettings,
  DownloadSummary,
  ManifestItemStatus,
  RunManifest,
  RunManifestItem,
} from './types';

export function createRunId() {
  return new Date().toISOString().replace(/[:.]/g, '-');
}

export function manifestItemFromJob(
  runId: string,
  job: DownloadJob,
  status: ManifestItemStatus,
  details?: {
    verification?: string;
    error?: string;
    finalName?: string;
  },
): RunManifestItem {
  return {
    runId,
    itemId: job.itemId,
    driveId: job.driveId,
    name: job.name,
    remotePath: job.remotePath,
    localPath: job.localPath,
    size: job.size,
    eTag: job.eTag,
    cTag: job.cTag,
    lastModifiedDateTime: job.lastModifiedDateTime,
    sha1Hash: job.hashes.sha1Hash,
    quickXorHash: job.hashes.quickXorHash,
    status,
    verification: details?.verification || job.summary,
    error: details?.error || job.errorMessage,
    finalName: details?.finalName,
    updatedAt: new Date().toISOString(),
  };
}

export function createManifest(
  runId: string,
  mode: DownloadRunMode,
  targetFolderName: string,
  startedAt: string,
  settings: DownloadSettings,
  summary: DownloadSummary,
  items: RunManifestItem[],
  totalBytes: number,
  completedFiles: number,
  failedFiles: number,
): RunManifest {
  return {
    runId,
    mode,
    source: 'OneDrive',
    targetFolderName,
    startedAt,
    finishedAt: new Date().toISOString(),
    settings,
    summary: {
      ...summary,
      totalFiles: items.length,
      totalBytes,
      completedFiles,
      failedFiles,
    },
    items,
  };
}
