export type DownloadJobStatus =
  | 'pending'
  | 'scanning'
  | 'queued'
  | 'downloading'
  | 'verifying'
  | 'completed'
  | 'skipped'
  | 'paused'
  | 'retrying'
  | 'throttled'
  | 'failed'
  | 'stale_remote_changed'
  | 'insufficient_space'
  | 'cancelled';

export type ConflictStrategy = 'skip_existing' | 'overwrite' | 'rename_new' | 'keep_newest' | 'keep_both';

export interface DownloadSettings {
  conflictStrategy: ConflictStrategy;
  maxGlobalConcurrentDownloads: number;
  maxChunksPerFile: number;
  chunkSize: number;
  retryCount: number;
  speedLimitBytesPerSecond: number;
  preserveTimestamps: boolean;
  useDeltaSync: boolean;
  skipHiddenSystemFiles: boolean;
  includeExtensions: string[];
  excludeExtensions: string[];
  maximumFileSize: number;
  smallFilesFirst: boolean;
  verifyAfterDownload: boolean;
  autoResumeOnStart: boolean;
}

export interface RemoteHashes {
  quickXorHash?: string;
  sha1Hash?: string;
  crc32Hash?: string;
}

export interface RemoteItemMetadata {
  driveId: string;
  itemId: string;
  name: string;
  remotePath: string;
  size: number;
  eTag?: string;
  cTag?: string;
  lastModifiedDateTime?: string;
  hashes: RemoteHashes;
  downloadUrl?: string;
}

export interface DownloadChunk {
  index: number;
  start: number;
  end: number;
  status: 'pending' | 'downloading' | 'verified' | 'failed';
  attempts: number;
  updatedAt: string;
  error?: string;
}

export interface DownloadJob {
  id: string;
  driveId: string;
  itemId: string;
  name: string;
  remotePath: string;
  localPath: string;
  partialPath: string;
  size: number;
  eTag?: string;
  cTag?: string;
  lastModifiedDateTime?: string;
  hashes: RemoteHashes;
  status: DownloadJobStatus;
  priority: number;
  retryCount: number;
  errorMessage?: string;
  technicalError?: string;
  httpStatus?: number;
  createdAt: string;
  updatedAt: string;
  downloadedBytes: number;
  chunks: DownloadChunk[];
  summary?: string;
}

export interface DownloadSummary {
  downloaded: number;
  resumed: number;
  verified: number;
  skipped: number;
  failed: number;
  changedRemotely: number;
  conflicts: number;
  insufficientDiskSpace: number;
}

export interface DownloadProgressSnapshot {
  totalBytes: number;
  downloadedBytes: number;
  completedFiles: number;
  failedFiles: number;
  queuedFiles: number;
  stagePercent: number;
  currentFile?: string;
  currentFolder?: string;
  speedBytesPerSecond: number;
  etaSeconds?: number;
  status: DownloadJobStatus | 'idle';
  throttledUntil?: number;
  summary: DownloadSummary;
}

export type ReporterEvent =
  | { type: 'log'; stage: string; message: string; job?: DownloadJob }
  | { type: 'progress'; snapshot: DownloadProgressSnapshot }
  | { type: 'job'; job: DownloadJob };

export type Reporter = (event: ReporterEvent) => void;

export const defaultDownloadSettings: DownloadSettings = {
  conflictStrategy: 'skip_existing',
  maxGlobalConcurrentDownloads: 4,
  maxChunksPerFile: 2,
  chunkSize: 16 * 1024 * 1024,
  retryCount: 5,
  speedLimitBytesPerSecond: 0,
  preserveTimestamps: true,
  useDeltaSync: false,
  skipHiddenSystemFiles: true,
  includeExtensions: [],
  excludeExtensions: [],
  maximumFileSize: 0,
  smallFilesFirst: true,
  verifyAfterDownload: true,
  autoResumeOnStart: true,
};
