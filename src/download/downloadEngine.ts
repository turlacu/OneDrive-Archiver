import { DownloadReporter } from './downloadReporter';
import { FileCommitter } from './fileCommitter';
import { IntegrityVerifier } from './integrityVerifier';
import { DownloadJobStore } from './jobStore';
import { OneDriveClient, type GraphDriveItem } from './oneDriveClient';
import { createManifest, createRunId, manifestItemFromJob } from './runManifest';
import { extensionOf, isTransientOfficeLockFile, normalizeRelativePath, shouldIncludeFile } from './pathTools';
import { backoffDelayMs, HttpDownloadError, isTransientStatus, parseRetryAfter, sleep } from './retry';
import type {
  DownloadChunk,
  DownloadJob,
  DownloadRunMode,
  DownloadSettings,
  RunManifestItem,
  RemoteItemMetadata,
  Reporter,
} from './types';
import { defaultDownloadSettings } from './types';

export interface SourceSelection {
  id: string;
  name: string;
  type: 'file' | 'folder';
}

interface EngineCounters {
  totalBytes: number;
  downloadedBytes: number;
  completedFiles: number;
  failedFiles: number;
  queuedFiles: number;
  speedBytesPerSecond: number;
  lastSpeedBytes: number;
  lastSpeedAt: number;
}

export class DownloadEngine {
  private readonly settings: DownloadSettings;
  private readonly store = new DownloadJobStore();
  private readonly verifier = new IntegrityVerifier();
  private readonly committer = new FileCommitter();
  private readonly reporter: DownloadReporter;
  private abortController = new AbortController();
  private paused = false;
  private counters: EngineCounters = {
    totalBytes: 0,
    downloadedBytes: 0,
    completedFiles: 0,
    failedFiles: 0,
    queuedFiles: 0,
    speedBytesPerSecond: 0,
    lastSpeedBytes: 0,
    lastSpeedAt: 0,
  };
  private scanProcessed = 0;
  private scanPending = 0;
  private failedJobs = new Map<string, DownloadJob>();
  private manifestItems = new Map<string, RunManifestItem>();
  private runId = createRunId();
  private mode: DownloadRunMode = 'normal';
  private startedAt = new Date().toISOString();
  private throttleEvents = 0;

  constructor(
    private readonly oneDrive: OneDriveClient,
    reporter: Reporter,
    settings?: Partial<DownloadSettings>,
  ) {
    this.settings = { ...defaultDownloadSettings, ...settings };
    this.reporter = new DownloadReporter(reporter);
  }

  pause() {
    this.paused = true;
    this.reporter.log('queue', 'Downloads paused.');
  }

  resume() {
    this.paused = false;
    this.reporter.log('queue', 'Downloads resumed.');
  }

  cancel() {
    this.abortController.abort();
    this.reporter.log('queue', 'Downloads cancelled.');
  }

  async start(
    selections: SourceSelection[],
    rootDirectory: FileSystemDirectoryHandle,
  ) {
    this.beginRun('normal', selections.length);

    this.reporter.log('scanner', `Scanning ${selections.length} selected item${selections.length === 1 ? '' : 's'}...`);
    this.emitProgress('scanning');
    const jobs = await this.scanSelections(selections, []);
    const runnableJobs = this.settings.smallFilesFirst
      ? jobs.sort((a, b) => a.size - b.size)
      : jobs;

    this.counters.totalBytes = runnableJobs.reduce((total, job) => total + job.size, 0);
    this.counters.downloadedBytes = runnableJobs.reduce((total, job) => total + job.downloadedBytes, 0);
    this.counters.lastSpeedBytes = this.counters.downloadedBytes;
    this.counters.lastSpeedAt = performance.now();
    this.counters.queuedFiles = runnableJobs.length;
    this.emitProgress('queued');

    await this.runQueue(runnableJobs, rootDirectory);

    if (!this.abortController.signal.aborted && this.failedJobs.size > 0) {
      await this.retryFailedPass(rootDirectory);
    }

    this.emitProgress(this.counters.failedFiles > 0 ? 'failed' : 'completed');
    this.reporter.log(
      'reporter',
      `Summary: ${this.counters.completedFiles} completed, ${this.counters.failedFiles} failed, ${this.reporter.getSummary().skipped} skipped.`,
    );
    await this.writeManifest(rootDirectory);

    if (!this.abortController.signal.aborted && this.counters.failedFiles === 0) {
      this.reporter.log('queue', 'Completed job records were retained for archive repair.');
    }
  }

  async dryRun(
    selections: SourceSelection[],
    rootDirectory: FileSystemDirectoryHandle,
  ) {
    this.beginRun('dry_run', selections.length);

    this.reporter.log('scanner', `Preflighting ${selections.length} selected item${selections.length === 1 ? '' : 's'}...`);
    this.emitProgress('scanning');
    const jobs = await this.scanSelections(selections, []);
    this.counters.totalBytes = jobs.reduce((total, job) => total + job.size, 0);
    this.counters.queuedFiles = jobs.length;

    for (const job of jobs) {
      const existing = await this.committer.getExistingFileByPath(rootDirectory, job.localPath);
      const verification = existing
        ? `Dry run: local file exists; conflict strategy is ${this.settings.conflictStrategy}.`
        : 'Dry run: file would be downloaded.';
      this.recordManifestItem(manifestItemFromJob(this.runId, job, 'unverified', { verification }));
      this.reporter.job({ ...job, status: 'queued', summary: verification });
    }

    this.emitProgress('completed');
    this.reporter.log('reporter', `Dry run summary: ${jobs.length} file${jobs.length === 1 ? '' : 's'}, ${this.counters.totalBytes} bytes.`);
  }

  async startIncremental(rootDirectory: FileSystemDirectoryHandle) {
    this.beginRun('incremental', 1);
    this.reporter.log('scanner', 'Checking OneDrive delta changes...');
    this.emitProgress('scanning');

    try {
      const tokenKey = 'me';
      let token: string | undefined;
      try {
        token = await this.store.getDeltaToken(tokenKey);
      } catch (error) {
        const message = error instanceof Error ? error.message : String(error);
        this.reporter.log('scanner', `Could not read saved delta token; starting a fresh delta scan: ${message}`);
      }
      this.reporter.log('scanner', token ? 'Using saved OneDrive delta token.' : 'No saved delta token found; starting delta baseline scan.');
      const delta = await this.oneDrive.delta(tokenKey, token);
      if (!token) {
        if (delta.deltaToken) await this.store.saveDeltaToken(tokenKey, delta.deltaToken);
        this.reporter.log('scanner', 'Baseline delta token saved. Incremental archive will download changes from the next run.');
        this.counters.queuedFiles = 0;
        this.emitProgress('completed');
        return;
      }
      this.reporter.log('scanner', `OneDrive delta returned ${delta.items.length} changed file${delta.items.length === 1 ? '' : 's'}.`);
      const jobs: DownloadJob[] = [];
      for (const metadata of delta.items) {
        const parentParts = this.parentPartsFromRemotePath(metadata);
        const job = await this.createOrResumeJob(metadata, parentParts);
        if (job) jobs.push(job);
      }

      this.counters.totalBytes = jobs.reduce((total, job) => total + job.size, 0);
      this.counters.downloadedBytes = jobs.reduce((total, job) => total + job.downloadedBytes, 0);
      this.counters.queuedFiles = jobs.length;
      this.counters.lastSpeedBytes = this.counters.downloadedBytes;
      this.counters.lastSpeedAt = performance.now();
      this.emitProgress('queued');
      this.reporter.log('scanner', `Incremental scan found ${jobs.length} changed file${jobs.length === 1 ? '' : 's'} to archive.`);

      await this.runQueue(jobs, rootDirectory);

      if (!this.abortController.signal.aborted && this.failedJobs.size > 0) {
        await this.retryFailedPass(rootDirectory);
      }

      this.emitProgress(this.counters.failedFiles > 0 ? 'failed' : 'completed');
      this.reporter.log('reporter', `Incremental summary: ${this.counters.completedFiles} completed, ${this.counters.failedFiles} failed.`);
      await this.writeManifest(rootDirectory);

      if (!this.abortController.signal.aborted && this.counters.failedFiles === 0) {
        if (delta.deltaToken) {
          try {
            await this.store.saveDeltaToken(tokenKey, delta.deltaToken);
            this.reporter.log('scanner', 'Saved OneDrive delta token after successful incremental archive.');
          } catch (error) {
            const message = error instanceof Error ? error.message : String(error);
            this.reporter.log('scanner', `Incremental archive completed, but the delta token could not be saved: ${message}`);
          }
        }
        this.reporter.log('queue', 'Completed job records were retained for archive repair.');
      }
    } catch (error) {
      const message = error instanceof Error ? error.message : String(error);
      throw new Error(`Incremental stage failed: ${message}`);
    }
  }

  async repairFromLatestManifest(rootDirectory: FileSystemDirectoryHandle) {
    this.beginRun('repair', 1);
    const latest = await this.store.getLatestManifest();
    const storedCompletedJobs = await this.store.getJobsByStatus(['completed', 'skipped']);
    if (!latest && storedCompletedJobs.length === 0) {
      this.reporter.log('verifier', 'No saved archive manifest or completed job records are available for repair.');
      this.emitProgress('idle');
      return;
    }
    const manifestJobs = latest?.items
      .filter(item => ['completed', 'verified', 'skipped'].includes(item.status))
      .map(item => this.jobFromManifestItem(item)) || [];
    const candidatesById = new Map<string, DownloadJob>();
    for (const job of [...manifestJobs, ...storedCompletedJobs]) {
      candidatesById.set(job.id, job);
    }
    const candidates = Array.from(candidatesById.values());
    this.counters.totalBytes = candidates.reduce((total, item) => total + item.size, 0);
    this.counters.queuedFiles = candidates.length;
    this.reporter.log('verifier', `Checking ${candidates.length} saved archive file${candidates.length === 1 ? '' : 's'}...`);
    this.emitProgress('verifying');

    const repairJobs: DownloadJob[] = [];
    for (const job of candidates) {
      const file = await this.committer.getExistingFileByPath(rootDirectory, job.localPath);
      if (!file) {
        this.recordManifestItem(manifestItemFromJob(this.runId, job, 'missing_local', { error: 'Local file is missing.' }));
        repairJobs.push(await this.store.upsertJob({ ...job, status: 'queued', downloadedBytes: 0, chunks: this.createChunks(job.size) }));
        continue;
      }

      const verification = await this.verifier.verify(file, job);
      if (verification.ok) {
        this.counters.completedFiles += 1;
        this.recordManifestItem(manifestItemFromJob(this.runId, job, 'verified', { verification: verification.message }));
        this.reporter.job({ ...job, status: 'completed', downloadedBytes: job.size, summary: verification.message });
      } else {
        this.recordManifestItem(manifestItemFromJob(this.runId, job, 'unverified', { error: verification.message }));
        repairJobs.push(await this.store.upsertJob({ ...job, status: 'queued', downloadedBytes: 0, chunks: this.createChunks(job.size) }));
      }
    }

    this.counters.queuedFiles = repairJobs.length;
    if (repairJobs.length > 0) {
      this.reporter.log('queue', `Repairing ${repairJobs.length} missing or unverified file${repairJobs.length === 1 ? '' : 's'}...`);
      await this.runQueue(repairJobs, rootDirectory);
    }

    this.emitProgress(this.counters.failedFiles > 0 ? 'failed' : 'completed');
    this.reporter.log('reporter', `Repair summary: ${this.counters.completedFiles} verified/completed, ${this.counters.failedFiles} failed.`);
    await this.writeManifest(rootDirectory);
  }

  async repairSelections(
    selections: SourceSelection[],
    rootDirectory: FileSystemDirectoryHandle,
  ) {
    this.beginRun('repair', selections.length);
    this.reporter.log('scanner', `Scanning ${selections.length} selected item${selections.length === 1 ? '' : 's'} for repair...`);
    this.emitProgress('scanning');
    const jobs = await this.scanSelections(selections, []);
    const repairJobs: DownloadJob[] = [];
    for (const job of jobs) {
      const existing = await this.committer.getExistingFileByPath(rootDirectory, job.localPath);
      if (!existing) {
        repairJobs.push({ ...job, status: 'queued', downloadedBytes: 0, chunks: this.createChunks(job.size) });
        this.recordManifestItem(manifestItemFromJob(this.runId, job, 'missing_local', { error: 'Local file is missing.' }));
        continue;
      }
      const verification = await this.verifier.verify(existing, job);
      if (verification.ok) {
        this.counters.completedFiles += 1;
        this.recordManifestItem(manifestItemFromJob(this.runId, job, 'verified', { verification: verification.message }));
      } else {
        repairJobs.push({ ...job, status: 'queued', downloadedBytes: 0, chunks: this.createChunks(job.size) });
        this.recordManifestItem(manifestItemFromJob(this.runId, job, 'unverified', { error: verification.message }));
      }
    }

    this.counters.totalBytes = repairJobs.reduce((total, job) => total + job.size, 0);
    this.counters.queuedFiles = repairJobs.length;
    this.emitProgress('queued');
    this.reporter.log('queue', `Repair selected found ${repairJobs.length} missing or unverified file${repairJobs.length === 1 ? '' : 's'}.`);
    await this.runQueue(repairJobs, rootDirectory);
    this.emitProgress(this.counters.failedFiles > 0 ? 'failed' : 'completed');
    this.reporter.log('reporter', `Repair selected summary: ${this.counters.completedFiles} verified/completed, ${this.counters.failedFiles} failed.`);
    await this.writeManifest(rootDirectory);
  }

  async retryJobs(
    jobIds: string[],
    rootDirectory: FileSystemDirectoryHandle,
  ) {
    this.beginRun('normal', jobIds.length);

    const jobs = (await Promise.all(jobIds.map(id => this.store.getJob(id))))
      .filter((job): job is DownloadJob => Boolean(job));
    const retryableJobs = jobs
      .filter(job => job.status === 'failed')
      .map(job => this.prepareRetryJob(job));

    if (retryableJobs.length === 0) {
      this.reporter.log('queue', 'No failed files are available to retry.');
      this.emitProgress('idle');
      return;
    }

    this.counters.totalBytes = retryableJobs.reduce((total, job) => total + job.size, 0);
    this.counters.downloadedBytes = retryableJobs.reduce((total, job) => total + job.downloadedBytes, 0);
    this.counters.queuedFiles = retryableJobs.length;
    this.counters.lastSpeedBytes = this.counters.downloadedBytes;
    this.counters.lastSpeedAt = performance.now();

    this.reporter.log('queue', `Retrying ${retryableJobs.length} failed file${retryableJobs.length === 1 ? '' : 's'}...`);
    this.emitProgress('retrying');
    await this.runQueue(retryableJobs, rootDirectory);

    this.emitProgress(this.counters.failedFiles > 0 ? 'failed' : 'completed');
    this.reporter.log(
      'reporter',
      `Retry summary: ${this.counters.completedFiles} completed, ${this.counters.failedFiles} failed.`,
    );
    await this.writeManifest(rootDirectory);

    if (!this.abortController.signal.aborted && this.counters.failedFiles === 0) {
      this.reporter.log('queue', 'Completed job records were retained for archive repair.');
    }
  }

  private async runWorker(queue: DownloadJob[], rootDirectory: FileSystemDirectoryHandle) {
    while (queue.length > 0 && !this.abortController.signal.aborted) {
      while (this.paused && !this.abortController.signal.aborted) {
        await sleep(500, this.abortController.signal);
      }
      const job = queue.shift();
      if (!job) return;
      await this.runJob(job, rootDirectory);
    }
  }

  private async runQueue(jobs: DownloadJob[], rootDirectory: FileSystemDirectoryHandle) {
    const queue = [...(this.settings.smallFilesFirst ? [...jobs].sort((a, b) => a.size - b.size) : jobs)];
    const baseConcurrency = Math.max(1, this.settings.maxGlobalConcurrentDownloads);
    const adaptiveConcurrency = Math.max(1, baseConcurrency - Math.min(this.throttleEvents, baseConcurrency - 1));
    if (adaptiveConcurrency < baseConcurrency) {
      this.reporter.log('queue', `Reduced concurrency to ${adaptiveConcurrency} after OneDrive throttling.`);
    }
    const workers = Array.from(
      { length: adaptiveConcurrency },
      () => this.runWorker(queue, rootDirectory),
    );
    await Promise.all(workers);
  }

  private async retryFailedPass(rootDirectory: FileSystemDirectoryHandle) {
    const retryJobs = Array.from(this.failedJobs.values())
      .filter(job => job.status === 'failed')
      .map(job => this.prepareRetryJob(job));

    this.failedJobs.clear();
    this.counters.failedFiles = 0;

    if (retryJobs.length === 0) return;

    this.reporter.log('queue', `Retrying ${retryJobs.length} failed file${retryJobs.length === 1 ? '' : 's'} once before finishing...`);
    this.emitProgress('retrying');

    await this.runQueue(retryJobs, rootDirectory);
  }

  private prepareRetryJob(job: DownloadJob): DownloadJob {
    const chunks = job.chunks.map(chunk => (
      chunk.status === 'verified'
        ? chunk
        : { ...chunk, status: 'pending' as const, error: undefined, updatedAt: new Date().toISOString() }
    ));

    return {
      ...job,
      status: 'retrying',
      errorMessage: undefined,
      technicalError: undefined,
      httpStatus: undefined,
      downloadedBytes: this.completedBytes(chunks),
      chunks,
    };
  }

  private async scanSelections(selections: SourceSelection[], parentParts: string[]) {
    const jobs: DownloadJob[] = [];
    for (const selection of selections) {
      if (this.abortController.signal.aborted) break;
      if (selection.type === 'folder') {
        await this.scanFolder(selection.id, [...parentParts, selection.name], jobs);
      } else {
        const metadata = await this.oneDrive.getItem(selection.id);
        const job = await this.createOrResumeJob(metadata, parentParts);
        if (job) jobs.push(job);
      }
      this.scanProcessed += 1;
      this.scanPending = Math.max(this.scanPending - 1, 0);
      this.emitProgress('scanning');
    }
    return jobs;
  }

  private async scanFolder(folderId: string, parentParts: string[], jobs: DownloadJob[]) {
    this.reporter.log('scanner', `Scanning folder ${parentParts.join('/')}`);
    const children = await this.oneDrive.listChildren(folderId);
    for (const child of children) {
      if (this.abortController.signal.aborted) return;
      if (child.folder || child.package) {
        this.scanPending += 1;
        this.emitProgress('scanning');
        await this.scanFolder(child.id, [...parentParts, child.name], jobs);
        this.scanProcessed += 1;
        this.scanPending = Math.max(this.scanPending - 1, 0);
        this.emitProgress('scanning');
      } else {
        const metadata = this.oneDrive.normalizeItem(child as GraphDriveItem);
        const job = await this.createOrResumeJob(metadata, parentParts);
        if (job) jobs.push(job);
        this.scanProcessed += 1;
        this.emitProgress('scanning');
      }
    }
  }

  private async createOrResumeJob(metadata: RemoteItemMetadata, parentParts: string[]) {
    if (isTransientOfficeLockFile(metadata.name)) {
      this.reporter.increment('skipped');
      this.reporter.log('scanner', `Skipped transient Office lock file: ${metadata.name}`);
      return undefined;
    }

    if (!shouldIncludeFile(metadata.name, this.settings.includeExtensions, this.settings.excludeExtensions)) {
      this.reporter.increment('skipped');
      this.reporter.log('scanner', `Skipped by extension rule: ${metadata.name}`);
      return undefined;
    }

    if (this.settings.maximumFileSize > 0 && metadata.size > this.settings.maximumFileSize) {
      this.reporter.increment('skipped');
      this.reporter.log('scanner', `Skipped because file exceeds maximum size: ${metadata.name}`);
      return undefined;
    }

    const localPath = normalizeRelativePath([...parentParts, metadata.name]);
    const id = `${metadata.driveId}:${metadata.itemId}:${localPath}`;
    const existing = await this.store.getJob(id);
    const now = new Date().toISOString();
    const canResume = existing
      && this.remoteMatches(existing, metadata)
      && ['queued', 'downloading', 'verifying', 'paused', 'retrying', 'throttled'].includes(existing.status);
    const chunks = canResume
      ? existing.chunks
      : this.createChunks(metadata.size);

    if (existing && !this.remoteMatches(existing, metadata)) {
      const stale = await this.store.updateJob(existing.id, {
        status: 'stale_remote_changed',
        errorMessage: 'Remote file changed since download started; restarting safely.',
      });
      if (stale) this.recordManifestItem(manifestItemFromJob(this.runId, stale, 'remote_changed', { error: stale.errorMessage }));
      this.reporter.increment('changedRemotely');
      this.reporter.log('scanner', `Remote file changed since download started; restarting safely: ${metadata.name}`);
    }

    const job: DownloadJob = {
      id,
      driveId: metadata.driveId,
      itemId: metadata.itemId,
      name: metadata.name,
      remotePath: metadata.remotePath,
      localPath,
      partialPath: `${localPath}.partial`,
      size: metadata.size,
      eTag: metadata.eTag,
      cTag: metadata.cTag,
      lastModifiedDateTime: metadata.lastModifiedDateTime,
      hashes: metadata.hashes,
      status: canResume ? existing.status : 'queued',
      priority: existing?.priority || 0,
      retryCount: canResume ? existing.retryCount : 0,
      createdAt: existing?.createdAt || now,
      updatedAt: now,
      downloadedBytes: canResume ? this.completedBytes(chunks) : 0,
      chunks,
    };

    if (job.downloadedBytes > 0 && job.downloadedBytes < job.size) {
      this.reporter.increment('resumed');
      this.reporter.log('queue', `Resume candidate found: ${job.name} (${job.downloadedBytes}/${job.size} bytes)`, job);
    }

    const saved = await this.store.upsertJob(job);
    this.reporter.job(saved);
    return saved;
  }

  private async runJob(job: DownloadJob, rootDirectory: FileSystemDirectoryHandle) {
    try {
      const directoryParts = job.localPath.split('/').slice(0, -1);
      const directory = await this.committer.getOrCreateDirectory(rootDirectory, directoryParts);
      job = await this.validateResumeState(job, directory);
      const existing = await this.committer.getExistingFile(directory, job.name);

      if (existing && this.settings.conflictStrategy === 'skip_existing') {
        const verified = await this.verifier.verify(existing, job);
        if (verified.ok) {
          await this.completeJob(job, 'skipped', job.size, verified.message, true);
          this.reporter.increment('skipped');
          this.reporter.log('queue', `Skipped existing verified file: ${job.name}`, job);
          return;
        }
        this.reporter.log('verifier', `Existing file failed verification; redownloading: ${job.name}`, job);
      }

      let current = await this.store.updateJob(job.id, { status: 'downloading', errorMessage: undefined });
      if (!current) return;
      this.reporter.job(current);
      this.reporter.log('downloader', `Downloading ${current.name}`, current);
      this.emitProgress('downloading', current);

      await this.downloadChunks(current, directory);

      current = await this.store.updateJob(job.id, { status: 'verifying' });
      if (!current) return;
      this.reporter.job(current);
      this.reporter.log('verifier', `Verifying ${current.name}`, current);
      this.emitProgress('verifying', current);

      const partialHandle = await this.committer.getPartialHandle(directory, current.name);
      const partialFile = await partialHandle.getFile();
      const verification = this.settings.verifyAfterDownload
        ? await this.verifier.verify(partialFile, current)
        : { ok: true, cryptographic: false, message: 'Verification disabled by settings.' };

      if (!verification.ok) {
        this.reporter.log('verifier', `Verification failed; redownloading: ${current.name}`, current);
        current = await this.store.updateJob(current.id, {
          chunks: this.createChunks(current.size),
          downloadedBytes: 0,
          status: 'retrying',
          errorMessage: verification.message,
        }) || current;
        await this.downloadChunks(current, directory);
      }

      const finalVerificationFile = await partialHandle.getFile();
      const finalVerification = this.settings.verifyAfterDownload
        ? await this.verifier.verify(finalVerificationFile, current)
        : verification;
      if (!finalVerification.ok) {
        throw new Error(finalVerification.message);
      }

      const commit = await this.committer.commit(directory, finalVerificationFile, current, this.settings.conflictStrategy);
      if (!commit.committed) {
        await this.completeJob(current, 'skipped', current.size, commit.reason, false, commit.finalName);
        this.reporter.increment('conflicts');
        this.reporter.log('committer', commit.reason || `Skipped existing file: ${current.name}`, current);
        return;
      }

      await this.committer.clearPartial(directory, current.name);
      await this.completeJob(current, 'completed', current.size, finalVerification.message, false, commit.finalName);
      this.reporter.increment('downloaded');
      this.reporter.increment('verified');
      if (this.settings.preserveTimestamps && current.lastModifiedDateTime) {
        this.reporter.log('committer', 'Browser File System Access API does not support setting downloaded file timestamps; remote modified time was retained in job metadata only.', current);
      }
      this.reporter.log('committer', `Committed verified file: ${commit.finalName}`, current);
    } catch (error) {
      const message = error instanceof Error ? error.message : String(error);
      const status = error instanceof HttpDownloadError ? error.status : undefined;
      const wasCancelled = this.abortController.signal.aborted;
      const failed = await this.store.updateJob(job.id, {
        status: wasCancelled ? 'cancelled' : 'failed',
        errorMessage: message,
        technicalError: error instanceof Error ? error.stack : String(error),
        httpStatus: status,
        retryCount: job.retryCount + 1,
      });
      if (!wasCancelled) {
        this.counters.failedFiles += 1;
        this.reporter.increment('failed');
      }
      if (failed) this.reporter.job(failed);
      if (failed && !wasCancelled) this.failedJobs.set(failed.id, failed);
      this.recordManifestItem(manifestItemFromJob(this.runId, failed || job, wasCancelled ? 'cancelled' : 'failed', { error: message }));
      this.reporter.log('downloader', `Failed ${job.name}: ${message}`, failed || job);
      this.emitProgress('failed', failed || job);
    }
  }

  private async downloadChunks(job: DownloadJob, directory: FileSystemDirectoryHandle) {
    const partialHandle = await this.committer.getPartialHandle(directory, job.name);
    const writable = await partialHandle.createWritable({ keepExistingData: true });
    let writeChain = Promise.resolve();
    const chunks = job.chunks.filter(chunk => chunk.status !== 'verified');
    let activeDownloadUrl: string | undefined;

    if (job.size === 0) {
      await writable.truncate(0);
      await this.persistChunk(job, { ...job.chunks[0], status: 'verified', updatedAt: new Date().toISOString() });
      await writable.close();
      return;
    }

    const runChunk = async (chunk: DownloadChunk) => {
      let attempt = 0;
      while (attempt <= this.settings.retryCount) {
        attempt += 1;
        try {
          activeDownloadUrl ||= (await this.oneDrive.refreshDownloadUrl(job.itemId)).downloadUrl;
          if (!activeDownloadUrl) throw new Error('Microsoft Graph did not provide a download URL.');

          const updatedChunk = { ...chunk, status: 'downloading' as const, attempts: attempt, updatedAt: new Date().toISOString() };
          await this.persistChunk(job, updatedChunk);

          const response = await this.oneDrive.fetchRange(activeDownloadUrl, chunk.start, chunk.end, this.abortController.signal);
          if (response.status === 401 || response.status === 403 || response.status === 404) {
            activeDownloadUrl = undefined;
            throw new HttpDownloadError('Temporary OneDrive download URL expired; refreshing metadata.', response.status);
          }
          if (!response.ok && response.status !== 206) {
            throw new HttpDownloadError(
              `HTTP ${response.status} while downloading chunk.`,
              response.status,
              parseRetryAfter(response.headers.get('Retry-After')),
            );
          }

          const expectedLength = chunk.end - chunk.start + 1;
          const received = await this.streamRangeToWritable(response, writable, writeChain, chunk.start, job);
          writeChain = received.writeChain;
          await writeChain;
          if (received.bytes !== expectedLength && job.size > 0) {
            throw new Error(`Chunk length mismatch: expected ${expectedLength}, received ${received.bytes}.`);
          }

          const verified = { ...updatedChunk, status: 'verified' as const, error: undefined, updatedAt: new Date().toISOString() };
          await this.persistChunk(job, verified);
          return;
        } catch (error) {
          const transient = error instanceof HttpDownloadError && isTransientStatus(error.status);
          if (attempt > this.settings.retryCount || (!transient && !(error instanceof TypeError))) {
            await this.persistChunk(job, {
              ...chunk,
              status: 'failed',
              attempts: attempt,
              error: error instanceof Error ? error.message : String(error),
              updatedAt: new Date().toISOString(),
            });
            throw error;
          }

          const retryAfter = error instanceof HttpDownloadError ? error.retryAfterSeconds : undefined;
          const delay = backoffDelayMs(attempt, retryAfter);
          if (retryAfter !== undefined) {
            this.throttleEvents += 1;
            const throttled = await this.store.updateJob(job.id, { status: 'throttled' });
            this.reporter.log('downloader', `OneDrive throttled requests; retrying after ${Math.ceil(retryAfter)} seconds.`, throttled || job);
            this.emitProgress('throttled', throttled || job, Date.now() + delay);
          } else {
            const retrying = await this.store.updateJob(job.id, { status: 'retrying' });
            this.reporter.log('downloader', `Retrying ${job.name} after ${Math.round(delay / 1000)} seconds.`, retrying || job);
          }
          await sleep(delay, this.abortController.signal);
        }
      }
    };

    try {
      const queue = [...chunks];
      const workers = Array.from({ length: Math.max(1, this.settings.maxChunksPerFile) }, async () => {
        while (queue.length > 0 && !this.abortController.signal.aborted) {
          const chunk = queue.shift();
          if (chunk) await runChunk(chunk);
        }
      });
      await Promise.all(workers);
    } finally {
      await writeChain;
      await writable.close();
    }
  }

  private async persistChunk(job: DownloadJob, chunk: DownloadChunk) {
    const latest = await this.store.getJob(job.id);
    const source = latest || job;
    const chunks = source.chunks.map(existing => existing.index === chunk.index ? chunk : existing);
    const downloadedBytes = this.completedBytes(chunks);
    const next = await this.store.updateJob(job.id, { chunks, downloadedBytes });
    if (next) this.reporter.job(next);
  }

  private async validateResumeState(job: DownloadJob, directory: FileSystemDirectoryHandle) {
    const verifiedChunks = job.chunks.filter(chunk => chunk.status === 'verified');
    if (verifiedChunks.length === 0) return job;

    const partialFile = await this.committer.getPartialFile(directory, job.name);
    const requiredSize = Math.max(...verifiedChunks.map(chunk => chunk.end + 1));
    if (partialFile && partialFile.size >= requiredSize) return job;

    const chunks = this.createChunks(job.size);
    const reset = await this.store.updateJob(job.id, {
      chunks,
      downloadedBytes: 0,
      status: 'queued',
      errorMessage: partialFile
        ? `Partial file was too short to resume safely (${partialFile.size}/${requiredSize} bytes).`
        : 'Partial file was missing; restarting safely.',
    });
    this.reporter.log('queue', reset?.errorMessage || 'Partial file did not match saved resume state; restarting safely.', reset || job);
    return reset || { ...job, chunks, downloadedBytes: 0, status: 'queued' as const };
  }

  private async streamRangeToWritable(
    response: Response,
    writable: FileSystemWritableFileStream,
    writeChain: Promise<void>,
    start: number,
    job: DownloadJob,
  ) {
    const reader = response.body?.getReader();
    if (!reader) throw new Error('Download response did not expose a readable stream.');

    let position = start;
    let bytes = 0;
    let chain = writeChain;
    while (true) {
      const { done, value } = await reader.read();
      if (done) break;
      const buffer = value.slice();
      const writePosition = position;
      position += buffer.byteLength;
      bytes += buffer.byteLength;
      chain = chain.then(async () => {
        await writable.write({ type: 'seek', position: writePosition });
        await writable.write(buffer);
      });
      this.addDownloadedBytes(buffer.byteLength, job);
    }

    return { bytes, writeChain: chain };
  }

  private async completeJob(
    job: DownloadJob,
    status: 'completed' | 'skipped',
    downloadedBytes: number,
    summary?: string,
    countBytes = false,
    finalName?: string,
  ) {
    const saved = await this.store.updateJob(job.id, {
      status,
      downloadedBytes,
      chunks: job.chunks.map(chunk => ({ ...chunk, status: 'verified' })),
      errorMessage: undefined,
      summary,
    });
    this.counters.completedFiles += 1;
    if (saved) this.reporter.job(saved);
    this.recordManifestItem(manifestItemFromJob(this.runId, saved || job, status, { verification: summary, finalName }));
    if (countBytes) {
      this.addDownloadedBytes(Math.max(0, job.size - job.downloadedBytes), saved || job);
    }
    this.emitProgress(status, saved || job);
  }

  private addDownloadedBytes(bytes: number, job: DownloadJob) {
    if (bytes <= 0) return;
    this.counters.downloadedBytes = Math.min(this.counters.downloadedBytes + bytes, this.counters.totalBytes);
    const now = performance.now();
    if (this.counters.lastSpeedAt > 0) {
      const elapsed = (now - this.counters.lastSpeedAt) / 1000;
      if (elapsed >= 0.5) {
        this.counters.speedBytesPerSecond = (this.counters.downloadedBytes - this.counters.lastSpeedBytes) / elapsed;
        this.counters.lastSpeedBytes = this.counters.downloadedBytes;
        this.counters.lastSpeedAt = now;
      }
    } else {
      this.counters.lastSpeedBytes = this.counters.downloadedBytes;
      this.counters.lastSpeedAt = now;
    }
    this.emitProgress('downloading', job);
  }

  private emitProgress(status: DownloadJob['status'] | 'idle', job?: DownloadJob, throttledUntil?: number) {
    const remainingBytes = Math.max(this.counters.totalBytes - this.counters.downloadedBytes, 0);
    const etaSeconds = this.counters.speedBytesPerSecond > 0 ? remainingBytes / this.counters.speedBytesPerSecond : undefined;
    const stagePercent = this.progressPercent(status, job);
    this.reporter.progress({
      totalBytes: this.counters.totalBytes,
      downloadedBytes: this.counters.downloadedBytes,
      completedFiles: this.counters.completedFiles,
      failedFiles: this.counters.failedFiles,
      queuedFiles: Math.max(this.counters.queuedFiles - this.counters.completedFiles - this.counters.failedFiles, 0),
      stagePercent,
      currentFile: job?.name,
      currentFolder: job?.localPath.split('/').slice(0, -1).join('/'),
      speedBytesPerSecond: this.counters.speedBytesPerSecond,
      etaSeconds,
      status,
      throttledUntil,
    });
  }

  private progressPercent(status: DownloadJob['status'] | 'idle', job?: DownloadJob) {
    if (status === 'scanning') {
      const total = this.scanProcessed + this.scanPending;
      return total > 0 ? Math.min((this.scanProcessed / total) * 100, 99) : 0;
    }
    if (status === 'verifying') return job ? 100 : 0;
    if (job && job.size > 0) return Math.min((job.downloadedBytes / job.size) * 100, 100);
    if (this.counters.totalBytes > 0) return Math.min((this.counters.downloadedBytes / this.counters.totalBytes) * 100, 100);
    return status === 'completed' ? 100 : 0;
  }

  private createChunks(size: number) {
    if (size === 0) {
      return [{ index: 0, start: 0, end: 0, status: 'pending' as const, attempts: 0, updatedAt: new Date().toISOString() }];
    }

    const chunks: DownloadChunk[] = [];
    let start = 0;
    let index = 0;
    while (start < size) {
      const end = Math.min(start + this.settings.chunkSize - 1, size - 1);
      chunks.push({ index, start, end, status: 'pending', attempts: 0, updatedAt: new Date().toISOString() });
      start = end + 1;
      index += 1;
    }
    return chunks;
  }

  private completedBytes(chunks: DownloadChunk[]) {
    return chunks
      .filter(chunk => chunk.status === 'verified')
      .reduce((total, chunk) => total + chunk.end - chunk.start + 1, 0);
  }

  private remoteMatches(job: DownloadJob, metadata: RemoteItemMetadata) {
    return job.itemId === metadata.itemId
      && job.size === metadata.size
      && (!job.eTag || !metadata.eTag || job.eTag === metadata.eTag)
      && (!job.cTag || !metadata.cTag || job.cTag === metadata.cTag);
  }

  private beginRun(mode: DownloadRunMode, scanPending: number) {
    this.abortController = new AbortController();
    this.paused = false;
    this.mode = mode;
    this.runId = createRunId();
    this.startedAt = new Date().toISOString();
    this.manifestItems.clear();
    this.resetCounters();
    this.failedJobs.clear();
    this.scanProcessed = 0;
    this.scanPending = scanPending;
    this.throttleEvents = 0;
  }

  private recordManifestItem(item: RunManifestItem) {
    this.manifestItems.set(`${item.driveId}:${item.itemId}:${item.localPath}`, item);
  }

  private async writeManifest(rootDirectory: FileSystemDirectoryHandle) {
    if (this.mode === 'dry_run') return;
    const items = Array.from(this.manifestItems.values());
    const manifest = createManifest(
      this.runId,
      this.mode,
      rootDirectory.name,
      this.startedAt,
      this.settings,
      this.reporter.getSummary(),
      items,
      this.counters.totalBytes,
      this.counters.completedFiles,
      this.counters.failedFiles,
    );
    try {
      await this.store.saveManifest(manifest);
      this.reporter.log('reporter', `Saved archive manifest in browser storage with ${items.length} item${items.length === 1 ? '' : 's'}.`);
    } catch (error) {
      const message = error instanceof Error ? error.message : String(error);
      this.reporter.log('reporter', `Archive completed, but the manifest could not be saved: ${message}`);
    }
  }

  private parentPartsFromRemotePath(metadata: RemoteItemMetadata) {
    const withoutGraphPrefix = metadata.remotePath.replace(/^.*?:\//, '');
    const parts = withoutGraphPrefix.split('/').filter(Boolean);
    if (parts.at(-1) === metadata.name) parts.pop();
    return parts;
  }

  private jobFromManifestItem(item: RunManifestItem): DownloadJob {
    return {
      id: `${item.driveId}:${item.itemId}:${item.localPath}`,
      driveId: item.driveId,
      itemId: item.itemId,
      name: item.name,
      remotePath: item.remotePath,
      localPath: item.localPath,
      partialPath: `${item.localPath}.partial`,
      size: item.size,
      eTag: item.eTag,
      cTag: item.cTag,
      lastModifiedDateTime: item.lastModifiedDateTime,
      hashes: {
        sha1Hash: item.sha1Hash,
        quickXorHash: item.quickXorHash,
      },
      status: 'queued',
      priority: 0,
      retryCount: 0,
      createdAt: new Date().toISOString(),
      updatedAt: new Date().toISOString(),
      downloadedBytes: 0,
      chunks: this.createChunks(item.size),
    };
  }

  private resetCounters() {
    this.counters = {
      totalBytes: 0,
      downloadedBytes: 0,
      completedFiles: 0,
      failedFiles: 0,
      queuedFiles: 0,
      speedBytesPerSecond: 0,
      lastSpeedBytes: 0,
      lastSpeedAt: 0,
    };
  }
}
