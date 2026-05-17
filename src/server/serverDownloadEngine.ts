import fs from 'node:fs/promises';
import { createWriteStream } from 'node:fs';
import path from 'node:path';
import { Readable } from 'node:stream';
import { pipeline } from 'node:stream/promises';
import { normalizeRelativePath, shouldIncludeFile } from '../download/pathTools';
import { defaultDownloadSettings, type DownloadSettings, type DownloadProgressSnapshot, type DownloadSummary, type RemoteHashes, type RemoteItemMetadata } from '../download/types';
import { ServerStateStore, type PersistedServerJob, type ServerUser } from './serverStateStore';

export interface ServerSourceSelection {
  id: string;
  name: string;
  type: 'file' | 'folder';
}

interface ServerGraphDriveItem {
  id: string;
  name: string;
  size?: number;
  eTag?: string;
  cTag?: string;
  lastModifiedDateTime?: string;
  folder?: unknown;
  file?: { hashes?: RemoteHashes };
  package?: unknown;
  deleted?: unknown;
  parentReference?: {
    driveId?: string;
    path?: string;
  };
  '@microsoft.graph.downloadUrl'?: string;
}

interface ServerDownloadJob {
  id: string;
  userEmail: string;
  mode: 'start' | 'dry-run' | 'repair' | 'incremental';
  status: 'queued' | 'scanning' | 'downloading' | 'completed' | 'failed' | 'cancelled' | 'interrupted';
  createdAt: string;
  updatedAt: string;
  targetRoot: string;
  selections: ServerSourceSelection[];
  settings: Record<string, unknown>;
  log: string[];
  snapshot: DownloadProgressSnapshot;
  abortController: AbortController;
  nextDeltaToken?: string;
}

const emptySummary: DownloadSummary = {
  downloaded: 0,
  resumed: 0,
  verified: 0,
  skipped: 0,
  failed: 0,
  changedRemotely: 0,
  conflicts: 0,
  insufficientDiskSpace: 0,
};

export function resolveInsideRoot(root: string, relativePath: string) {
  const resolvedRoot = path.resolve(root);
  const resolvedPath = path.resolve(resolvedRoot, relativePath);
  const relative = path.relative(resolvedRoot, resolvedPath);
  if (relative.startsWith('..') || path.isAbsolute(relative)) {
    throw new Error('Resolved download path escapes SERVER_DOWNLOAD_ROOT.');
  }
  return resolvedPath;
}

export function sanitizeUserFolder(email: string) {
  return email.toLowerCase().replace(/[^a-z0-9._-]+/g, '_').replace(/^[._-]+|[._-]+$/g, '') || 'user';
}

class ServerOneDriveClient {
  constructor(private readonly getAccessToken: () => Promise<string>) {}

  async getItem(itemId: string) {
    return this.normalizeItem(await this.graphGet(`/me/drive/items/${itemId}`));
  }

  async listChildren(folderId: string) {
    const endpoint = folderId === 'root'
      ? '/me/drive/root/children'
      : `/me/drive/items/${folderId}/children`;
    return this.readPagedItems(endpoint);
  }

  async delta(token?: string) {
    const endpoint = token || '/me/drive/root/delta?$select=id,name,size,eTag,cTag,lastModifiedDateTime,folder,file,package,parentReference&$top=200';
    let response = await this.graphGet(endpoint);
    const items: RemoteItemMetadata[] = [];
    let deltaToken: string | undefined;

    while (response) {
      for (const item of response.value || []) {
        if (!item.deleted && !item.folder && !item.package) items.push(this.normalizeItem(item));
      }
      deltaToken = response['@odata.deltaLink'];
      response = response['@odata.nextLink'] ? await this.graphGet(response['@odata.nextLink']) : null;
    }

    return { items, deltaToken };
  }

  async refreshDownloadUrl(itemId: string) {
    const metadata = await this.getItem(itemId);
    if (metadata.downloadUrl) return metadata;
    const accessToken = await this.getAccessToken();
    const response = await fetch(`https://graph.microsoft.com/v1.0/me/drive/items/${itemId}/content`, {
      headers: { Authorization: `Bearer ${accessToken}` },
      redirect: 'manual',
    });
    const location = response.headers.get('Location');
    return location ? { ...metadata, downloadUrl: location } : metadata;
  }

  private async readPagedItems(endpoint: string) {
    const items: ServerGraphDriveItem[] = [];
    let response = await this.graphGet(`${endpoint}?$select=id,name,size,eTag,cTag,lastModifiedDateTime,folder,file,package,parentReference&$top=200`);
    while (response) {
      items.push(...(response.value || []));
      const nextLink = response['@odata.nextLink'];
      response = nextLink ? await this.graphGet(nextLink) : null;
    }
    return items;
  }

  private async graphGet(endpointOrUrl: string) {
    const accessToken = await this.getAccessToken();
    const url = endpointOrUrl.startsWith('https://')
      ? endpointOrUrl
      : `https://graph.microsoft.com/v1.0${endpointOrUrl}`;
    const response = await fetch(url, {
      headers: { Authorization: `Bearer ${accessToken}` },
    });
    if (!response.ok) {
      throw new Error(`Microsoft Graph ${response.status}: ${await response.text()}`);
    }
    return response.json();
  }

  normalizeItem(item: ServerGraphDriveItem): RemoteItemMetadata {
    const pathValue = item.parentReference?.path || '';
    const remotePath = `${pathValue}/${item.name}`.replace(/^\/+/, '');
    return {
      driveId: item.parentReference?.driveId || 'me',
      itemId: item.id,
      name: item.name,
      remotePath,
      size: item.size || 0,
      eTag: item.eTag,
      cTag: item.cTag,
      lastModifiedDateTime: item.lastModifiedDateTime,
      hashes: item.file?.hashes || {},
      downloadUrl: item['@microsoft.graph.downloadUrl'],
    };
  }
}

export class ServerDownloadManager {
  private jobs = new Map<string, ServerDownloadJob>();
  private activeJobIds = new Set<string>();

  constructor(
    private readonly targetRoot?: string,
    private readonly store?: ServerStateStore,
  ) {
    this.store?.markActiveJobsInterrupted();
  }

  get configuredRoot() {
    return this.targetRoot;
  }

  targetRootForUser(userEmail: string) {
    return this.userTargetRoot(userEmail);
  }

  listJobs(userEmail: string) {
    if (this.store) return this.store.listJobs(userEmail).map(job => this.publicJob(job));
    return Array.from(this.jobs.values())
      .filter(job => job.userEmail === userEmail)
      .map(job => this.publicJob(job));
  }

  getJob(userEmail: string, id: string) {
    const job = this.jobs.get(id);
    if (job?.userEmail === userEmail) return this.publicJob(job);
    const persisted = this.store?.getJob(userEmail, id);
    return persisted ? this.publicJob(persisted) : undefined;
  }

  cancel(userEmail: string, id: string) {
    const job = this.jobs.get(id);
    if (!job || job.userEmail !== userEmail) return false;
    job.abortController.abort();
    job.status = 'cancelled';
    job.updatedAt = new Date().toISOString();
    this.log(job, 'Cancelled by user.');
    this.persist(job);
    return true;
  }

  start(
    user: ServerUser,
    mode: 'start' | 'dry-run' | 'repair' | 'incremental',
    selections: ServerSourceSelection[],
    settings: Partial<DownloadSettings>,
    getAccessToken: () => Promise<string>,
  ) {
    if (!this.targetRoot) throw new Error('SERVER_DOWNLOAD_ROOT is not configured.');
    if (mode !== 'incremental' && selections.length === 0) {
      throw new Error('Select OneDrive files or folders before starting a server job.');
    }
    const job = this.createJob(user, mode, selections, settings as Record<string, unknown>);
    this.jobs.set(job.id, job);
    this.persist(job);
    void this.run(job, getAccessToken);
    return this.publicJob(job);
  }

  resumeForUser(userEmail: string, getAccessToken: () => Promise<string>) {
    const jobs = this.store?.listResumableJobs(userEmail) || [];
    for (const persisted of jobs) {
      if (this.activeJobIds.has(persisted.id)) continue;
      const job = this.fromPersistedJob(persisted);
      this.jobs.set(job.id, job);
      this.log(job, 'Resuming server job after reconnect.');
      void this.run(job, getAccessToken);
    }
  }

  private createJob(
    user: ServerUser,
    mode: ServerDownloadJob['mode'],
    selections: ServerSourceSelection[],
    settings: Record<string, unknown>,
  ): ServerDownloadJob {
    const now = new Date().toISOString();
    return {
      id: `${Date.now()}-${Math.random().toString(36).slice(2)}`,
      userEmail: user.email,
      mode,
      status: 'queued',
      createdAt: now,
      updatedAt: now,
      targetRoot: this.userTargetRoot(user.email),
      selections,
      settings,
      log: [],
      abortController: new AbortController(),
      snapshot: {
        totalBytes: 0,
        downloadedBytes: 0,
        completedFiles: 0,
        failedFiles: 0,
        queuedFiles: 0,
        stagePercent: 0,
        speedBytesPerSecond: 0,
        status: 'queued',
        summary: { ...emptySummary },
      },
    };
  }

  private fromPersistedJob(job: PersistedServerJob): ServerDownloadJob {
    return {
      ...job,
      abortController: new AbortController(),
    };
  }

  private async run(job: ServerDownloadJob, getAccessToken: () => Promise<string>) {
    if (this.activeJobIds.has(job.id)) return;
    this.activeJobIds.add(job.id);
    const settings = { ...defaultDownloadSettings, ...job.settings };
    const oneDrive = new ServerOneDriveClient(getAccessToken);
    try {
      job.status = 'scanning';
      job.snapshot.status = 'scanning';
      this.persist(job);
      const items = job.mode === 'incremental'
        ? await this.scanIncremental(oneDrive, job)
        : await this.scanSelections(oneDrive, job.selections, [], settings, job);
      const runnable = job.mode === 'repair'
        ? await this.filterMissingOrChanged(items, job)
        : items;
      job.snapshot.totalBytes = runnable.reduce((total, item) => total + item.size, 0);
      job.snapshot.queuedFiles = runnable.length;
      job.snapshot.status = job.mode === 'dry-run' ? 'completed' : 'queued';
      if (job.mode === 'dry-run') {
        job.status = 'completed';
        this.log(job, `Dry run found ${runnable.length} file${runnable.length === 1 ? '' : 's'} and ${job.snapshot.totalBytes} bytes.`);
        this.persist(job);
        return;
      }
      this.log(job, `${job.mode === 'repair' ? 'Repair' : job.mode === 'incremental' ? 'Incremental archive' : 'Archive'} queued ${runnable.length} file${runnable.length === 1 ? '' : 's'}.`);
      job.status = 'downloading';
      job.snapshot.status = 'downloading';
      this.persist(job);
      for (const item of runnable) {
        if (job.abortController.signal.aborted) throw new DOMException('Operation cancelled', 'AbortError');
        await this.downloadItem(oneDrive, item, settings, job);
      }
      job.status = 'completed';
      job.snapshot.status = 'completed';
      job.snapshot.stagePercent = 100;
      if (job.mode === 'incremental' && job.snapshot.failedFiles === 0 && job.nextDeltaToken) {
        this.store?.saveDeltaToken(job.userEmail, job.nextDeltaToken);
        this.log(job, 'Saved OneDrive delta token after successful incremental archive.');
      }
      this.log(job, `Completed ${job.snapshot.completedFiles} file${job.snapshot.completedFiles === 1 ? '' : 's'} with ${job.snapshot.failedFiles} failure${job.snapshot.failedFiles === 1 ? '' : 's'}.`);
    } catch (error) {
      if (job.abortController.signal.aborted) {
        job.status = 'cancelled';
        job.snapshot.status = 'cancelled';
        this.log(job, 'Server job cancelled.');
      } else {
        job.status = 'failed';
        job.snapshot.status = 'failed';
        job.snapshot.failedFiles += 1;
        this.log(job, `Server job failed: ${error instanceof Error ? error.message : String(error)}`);
      }
    } finally {
      job.updatedAt = new Date().toISOString();
      this.persist(job);
      this.activeJobIds.delete(job.id);
    }
  }

  private async scanIncremental(oneDrive: ServerOneDriveClient, job: ServerDownloadJob) {
    this.log(job, 'Checking OneDrive delta changes...');
    const token = this.store?.getDeltaToken(job.userEmail);
    this.log(job, token ? 'Using saved OneDrive delta token.' : 'No saved delta token found; creating baseline.');
    const delta = await oneDrive.delta(token);
    job.nextDeltaToken = delta.deltaToken;
    if (!token) {
      if (delta.deltaToken) this.store?.saveDeltaToken(job.userEmail, delta.deltaToken);
      this.log(job, 'Baseline delta token saved. Incremental archive will download changes from the next run.');
      return [];
    }
    this.log(job, `OneDrive delta returned ${delta.items.length} changed file${delta.items.length === 1 ? '' : 's'}.`);
    return delta.items.map(item => {
      const remotePath = item.remotePath.includes('/root:')
        ? item.remotePath.split('/root:').pop() || item.name
        : item.remotePath;
      return { ...item, remotePath: normalizeRelativePath([remotePath]) };
    });
  }

  private async scanSelections(
    oneDrive: ServerOneDriveClient,
    selections: ServerSourceSelection[],
    parentParts: string[],
    settings: DownloadSettings,
    job: ServerDownloadJob,
  ) {
    const items: RemoteItemMetadata[] = [];
    for (const selection of selections) {
      if (selection.type === 'folder') {
        const folderParts = [...parentParts, selection.name];
        const children = await oneDrive.listChildren(selection.id);
        for (const child of children) {
          if (child.folder || child.package) {
            items.push(...await this.scanSelections(oneDrive, [{ id: child.id, name: child.name, type: 'folder' }], folderParts, settings, job));
          } else {
            this.addIfIncluded(items, oneDrive.normalizeItem(child), folderParts, settings);
          }
        }
      } else {
        this.addIfIncluded(items, await oneDrive.getItem(selection.id), parentParts, settings);
      }
      job.snapshot.stagePercent = Math.min(job.snapshot.stagePercent + 1, 99);
      job.updatedAt = new Date().toISOString();
      this.persist(job);
    }
    return items;
  }

  private addIfIncluded(items: RemoteItemMetadata[], item: RemoteItemMetadata, parentParts: string[], settings: DownloadSettings) {
    if (!shouldIncludeFile(item.name, settings.includeExtensions, settings.excludeExtensions)) return;
    if (settings.maximumFileSize > 0 && item.size > settings.maximumFileSize) return;
    items.push({ ...item, remotePath: normalizeRelativePath([...parentParts, item.name]) });
  }

  private async filterMissingOrChanged(items: RemoteItemMetadata[], job: ServerDownloadJob) {
    const missing: RemoteItemMetadata[] = [];
    for (const item of items) {
      const finalPath = resolveInsideRoot(job.targetRoot, item.remotePath);
      try {
        const stat = await fs.stat(finalPath);
        if (stat.size !== item.size) missing.push(item);
      } catch {
        missing.push(item);
      }
    }
    return missing;
  }

  private async downloadItem(
    oneDrive: ServerOneDriveClient,
    item: RemoteItemMetadata,
    settings: DownloadSettings,
    job: ServerDownloadJob,
  ) {
    const finalPath = await this.resolveFinalPath(item, settings, job);
    if (!finalPath) {
      job.snapshot.completedFiles += 1;
      job.updatedAt = new Date().toISOString();
      this.persist(job);
      return;
    }
    const partialPath = `${finalPath}.partial`;
    await fs.mkdir(path.dirname(finalPath), { recursive: true });
    const metadata = await oneDrive.refreshDownloadUrl(item.itemId);
    if (!metadata.downloadUrl) throw new Error(`Microsoft Graph did not provide a download URL for ${item.name}.`);
    const response = await fetch(metadata.downloadUrl, { signal: job.abortController.signal });
    if (!response.ok || !response.body) throw new Error(`HTTP ${response.status} while downloading ${item.name}.`);
    job.snapshot.currentFile = item.name;
    this.log(job, `Downloading ${item.remotePath}`);
    const startingDownloadedBytes = job.snapshot.downloadedBytes;
    let currentFileBytes = 0;
    let lastProgressAt = 0;
    let lastSpeedAt = Date.now();
    let lastSpeedBytes = startingDownloadedBytes;
    const stream = Readable.fromWeb(response.body as any);
    stream.on('data', (chunk: Buffer) => {
      currentFileBytes += chunk.length;
      job.snapshot.downloadedBytes = Math.min(startingDownloadedBytes + currentFileBytes, job.snapshot.totalBytes);
      const now = Date.now();
      const elapsedSeconds = (now - lastSpeedAt) / 1000;
      if (elapsedSeconds >= 1) {
        job.snapshot.speedBytesPerSecond = Math.max(0, (job.snapshot.downloadedBytes - lastSpeedBytes) / elapsedSeconds);
        lastSpeedAt = now;
        lastSpeedBytes = job.snapshot.downloadedBytes;
      }
      if (job.snapshot.totalBytes > 0) {
        job.snapshot.stagePercent = Math.min((job.snapshot.downloadedBytes / job.snapshot.totalBytes) * 100, 100);
      }
      const remainingBytes = Math.max(job.snapshot.totalBytes - job.snapshot.downloadedBytes, 0);
      job.snapshot.etaSeconds = job.snapshot.speedBytesPerSecond > 0
        ? remainingBytes / job.snapshot.speedBytesPerSecond
        : undefined;
      if (now - lastProgressAt >= 1000) {
        job.updatedAt = new Date().toISOString();
        this.persist(job);
        lastProgressAt = now;
      }
    });
    await pipeline(stream, createWriteStream(partialPath));
    const stat = await fs.stat(partialPath);
    if (stat.size !== item.size) throw new Error(`Size mismatch for ${item.name}: expected ${item.size}, wrote ${stat.size}.`);
    await fs.rename(partialPath, finalPath);
    job.snapshot.completedFiles += 1;
    job.snapshot.downloadedBytes = Math.min(startingDownloadedBytes + item.size, job.snapshot.totalBytes);
    job.snapshot.stagePercent = job.snapshot.totalBytes > 0
      ? Math.min((job.snapshot.downloadedBytes / job.snapshot.totalBytes) * 100, 100)
      : 100;
    job.snapshot.summary.downloaded += 1;
    job.updatedAt = new Date().toISOString();
    this.persist(job);
  }

  private async resolveFinalPath(item: RemoteItemMetadata, settings: DownloadSettings, job: ServerDownloadJob) {
    const requested = resolveInsideRoot(job.targetRoot, item.remotePath);
    if (settings.conflictStrategy === 'overwrite') return requested;
    try {
      const existing = await fs.stat(requested);
      if (settings.conflictStrategy === 'skip_existing') {
        if (existing.size === item.size) {
          job.snapshot.summary.skipped += 1;
          this.log(job, `Skipped existing file: ${item.remotePath}`);
          return null;
        }
        return requested;
      }
    } catch (error) {
      if ((error as NodeJS.ErrnoException).code === 'ENOENT') return requested;
    }
    const parsed = path.parse(requested);
    for (let index = 1; index < 10000; index += 1) {
      const candidate = path.join(parsed.dir, `${parsed.name} (${index})${parsed.ext}`);
      try {
        await fs.stat(candidate);
      } catch {
        return candidate;
      }
    }
    throw new Error(`Unable to allocate conflict-free filename for ${item.remotePath}`);
  }

  private log(job: ServerDownloadJob, message: string) {
    job.log.unshift(`${new Date().toLocaleTimeString()} ${message}`);
    job.log = job.log.slice(0, 200);
    job.updatedAt = new Date().toISOString();
    this.persist(job);
  }

  private userTargetRoot(userEmail: string) {
    if (!this.targetRoot) return '';
    return path.join(this.targetRoot, 'users', sanitizeUserFolder(userEmail));
  }

  private persist(job: ServerDownloadJob) {
    this.store?.updateJob({
      id: job.id,
      userEmail: job.userEmail,
      mode: job.mode,
      status: job.status,
      createdAt: job.createdAt,
      updatedAt: job.updatedAt,
      targetRoot: job.targetRoot,
      selections: job.selections,
      settings: job.settings,
      log: job.log,
      snapshot: job.snapshot,
    });
  }

  private publicJob(job: ServerDownloadJob | PersistedServerJob) {
    return {
      id: job.id,
      mode: job.mode,
      status: job.status,
      createdAt: job.createdAt,
      updatedAt: job.updatedAt,
      targetRoot: job.targetRoot,
      log: job.log,
      snapshot: job.snapshot,
    };
  }
}
