import type { DownloadJob, DownloadJobStatus, RunManifest } from './types';

interface StoredDeltaToken {
  driveId: string;
  token: string;
  updatedAt: string;
}

const databaseName = 'syncpoint-downloads';
const databaseVersion = 3;
const staleJobAgeMs = 14 * 24 * 60 * 60 * 1000;

export class DownloadJobStore {
  private dbPromise?: Promise<IDBDatabase>;

  private open() {
    if (!this.dbPromise) {
      this.dbPromise = new Promise((resolve, reject) => {
        const request = indexedDB.open(databaseName, databaseVersion);
        request.onerror = () => reject(request.error);
        request.onsuccess = () => {
          const db = request.result;
          db.onversionchange = () => {
            db.close();
            this.dbPromise = undefined;
          };
          resolve(db);
        };
        request.onupgradeneeded = () => {
          const db = request.result;
          if (!db.objectStoreNames.contains('jobs')) {
            const jobs = db.createObjectStore('jobs', { keyPath: 'id' });
            jobs.createIndex('status', 'status', { unique: false });
            jobs.createIndex('item', ['driveId', 'itemId'], { unique: false });
            jobs.createIndex('priority', 'priority', { unique: false });
            jobs.createIndex('updatedAt', 'updatedAt', { unique: false });
          }
          if (!db.objectStoreNames.contains('deltaTokens')) {
            db.createObjectStore('deltaTokens', { keyPath: 'driveId' });
          }
          if (!db.objectStoreNames.contains('settings')) {
            db.createObjectStore('settings', { keyPath: 'key' });
          }
          if (!db.objectStoreNames.contains('manifests')) {
            const manifests = db.createObjectStore('manifests', { keyPath: 'runId' });
            manifests.createIndex('finishedAt', 'finishedAt', { unique: false });
          }
          const settings = request.transaction?.objectStore('settings');
          settings?.put({ key: 'schemaVersion', value: databaseVersion, updatedAt: new Date().toISOString() });
        };
      });
    }

    return this.dbPromise;
  }

  private async reopen() {
    const db = await this.dbPromise?.catch(() => undefined);
    db?.close();
    this.dbPromise = undefined;
    return this.open();
  }

  private isInvalidObjectHandle(error: unknown) {
    const message = error instanceof Error ? error.message : String(error);
    return message.includes('ObjectHandle is Invalid') || message.includes('invalid object handle');
  }

  private requestToPromise<T>(request: IDBRequest<T>) {
    return new Promise<T>((resolve, reject) => {
      request.onsuccess = () => resolve(request.result);
      request.onerror = () => reject(request.error);
    });
  }

  private transactionDone(transaction: IDBTransaction) {
    return new Promise<void>((resolve, reject) => {
      transaction.oncomplete = () => resolve();
      transaction.onerror = () => reject(transaction.error);
      transaction.onabort = () => reject(transaction.error);
    });
  }

  private async withStore<T>(
    storeName: string,
    mode: IDBTransactionMode,
    run: (store: IDBObjectStore) => Promise<T>,
    retry = true,
  ) {
    try {
      const db = await this.open();
      const transaction = db.transaction(storeName, mode);
      const done = this.transactionDone(transaction);
      const result = await run(transaction.objectStore(storeName));
      await done;
      return result;
    } catch (error) {
      if (retry && this.isInvalidObjectHandle(error)) {
        await this.reopen();
        return this.withStore(storeName, mode, run, false);
      }
      throw error;
    }
  }

  async upsertJob(job: DownloadJob) {
    const now = new Date().toISOString();
    const next = { ...job, updatedAt: now, createdAt: job.createdAt || now };
    await this.withStore('jobs', 'readwrite', store => this.requestToPromise(store.put(next)));
    return next;
  }

  async getJob(id: string) {
    return this.withStore('jobs', 'readonly', store => this.requestToPromise(store.get(id)));
  }

  async getJobs() {
    return this.withStore('jobs', 'readonly', store => this.requestToPromise(store.getAll()));
  }

  async getJobsByStatus(statuses: DownloadJobStatus[]) {
    const jobs = await this.getJobs();
    return jobs.filter(job => statuses.includes(job.status));
  }

  async updateJob(id: string, patch: Partial<DownloadJob>) {
    const existing = await this.getJob(id);
    if (!existing) return undefined;
    const next = { ...existing, ...patch, updatedAt: new Date().toISOString() };
    await this.upsertJob(next);
    return next;
  }

  async saveDeltaToken(driveId: string, token: string) {
    const record: StoredDeltaToken = { driveId, token, updatedAt: new Date().toISOString() };
    await this.withStore('deltaTokens', 'readwrite', store => this.requestToPromise(store.put(record)));
  }

  async getDeltaToken(driveId: string) {
    const record = await this.withStore<StoredDeltaToken | undefined>('deltaTokens', 'readonly', store => this.requestToPromise(store.get(driveId)));
    return record?.token;
  }

  async saveManifest(manifest: RunManifest) {
    await this.withStore('manifests', 'readwrite', store => this.requestToPromise(store.put(manifest)));
    await this.withStore('settings', 'readwrite', store => this.requestToPromise(store.put({
      key: 'latestManifestRunId',
      value: manifest.runId,
      updatedAt: new Date().toISOString(),
    })));
  }

  async getLatestManifest() {
    const pointer = await this.withStore<{ key: string; value: string } | undefined>(
      'settings',
      'readonly',
      store => this.requestToPromise(store.get('latestManifestRunId')),
    );
    if (pointer?.value) {
      const manifest = await this.withStore<RunManifest | undefined>(
        'manifests',
        'readonly',
        store => this.requestToPromise(store.get(pointer.value)),
      );
      if (manifest) return manifest;
    }

    const manifests = await this.withStore<RunManifest[]>('manifests', 'readonly', store => this.requestToPromise(store.getAll()));
    return manifests.sort((a, b) => Date.parse(b.finishedAt) - Date.parse(a.finishedAt))[0];
  }

  async resetJobs() {
    await this.withStore('jobs', 'readwrite', store => this.requestToPromise(store.clear()));
  }

  async deleteCompletedJobs() {
    const completed = await this.getJobsByStatus(['completed']);
    if (completed.length === 0) return 0;

    await this.withStore('jobs', 'readwrite', async store => {
      await Promise.all(completed.map(job => this.requestToPromise(store.delete(job.id))));
    });
    return completed.length;
  }

  async cleanupStaleJobs(maxAgeMs = staleJobAgeMs) {
    const cutoff = Date.now() - maxAgeMs;
    const jobs = await this.getJobs();
    const stale = jobs.filter(job => {
      const updated = Date.parse(job.updatedAt);
      return Number.isFinite(updated)
        && updated < cutoff
        && ['completed', 'skipped', 'failed', 'cancelled', 'stale_remote_changed'].includes(job.status);
    });
    if (stale.length === 0) return 0;

    await this.withStore('jobs', 'readwrite', async store => {
      await Promise.all(stale.map(job => this.requestToPromise(store.delete(job.id))));
    });
    return stale.length;
  }
}
