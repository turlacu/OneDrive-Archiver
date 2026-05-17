import fs from 'node:fs';
import path from 'node:path';
import { DatabaseSync } from 'node:sqlite';
import type { DownloadProgressSnapshot, RemoteItemMetadata } from '../download/types';
import type { ServerSourceSelection } from './serverDownloadEngine';

export interface ServerUser {
  email: string;
  displayName?: string;
}

export interface PersistedServerJob {
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
}

export interface ServerArchiveRecord {
  userEmail: string;
  item: RemoteItemMetadata;
  localPath: string;
  status: 'completed' | 'skipped';
  verificationMessage?: string;
  updatedAt: string;
}

interface JobRow {
  id: string;
  user_email: string;
  mode: PersistedServerJob['mode'];
  status: PersistedServerJob['status'];
  created_at: string;
  updated_at: string;
  target_root: string;
  selections_json: string;
  settings_json: string;
  log_json: string;
  snapshot_json: string;
}

interface ArchiveRecordRow {
  user_email: string;
  remote_path: string;
  item_id: string;
  local_path: string;
  status: ServerArchiveRecord['status'];
  verification_message: string | null;
  item_json: string;
  updated_at: string;
}

export class ServerStateStore {
  private readonly db: DatabaseSync;

  constructor(appDataDir: string) {
    fs.mkdirSync(appDataDir, { recursive: true });
    this.db = new DatabaseSync(path.join(appDataDir, 'onedrive-archiver.db'));
    this.db.exec(`
      CREATE TABLE IF NOT EXISTS users (
        email TEXT PRIMARY KEY,
        display_name TEXT,
        updated_at TEXT NOT NULL
      );

      CREATE TABLE IF NOT EXISTS server_jobs (
        id TEXT PRIMARY KEY,
        user_email TEXT NOT NULL,
        mode TEXT NOT NULL,
        status TEXT NOT NULL,
        created_at TEXT NOT NULL,
        updated_at TEXT NOT NULL,
        target_root TEXT NOT NULL,
        selections_json TEXT NOT NULL,
        settings_json TEXT NOT NULL,
        log_json TEXT NOT NULL,
        snapshot_json TEXT NOT NULL
      );

      CREATE INDEX IF NOT EXISTS idx_server_jobs_user_updated
        ON server_jobs(user_email, updated_at DESC);

      CREATE TABLE IF NOT EXISTS delta_tokens (
        user_email TEXT PRIMARY KEY,
        token TEXT NOT NULL,
        updated_at TEXT NOT NULL
      );

      CREATE TABLE IF NOT EXISTS server_archive_records (
        user_email TEXT NOT NULL,
        remote_path TEXT NOT NULL,
        item_id TEXT NOT NULL,
        local_path TEXT NOT NULL,
        status TEXT NOT NULL,
        verification_message TEXT,
        item_json TEXT NOT NULL,
        updated_at TEXT NOT NULL,
        PRIMARY KEY (user_email, remote_path)
      );

      CREATE INDEX IF NOT EXISTS idx_server_archive_records_user_updated
        ON server_archive_records(user_email, updated_at DESC);
    `);
  }

  upsertUser(user: ServerUser) {
    this.db.prepare(`
      INSERT INTO users (email, display_name, updated_at)
      VALUES (?, ?, ?)
      ON CONFLICT(email) DO UPDATE SET
        display_name = excluded.display_name,
        updated_at = excluded.updated_at
    `).run(user.email, user.displayName || null, new Date().toISOString());
  }

  listJobs(userEmail: string) {
    const rows = this.db.prepare(`
      SELECT * FROM server_jobs
      WHERE user_email = ?
      ORDER BY updated_at DESC
      LIMIT 50
    `).all(userEmail) as unknown as JobRow[];
    return rows.map(row => this.rowToJob(row));
  }

  getJob(userEmail: string, id: string) {
    const row = this.db.prepare(`
      SELECT * FROM server_jobs
      WHERE user_email = ? AND id = ?
    `).get(userEmail, id) as unknown as JobRow | undefined;
    return row ? this.rowToJob(row) : undefined;
  }

  upsertJob(job: PersistedServerJob) {
    this.db.prepare(`
      INSERT INTO server_jobs (
        id, user_email, mode, status, created_at, updated_at, target_root,
        selections_json, settings_json, log_json, snapshot_json
      )
      VALUES (?, ?, ?, ?, ?, ?, ?, ?, ?, ?, ?)
      ON CONFLICT(id) DO UPDATE SET
        status = excluded.status,
        updated_at = excluded.updated_at,
        target_root = excluded.target_root,
        selections_json = excluded.selections_json,
        settings_json = excluded.settings_json,
        log_json = excluded.log_json,
        snapshot_json = excluded.snapshot_json
    `).run(
      job.id,
      job.userEmail,
      job.mode,
      job.status,
      job.createdAt,
      job.updatedAt,
      job.targetRoot,
      JSON.stringify(job.selections),
      JSON.stringify(job.settings),
      JSON.stringify(job.log),
      JSON.stringify(job.snapshot),
    );
  }

  updateJob(job: PersistedServerJob) {
    job.updatedAt = new Date().toISOString();
    this.upsertJob(job);
  }

  listResumableJobs(userEmail: string) {
    const rows = this.db.prepare(`
      SELECT * FROM server_jobs
      WHERE user_email = ?
        AND status IN ('queued', 'scanning', 'downloading', 'interrupted')
      ORDER BY created_at ASC
    `).all(userEmail) as unknown as JobRow[];
    return rows.map(row => this.rowToJob(row));
  }

  markActiveJobsInterrupted() {
    this.db.prepare(`
      UPDATE server_jobs
      SET status = 'interrupted', updated_at = ?
      WHERE status IN ('queued', 'scanning', 'downloading')
    `).run(new Date().toISOString());
  }

  getDeltaToken(userEmail: string) {
    const row = this.db.prepare(`
      SELECT token FROM delta_tokens
      WHERE user_email = ?
    `).get(userEmail) as { token: string } | undefined;
    return row?.token;
  }

  saveDeltaToken(userEmail: string, token: string) {
    this.db.prepare(`
      INSERT INTO delta_tokens (user_email, token, updated_at)
      VALUES (?, ?, ?)
      ON CONFLICT(user_email) DO UPDATE SET
        token = excluded.token,
        updated_at = excluded.updated_at
    `).run(userEmail, token, new Date().toISOString());
  }

  upsertArchiveRecord(record: Omit<ServerArchiveRecord, 'updatedAt'> & { updatedAt?: string }) {
    const updatedAt = record.updatedAt || new Date().toISOString();
    this.db.prepare(`
      INSERT INTO server_archive_records (
        user_email, remote_path, item_id, local_path, status,
        verification_message, item_json, updated_at
      )
      VALUES (?, ?, ?, ?, ?, ?, ?, ?)
      ON CONFLICT(user_email, remote_path) DO UPDATE SET
        item_id = excluded.item_id,
        local_path = excluded.local_path,
        status = excluded.status,
        verification_message = excluded.verification_message,
        item_json = excluded.item_json,
        updated_at = excluded.updated_at
    `).run(
      record.userEmail,
      record.item.remotePath,
      record.item.itemId,
      record.localPath,
      record.status,
      record.verificationMessage || null,
      JSON.stringify(record.item),
      updatedAt,
    );
  }

  listArchiveRecords(userEmail: string) {
    const rows = this.db.prepare(`
      SELECT * FROM server_archive_records
      WHERE user_email = ?
      ORDER BY updated_at DESC
    `).all(userEmail) as unknown as ArchiveRecordRow[];
    return rows.map(row => this.rowToArchiveRecord(row));
  }

  private rowToJob(row: JobRow): PersistedServerJob {
    return {
      id: row.id,
      userEmail: row.user_email,
      mode: row.mode,
      status: row.status,
      createdAt: row.created_at,
      updatedAt: row.updated_at,
      targetRoot: row.target_root,
      selections: JSON.parse(row.selections_json),
      settings: JSON.parse(row.settings_json),
      log: JSON.parse(row.log_json),
      snapshot: JSON.parse(row.snapshot_json),
    };
  }

  private rowToArchiveRecord(row: ArchiveRecordRow): ServerArchiveRecord {
    return {
      userEmail: row.user_email,
      item: JSON.parse(row.item_json),
      localPath: row.local_path,
      status: row.status,
      verificationMessage: row.verification_message || undefined,
      updatedAt: row.updated_at,
    };
  }
}
