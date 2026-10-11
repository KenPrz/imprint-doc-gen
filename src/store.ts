import Database from 'better-sqlite3';
import { mkdirSync } from 'node:fs';
import { join } from 'node:path';
import { createHash, randomBytes, randomUUID } from 'node:crypto';
import { ServiceError } from './service-error.ts';
import type { DocumentFormat } from './types.ts';
import type { RenderInput, RenderResult } from './render-service.ts';

const hash = (value: string) => createHash('sha256').update(value).digest('hex');
const RESULT_RESERVE = 32 * 1024 * 1024;
export const LEASE_MS = 150_000; // Longer than the 30s render + 65s conversion deadlines.
export interface JobInput extends RenderInput { templateId: string; versionId?: string }
export interface Job {
  id: string; tenant_id: string; template_id: string; version_id: string;
  status: 'queued' | 'running' | 'succeeded' | 'failed'; attempts: number;
  created_at: number; updated_at: number; available_at: number; deadline: number;
  error: string | null; format: DocumentFormat | 'pdf' | null;
}
export interface ClaimedJob extends Job { payload: string; lease: string; template: Buffer }
interface Version { id: string; template_id: string; format: DocumentFormat; created_at: number; bytes?: Buffer }
interface Template { id: string; name: string; published_version: string | null; created_at: number }

// The standalone adapter keeps bounded binary values and metadata in one transaction.
// No caller-controlled filesystem paths; a single mounted directory contains the database and WAL.
export class StandaloneStore {
  readonly db: Database.Database;
  readonly now: () => number;
  readonly tenantQuota: number;
  readonly totalQuota: number;
  constructor(directory: string, options: { now?: () => number; tenantQuota?: number; totalQuota?: number } = {}) {
    mkdirSync(directory, { recursive: true, mode: 0o700 });
    this.db = new Database(join(directory, 'imprint.sqlite'));
    this.now = options.now ?? Date.now;
    this.tenantQuota = options.tenantQuota ?? 1024 ** 3;
    this.totalQuota = options.totalQuota ?? 4 * 1024 ** 3;
    this.db.pragma('journal_mode = WAL');
    this.db.pragma('synchronous = FULL');
    this.db.pragma('foreign_keys = ON');
    this.db.pragma('busy_timeout = 5000');
    const version = this.db.pragma('user_version', { simple: true }) as number;
    if (version > 1) { this.db.close(); throw new Error('Database schema is newer than this service'); }
    this.db.transaction(() => {
      this.db.exec(`
        CREATE TABLE IF NOT EXISTS tenants (id TEXT PRIMARY KEY, name TEXT NOT NULL, key_hash TEXT UNIQUE);
        CREATE TABLE IF NOT EXISTS templates (id TEXT PRIMARY KEY, tenant_id TEXT NOT NULL REFERENCES tenants(id), name TEXT NOT NULL, published_version TEXT, created_at INTEGER NOT NULL);
        CREATE INDEX IF NOT EXISTS templates_tenant ON templates(tenant_id);
        CREATE TABLE IF NOT EXISTS versions (id TEXT PRIMARY KEY, template_id TEXT NOT NULL REFERENCES templates(id), format TEXT NOT NULL, bytes BLOB NOT NULL, created_at INTEGER NOT NULL);
        CREATE INDEX IF NOT EXISTS versions_template ON versions(template_id);
        CREATE TABLE IF NOT EXISTS jobs (
          id TEXT PRIMARY KEY, tenant_id TEXT NOT NULL REFERENCES tenants(id), template_id TEXT NOT NULL REFERENCES templates(id), version_id TEXT NOT NULL REFERENCES versions(id),
          status TEXT NOT NULL, attempts INTEGER NOT NULL DEFAULT 0, payload TEXT NOT NULL, reserved_bytes INTEGER NOT NULL,
          idempotency_key TEXT, request_hash TEXT NOT NULL, created_at INTEGER NOT NULL, updated_at INTEGER NOT NULL,
          available_at INTEGER NOT NULL, deadline INTEGER NOT NULL, lease TEXT, lease_until INTEGER, error TEXT, format TEXT, result BLOB,
          UNIQUE(tenant_id, idempotency_key)
        );
        CREATE INDEX IF NOT EXISTS jobs_queue ON jobs(status, available_at, created_at);
        CREATE INDEX IF NOT EXISTS jobs_tenant ON jobs(tenant_id, created_at);
        INSERT OR IGNORE INTO tenants(id, name) VALUES ('default', 'Default');
        PRAGMA user_version = 1;
      `);
    })();
  }
  close() { this.db.close(); }
  authenticate(key: string): string | undefined {
    return (this.db.prepare('SELECT id FROM tenants WHERE key_hash = ?').get(hash(key)) as { id: string } | undefined)?.id;
  }
  listTenants() { return this.db.prepare('SELECT id,name,key_hash IS NOT NULL AS keyActive FROM tenants ORDER BY name LIMIT 1000').all(); }
  createTenant(name: string) {
    if ((this.db.prepare('SELECT COUNT(*) AS n FROM tenants').get() as { n: number }).n >= 1000) throw new ServiceError(413, 'TENANT_LIMIT');
    const id = randomUUID(), apiKey = randomBytes(32).toString('base64url');
    this.db.prepare('INSERT INTO tenants(id, name, key_hash) VALUES (?, ?, ?)').run(id, name, hash(apiKey));
    return { id, name, apiKey };
  }
  rotateKey(id: string, revoke = false) {
    if (id === 'default') throw new ServiceError(409, 'DEFAULT_KEY_USE_ENV');
    const apiKey = randomBytes(32).toString('base64url');
    if (!this.db.prepare('UPDATE tenants SET key_hash = ? WHERE id = ?').run(revoke ? null : hash(apiKey), id).changes) throw new ServiceError(404, 'NOT_FOUND');
    return revoke ? { id } : { id, apiKey };
  }
  private usage(tenant?: string): number {
    const templates = this.db.prepare(`SELECT COALESCE(SUM(length(v.bytes)),0) AS n FROM versions v JOIN templates t ON t.id=v.template_id ${tenant ? 'WHERE t.tenant_id=?' : ''}`).get(...(tenant ? [tenant] : [])) as { n: number };
    const jobs = this.db.prepare(`SELECT COALESCE(SUM(reserved_bytes),0) AS n FROM jobs ${tenant ? 'WHERE tenant_id=?' : ''}`).get(...(tenant ? [tenant] : [])) as { n: number };
    return templates.n + jobs.n;
  }
  private capacity(tenant: string, bytes: number) {
    if (this.usage(tenant) + bytes > this.tenantQuota || this.usage() + bytes > this.totalQuota) throw new ServiceError(413, 'STORAGE_QUOTA', 'Delete completed jobs or unused templates to release storage');
  }
  template(tenant: string, id: string): Template {
    const row = this.db.prepare('SELECT id,name,published_version,created_at FROM templates WHERE tenant_id=? AND id=?').get(tenant, id) as Template | undefined;
    if (!row) throw new ServiceError(404, 'NOT_FOUND');
    return row;
  }
  listTemplates(tenant: string) { return this.db.prepare('SELECT id,name,published_version,created_at FROM templates WHERE tenant_id=? ORDER BY created_at DESC LIMIT 100').all(tenant) as Template[]; }
  versions(tenant: string, id: string) {
    this.template(tenant, id);
    return this.db.prepare('SELECT id,template_id,format,created_at FROM versions WHERE template_id=? ORDER BY created_at,id').all(id) as Version[];
  }
  upload(tenant: string, name: string, bytes: Buffer, format: DocumentFormat, templateId?: string) {
    return this.db.transaction(() => {
      if (templateId) this.template(tenant, templateId);
      this.capacity(tenant, bytes.length);
      const count = this.db.prepare('SELECT COUNT(*) AS n FROM versions v JOIN templates t ON t.id=v.template_id WHERE t.tenant_id=?').get(tenant) as { n: number };
      if (count.n >= 1000) throw new ServiceError(413, 'VERSION_LIMIT');
      const id = templateId ?? randomUUID(), versionId = randomUUID(), now = this.now();
      if (!templateId) this.db.prepare('INSERT INTO templates(id,tenant_id,name,created_at) VALUES (?,?,?,?)').run(id, tenant, name, now);
      this.db.prepare('INSERT INTO versions VALUES (?,?,?,?,?)').run(versionId, id, format, bytes, now);
      return { templateId: id, versionId, format, createdAt: now };
    }).immediate();
  }
  publish(tenant: string, id: string, versionId: string) {
    this.template(tenant, id);
    if (!this.db.prepare('SELECT id FROM versions WHERE id=? AND template_id=?').get(versionId, id)) throw new ServiceError(404, 'NOT_FOUND');
    this.db.prepare('UPDATE templates SET published_version=? WHERE id=? AND tenant_id=?').run(versionId, id, tenant);
    return { templateId: id, versionId };
  }
  deleteTemplate(tenant: string, id: string) {
    this.db.transaction(() => {
      this.template(tenant, id);
      if (this.db.prepare('SELECT id FROM jobs WHERE template_id=? LIMIT 1').get(id)) throw new ServiceError(409, 'TEMPLATE_IN_USE', 'Delete its terminal jobs first');
      this.db.prepare('DELETE FROM versions WHERE template_id=?').run(id);
      this.db.prepare('DELETE FROM templates WHERE id=? AND tenant_id=?').run(id, tenant);
    }).immediate();
  }
  enqueue(tenant: string, input: JobInput, idempotencyKey?: string): { job: Job; reused: boolean } {
    return this.db.transaction(() => {
      const payload = JSON.stringify(input), fingerprint = hash(payload);
      if (idempotencyKey) {
        const existing = this.db.prepare('SELECT id,request_hash FROM jobs WHERE tenant_id=? AND idempotency_key=?').get(tenant, idempotencyKey) as { id: string; request_hash: string } | undefined;
        if (existing) {
          if (existing.request_hash !== fingerprint) throw new ServiceError(409, 'IDEMPOTENCY_CONFLICT');
          return { job: this.job(tenant, existing.id), reused: true };
        }
      }
      const template = this.template(tenant, input.templateId);
      const versionId = input.versionId ?? template.published_version;
      if (!versionId) throw new ServiceError(409, 'TEMPLATE_NOT_PUBLISHED');
      const version = this.db.prepare('SELECT format FROM versions WHERE id=? AND template_id=?').get(versionId, input.templateId) as { format: string } | undefined;
      if (!version) throw new ServiceError(404, 'NOT_FOUND');
      if (input.output && input.output !== 'pdf' && input.output !== version.format) throw new ServiceError(422, 'OUTPUT_FORMAT_MISMATCH');
      const stored = this.db.prepare('SELECT COUNT(*) AS total, SUM(tenant_id=?) AS tenant FROM jobs').get(tenant) as { total: number; tenant: number };
      if (stored.total >= 10000 || stored.tenant >= 2000) throw new ServiceError(413, 'JOB_HISTORY_LIMIT', 'Delete completed jobs before submitting more');
      const counts = this.db.prepare("SELECT COUNT(*) AS total, SUM(tenant_id=?) AS tenant FROM jobs WHERE status IN ('queued','running')").get(tenant) as { total: number; tenant: number };
      if (counts.total >= 1000 || counts.tenant >= 100) throw new ServiceError(503, 'QUEUE_FULL');
      const reserved = Buffer.byteLength(payload) + RESULT_RESERVE;
      this.capacity(tenant, reserved);
      const id = randomUUID(), now = this.now();
      this.db.prepare(`INSERT INTO jobs(id,tenant_id,template_id,version_id,status,payload,reserved_bytes,idempotency_key,request_hash,created_at,updated_at,available_at,deadline)
        VALUES (?,?,?,?,'queued',?,?,?,?,?,?,?,?)`).run(id, tenant, input.templateId, versionId, payload, reserved, idempotencyKey ?? null, fingerprint, now, now, now, now + 86_400_000);
      return { job: this.job(tenant, id), reused: false };
    }).immediate();
  }
  job(tenant: string, id: string): Job {
    const row = this.db.prepare('SELECT id,tenant_id,template_id,version_id,status,attempts,created_at,updated_at,available_at,deadline,error,format FROM jobs WHERE tenant_id=? AND id=?').get(tenant, id) as Job | undefined;
    if (!row) throw new ServiceError(404, 'NOT_FOUND');
    return row;
  }
  result(tenant: string, id: string): RenderResult {
    const job = this.job(tenant, id);
    if (job.status !== 'succeeded') throw new ServiceError(409, 'RESULT_NOT_READY');
    const row = this.db.prepare('SELECT result FROM jobs WHERE tenant_id=? AND id=?').get(tenant, id) as { result: Buffer };
    return { bytes: row.result, format: job.format! };
  }
  deleteJob(tenant: string, id: string) {
    if (['running', 'queued'].includes(this.job(tenant, id).status)) throw new ServiceError(409, 'JOB_ACTIVE');
    this.db.prepare('DELETE FROM jobs WHERE tenant_id=? AND id=?').run(tenant, id);
  }
  claim(): ClaimedJob | undefined {
    return this.db.transaction(() => {
      const now = this.now();
      // Interrupted attempts have a finite lease. Never let an old worker overwrite a reclaimed result.
      this.db.prepare(`UPDATE jobs SET status=CASE WHEN attempts>=2 OR deadline<=? THEN 'failed' ELSE 'queued' END,
        error='WORKER_INTERRUPTED',lease=NULL,lease_until=NULL,updated_at=?,available_at=? WHERE status='running' AND lease_until<=?`).run(now, now, now, now);
      this.db.prepare("UPDATE jobs SET status='failed',error='JOB_EXPIRED',updated_at=? WHERE status='queued' AND deadline<=?").run(now, now);
      this.db.prepare("UPDATE jobs SET payload='{}',reserved_bytes=0 WHERE status='failed' AND reserved_bytes<>0").run();
      const row = this.db.prepare("SELECT * FROM jobs WHERE status='queued' AND available_at<=? ORDER BY created_at,id LIMIT 1").get(now) as (Job & { payload: string }) | undefined;
      if (!row) return;
      const lease = randomUUID();
      this.db.prepare("UPDATE jobs SET status='running',attempts=attempts+1,lease=?,lease_until=?,updated_at=? WHERE id=?").run(lease, now + LEASE_MS, now, row.id);
      const version = this.db.prepare('SELECT bytes FROM versions WHERE id=?').get(row.version_id) as { bytes: Buffer };
      return { ...row, status: 'running' as const, attempts: row.attempts + 1, lease, template: version.bytes };
    }).immediate();
  }
  succeed(job: ClaimedJob, result: RenderResult) {
    if (result.bytes.length > RESULT_RESERVE) throw new ServiceError(413, 'LIMIT_EXCEEDED');
    this.db.prepare("UPDATE jobs SET status='succeeded',result=?,format=?,payload='{}',reserved_bytes=?,error=NULL,lease=NULL,lease_until=NULL,updated_at=? WHERE id=? AND status='running' AND lease=?")
      .run(result.bytes, result.format, result.bytes.length, this.now(), job.id, job.lease);
  }
  fail(job: ClaimedJob, code: string, retryable: boolean) {
    const retry = retryable && job.attempts < 2 && this.now() + 2000 < job.deadline;
    this.db.prepare(`UPDATE jobs SET status=?,error=?,available_at=?,updated_at=?,lease=NULL,lease_until=NULL,
      payload=CASE WHEN ? THEN payload ELSE '{}' END,reserved_bytes=CASE WHEN ? THEN reserved_bytes ELSE 0 END WHERE id=? AND status='running' AND lease=?`)
      .run(retry ? 'queued' : 'failed', code, this.now() + 2000, this.now(), Number(retry), Number(retry), job.id, job.lease);
  }
  cleanup() {
    this.db.prepare("DELETE FROM jobs WHERE status IN ('succeeded','failed') AND updated_at<?").run(this.now() - 7 * 86_400_000);
  }
}
