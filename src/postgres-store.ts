import pg from 'pg';
import { randomUUID, randomBytes, createHash } from 'node:crypto';
import { ServiceError } from './service-error.ts';
import { LEASE_MS, type Job, type JobInput, type ClaimedJob } from './store.ts';
import type { ServiceStore } from './store-interface.ts';
import type { DocumentFormat } from './types.ts';
import type { RenderResult } from './render-service.ts';
import { digest, type ObjectStore } from './object-store.ts';
const hash = (value: string) => createHash('sha256').update(value).digest('hex');
const reserve = 32 * 1024 * 1024;
const fields =
  'id,tenant_id,template_id,version_id,status,attempts,created_at,updated_at,available_at,deadline,error,format';
type Query = Pick<pg.PoolClient, 'query'>;
export class PostgresStore implements ServiceStore {
  readonly kind = 'postgres' as const;
  readonly pool: pg.Pool;
  readonly objects: ObjectStore;
  readonly now: () => number;
  readonly tenantQuota: number;
  readonly totalQuota: number;
  constructor(
    url: string,
    objects: ObjectStore,
    options: { now?: () => number; tenantQuota?: number; totalQuota?: number } = {},
  ) {
    this.pool = new pg.Pool({
      connectionString: url,
      max: 10,
      connectionTimeoutMillis: 10000,
      statement_timeout: 10000,
      options: '-c search_path=imprint,public',
    });
    this.pool.on('error', () => {
      /* Idle connections are replaced; readiness reports outages without leaking connection strings. */
    });
    this.objects = objects;
    this.now = options.now ?? Date.now;
    this.tenantQuota = options.tenantQuota ?? 1024 ** 3;
    this.totalQuota = options.totalQuota ?? 4 * 1024 ** 3;
  }
  async initialize() {
    await this.transaction(async (c) => {
      await c.query('SELECT pg_advisory_xact_lock(774301)');
      await c.query(`CREATE SCHEMA IF NOT EXISTS imprint;
   CREATE TABLE IF NOT EXISTS imprint.schema_version(version INTEGER PRIMARY KEY);
   INSERT INTO imprint.schema_version VALUES(1) ON CONFLICT DO NOTHING;`);
      const version = await c.query('SELECT max(version) AS version FROM imprint.schema_version');
      if (version.rows[0].version !== 1) throw new Error('Unsupported database schema');
      await c.query(`
   CREATE TABLE IF NOT EXISTS tenants(id TEXT PRIMARY KEY,name TEXT NOT NULL,key_hash TEXT UNIQUE);
   CREATE TABLE IF NOT EXISTS objects(key TEXT PRIMARY KEY,bytes INTEGER NOT NULL,digest TEXT NOT NULL,created_at DOUBLE PRECISION NOT NULL);
   CREATE TABLE IF NOT EXISTS templates(id TEXT PRIMARY KEY,tenant_id TEXT NOT NULL REFERENCES tenants(id),name TEXT NOT NULL,published_version TEXT,created_at DOUBLE PRECISION NOT NULL);
   CREATE TABLE IF NOT EXISTS versions(id TEXT PRIMARY KEY,template_id TEXT NOT NULL REFERENCES templates(id),format TEXT NOT NULL,object_key TEXT NOT NULL REFERENCES objects(key),created_at DOUBLE PRECISION NOT NULL);
   CREATE TABLE IF NOT EXISTS jobs(id TEXT PRIMARY KEY,tenant_id TEXT NOT NULL REFERENCES tenants(id),template_id TEXT NOT NULL REFERENCES templates(id),version_id TEXT NOT NULL REFERENCES versions(id),status TEXT NOT NULL,attempts INTEGER NOT NULL DEFAULT 0,payload TEXT NOT NULL,reserved_bytes INTEGER NOT NULL,idempotency_key TEXT,request_hash TEXT NOT NULL,created_at DOUBLE PRECISION NOT NULL,updated_at DOUBLE PRECISION NOT NULL,available_at DOUBLE PRECISION NOT NULL,deadline DOUBLE PRECISION NOT NULL,lease TEXT,lease_until DOUBLE PRECISION,error TEXT,format TEXT,result_key TEXT REFERENCES objects(key),UNIQUE(tenant_id,idempotency_key));
   CREATE INDEX IF NOT EXISTS jobs_queue ON jobs(status,available_at,created_at);
   CREATE INDEX IF NOT EXISTS jobs_tenant ON jobs(tenant_id,created_at);
   CREATE INDEX IF NOT EXISTS templates_tenant ON templates(tenant_id);
   CREATE INDEX IF NOT EXISTS versions_template ON versions(template_id);
   CREATE TABLE IF NOT EXISTS webhooks(id TEXT PRIMARY KEY,job_id TEXT UNIQUE NOT NULL,payload TEXT NOT NULL,attempts INTEGER NOT NULL DEFAULT 0,available_at DOUBLE PRECISION NOT NULL,lease TEXT,lease_until DOUBLE PRECISION,delivered_at DOUBLE PRECISION,created_at DOUBLE PRECISION NOT NULL,last_error TEXT);
   INSERT INTO tenants(id,name) VALUES('default','Default') ON CONFLICT DO NOTHING;
  `);
    });
    await this.ready();
  }
  async transaction<T>(fn: (client: pg.PoolClient) => Promise<T>): Promise<T> {
    const client = await this.pool.connect();
    try {
      await client.query('BEGIN');
      const value = await fn(client);
      await client.query('COMMIT');
      return value;
    } catch (error) {
      await client.query('ROLLBACK').catch(() => {});
      throw error;
    } finally {
      client.release();
    }
  }
  private async lock(c: Query) {
    await c.query('SELECT pg_advisory_xact_lock(774302)');
  }
  async ready() {
    await this.pool.query('SELECT 1');
    await this.objects.ready();
  }
  async queued() {
    return Number((await this.pool.query("SELECT COUNT(*) AS n FROM jobs WHERE status='queued'")).rows[0].n);
  }
  async close() {
    await this.pool.end();
    this.objects.close();
  }
  async authenticate(key: string) {
    return (await this.pool.query('SELECT id FROM tenants WHERE key_hash=$1', [hash(key)])).rows[0]?.id as
      string | undefined;
  }
  async listTenants() {
    return (
      await this.pool.query(
        'SELECT id,name,key_hash IS NOT NULL AS "keyActive" FROM tenants ORDER BY name LIMIT 1000',
      )
    ).rows;
  }
  async createTenant(name: string) {
    return this.transaction(async (c) => {
      await this.lock(c);
      if (Number((await c.query('SELECT COUNT(*) AS n FROM tenants')).rows[0].n) >= 1000)
        throw new ServiceError(413, 'TENANT_LIMIT');
      const id = randomUUID(),
        apiKey = randomBytes(32).toString('base64url');
      await c.query('INSERT INTO tenants VALUES($1,$2,$3)', [id, name, hash(apiKey)]);
      return { id, name, apiKey };
    });
  }
  async rotateKey(id: string, revoke = false) {
    if (id === 'default') throw new ServiceError(409, 'DEFAULT_KEY_USE_ENV');
    const apiKey = randomBytes(32).toString('base64url');
    if (
      !(
        await this.pool.query('UPDATE tenants SET key_hash=$1 WHERE id=$2', [
          revoke ? null : hash(apiKey),
          id,
        ])
      ).rowCount
    )
      throw new ServiceError(404, 'NOT_FOUND');
    return revoke ? { id } : { id, apiKey };
  }
  private async getTemplate(c: Query, tenant: string, id: string) {
    const row = (
      await c.query(
        'SELECT id,name,published_version,created_at FROM templates WHERE tenant_id=$1 AND id=$2',
        [tenant, id],
      )
    ).rows[0];
    if (!row) throw new ServiceError(404, 'NOT_FOUND');
    return row as { id: string; name: string; published_version: string | null; created_at: number };
  }
  template(tenant: string, id: string) {
    return this.getTemplate(this.pool, tenant, id);
  }
  async listTemplates(tenant: string) {
    return (
      await this.pool.query(
        'SELECT id,name,published_version,created_at FROM templates WHERE tenant_id=$1 ORDER BY created_at DESC LIMIT 10',
        [tenant],
      )
    ).rows as Awaited<ReturnType<ServiceStore['listTemplates']>>;
  }
  async versions(tenant: string, id: string) {
    await this.template(tenant, id);
    return (
      await this.pool.query(
        'SELECT id,template_id,format,created_at FROM versions WHERE template_id=$1 ORDER BY created_at,id',
        [id],
      )
    ).rows as Awaited<ReturnType<ServiceStore['versions']>>;
  }
  private async capacity(c: Query, tenant: string, bytes: number) {
    const result = await c.query(
      `SELECT COALESCE(SUM(n),0) AS total,COALESCE(SUM(n) FILTER(WHERE tenant_id=$1),0) AS tenant FROM(SELECT t.tenant_id,o.bytes AS n FROM versions v JOIN templates t ON t.id=v.template_id JOIN objects o ON o.key=v.object_key UNION ALL SELECT tenant_id,reserved_bytes FROM jobs) u`,
      [tenant],
    );
    if (
      Number(result.rows[0].total) + bytes > this.totalQuota ||
      Number(result.rows[0].tenant) + bytes > this.tenantQuota
    )
      throw new ServiceError(413, 'STORAGE_QUOTA');
  }
  private async put(bytes: Buffer) {
    if (bytes.length > reserve) throw new ServiceError(413, 'LIMIT_EXCEEDED');
    const key = randomUUID();
    await this.pool.query('INSERT INTO objects VALUES($1,$2,$3,$4)', [
      key,
      bytes.length,
      digest(bytes),
      this.now(),
    ]);
    try {
      await this.objects.put(key, bytes);
    } catch {
      throw new ServiceError(502, 'OBJECT_UNAVAILABLE');
    }
    return key;
  }
  private async get(key: string) {
    const row = (await this.pool.query('SELECT bytes,digest FROM objects WHERE key=$1', [key])).rows[0];
    if (!row) throw new ServiceError(502, 'OBJECT_UNAVAILABLE');
    try {
      return await this.objects.get(key, row.bytes, row.digest);
    } catch {
      throw new ServiceError(502, 'OBJECT_UNAVAILABLE');
    }
  }
  async upload(tenant: string, name: string, bytes: Buffer, format: DocumentFormat, templateId?: string) {
    // Validate ownership before writing the staged object. Staged objects are tracked for garbage collection.
    if (templateId) await this.template(tenant, templateId);
    await this.capacity(this.pool, tenant, bytes.length);
    const key = await this.put(bytes);
    return this.transaction(async (c) => {
      await this.lock(c);
      if (templateId) await this.getTemplate(c, tenant, templateId);
      await this.capacity(c, tenant, bytes.length);
      if (
        Number(
          (
            await c.query(
              'SELECT COUNT(*) AS n FROM versions v JOIN templates t ON t.id=v.template_id WHERE t.tenant_id=$1',
              [tenant],
            )
          ).rows[0].n,
        ) >= 1000
      )
        throw new ServiceError(413, 'VERSION_LIMIT');
      const id = templateId ?? randomUUID(),
        versionId = randomUUID(),
        now = this.now();
      if (!templateId)
        await c.query('INSERT INTO templates(id,tenant_id,name,created_at) VALUES($1,$2,$3,$4)', [
          id,
          tenant,
          name,
          now,
        ]);
      await c.query('INSERT INTO versions VALUES($1,$2,$3,$4,$5)', [versionId, id, format, key, now]);
      return { templateId: id, versionId, format, createdAt: now };
    });
  }
  async publish(tenant: string, id: string, versionId: string) {
    return this.transaction(async (c) => {
      await this.lock(c);
      await this.getTemplate(c, tenant, id);
      if (
        !(await c.query('SELECT id FROM versions WHERE id=$1 AND template_id=$2', [versionId, id])).rowCount
      )
        throw new ServiceError(404, 'NOT_FOUND');
      await c.query('UPDATE templates SET published_version=$1 WHERE id=$2', [versionId, id]);
      return { templateId: id, versionId };
    });
  }
  async deleteTemplate(tenant: string, id: string) {
    await this.transaction(async (c) => {
      await this.lock(c);
      await this.getTemplate(c, tenant, id);
      if ((await c.query('SELECT id FROM jobs WHERE template_id=$1 LIMIT 1', [id])).rowCount)
        throw new ServiceError(409, 'TEMPLATE_IN_USE');
      await c.query('DELETE FROM versions WHERE template_id=$1', [id]);
      await c.query('DELETE FROM templates WHERE id=$1', [id]);
    });
  }
  async enqueue(tenant: string, input: JobInput, idempotencyKey?: string) {
    return this.transaction(async (c) => {
      await this.lock(c);
      const payload = JSON.stringify(input),
        fingerprint = hash(payload);
      if (idempotencyKey) {
        const existing = (
          await c.query('SELECT * FROM jobs WHERE tenant_id=$1 AND idempotency_key=$2', [
            tenant,
            idempotencyKey,
          ])
        ).rows[0];
        if (existing) {
          if (existing.request_hash !== fingerprint) throw new ServiceError(409, 'IDEMPOTENCY_CONFLICT');
          return { job: existing as Job, reused: true };
        }
      }
      const template = await this.getTemplate(c, tenant, input.templateId),
        versionId = input.versionId ?? template.published_version;
      if (!versionId) throw new ServiceError(409, 'TEMPLATE_NOT_PUBLISHED');
      const version = (
        await c.query('SELECT format FROM versions WHERE id=$1 AND template_id=$2', [
          versionId,
          input.templateId,
        ])
      ).rows[0];
      if (!version) throw new ServiceError(404, 'NOT_FOUND');
      if (input.output && input.output !== 'pdf' && input.output !== version.format)
        throw new ServiceError(422, 'OUTPUT_FORMAT_MISMATCH');
      const counts = (
        await c.query(
          `SELECT COUNT(*) AS total,COUNT(*) FILTER(WHERE tenant_id=$1) AS tenant,COUNT(*) FILTER(WHERE status IN('queued','running')) AS active,COUNT(*) FILTER(WHERE tenant_id=$1 AND status IN('queued','running')) AS tenant_active FROM jobs`,
          [tenant],
        )
      ).rows[0];
      if (Number(counts.total) >= 10000 || Number(counts.tenant) >= 2000)
        throw new ServiceError(413, 'JOB_HISTORY_LIMIT');
      if (Number(counts.active) >= 1000 || Number(counts.tenant_active) >= 100)
        throw new ServiceError(503, 'QUEUE_FULL');
      const reserved = Buffer.byteLength(payload) + reserve;
      await this.capacity(c, tenant, reserved);
      const id = randomUUID(),
        now = this.now();
      const result = await c.query(
        `INSERT INTO jobs(id,tenant_id,template_id,version_id,status,payload,reserved_bytes,idempotency_key,request_hash,created_at,updated_at,available_at,deadline) VALUES($1,$2,$3,$4,'queued',$5,$6,$7,$8,$9,$9,$9,$10) RETURNING ${fields}`,
        [
          id,
          tenant,
          input.templateId,
          versionId,
          payload,
          reserved,
          idempotencyKey ?? null,
          fingerprint,
          now,
          now + 86400000,
        ],
      );
      return { job: result.rows[0] as Job, reused: false };
    });
  }
  async job(tenant: string, id: string) {
    const row = (
      await this.pool.query(`SELECT ${fields} FROM jobs WHERE tenant_id=$1 AND id=$2`, [tenant, id])
    ).rows[0];
    if (!row) throw new ServiceError(404, 'NOT_FOUND');
    return row as Job;
  }
  async result(tenant: string, id: string): Promise<RenderResult> {
    const row = (
      await this.pool.query('SELECT status,result_key,format FROM jobs WHERE tenant_id=$1 AND id=$2', [
        tenant,
        id,
      ])
    ).rows[0];
    if (!row) throw new ServiceError(404, 'NOT_FOUND');
    if (row.status !== 'succeeded') throw new ServiceError(409, 'RESULT_NOT_READY');
    return { bytes: await this.get(row.result_key), format: row.format };
  }
  async deleteJob(tenant: string, id: string) {
    await this.transaction(async (c) => {
      await this.lock(c);
      const row = (
        await c.query('SELECT status FROM jobs WHERE tenant_id=$1 AND id=$2 FOR UPDATE', [tenant, id])
      ).rows[0];
      if (!row) throw new ServiceError(404, 'NOT_FOUND');
      if (['queued', 'running'].includes(row.status)) throw new ServiceError(409, 'JOB_ACTIVE');
      await c.query('DELETE FROM jobs WHERE id=$1', [id]);
    });
  }
  private async events(c: Query) {
    // Inserted in the same transaction as terminal state changes; payload never contains document data.
    await c.query(
      `INSERT INTO webhooks(id,job_id,payload,available_at,created_at) SELECT id,id,json_build_object('id',id,'type','job.'||status,'tenantId',tenant_id,'jobId',id,'status',status,'format',format,'error',error)::text,$1,$1 FROM jobs WHERE status IN('succeeded','failed') ON CONFLICT(job_id) DO NOTHING`,
      [this.now()],
    );
  }
  async claim(): Promise<ClaimedJob | undefined> {
    const claimed = await this.transaction(async (c) => {
      const now = this.now();
      await c.query(
        `UPDATE jobs SET status=CASE WHEN attempts>=2 OR deadline<=$1 THEN 'failed' ELSE 'queued' END,error='WORKER_INTERRUPTED',lease=NULL,lease_until=NULL,updated_at=$1,available_at=$1 WHERE status='running' AND lease_until<=$1`,
        [now],
      );
      await c.query(
        "UPDATE jobs SET status='failed',error='JOB_EXPIRED',updated_at=$1 WHERE status='queued' AND deadline<=$1",
        [now],
      );
      await c.query(
        "UPDATE jobs SET payload='{}',reserved_bytes=0 WHERE status='failed' AND reserved_bytes<>0",
      );
      await this.events(c);
      const row = (
        await c.query(
          "SELECT * FROM jobs WHERE status='queued' AND available_at<=$1 ORDER BY created_at,id FOR UPDATE SKIP LOCKED LIMIT 1",
          [now],
        )
      ).rows[0];
      if (!row) return;
      const lease = randomUUID();
      await c.query(
        "UPDATE jobs SET status='running',attempts=attempts+1,lease=$1,lease_until=$2,updated_at=$3 WHERE id=$4",
        [lease, now + LEASE_MS, now, row.id],
      );
      const version = (await c.query('SELECT object_key FROM versions WHERE id=$1', [row.version_id]))
        .rows[0];
      return { ...row, status: 'running', attempts: row.attempts + 1, lease, key: version.object_key };
    });
    if (!claimed) return;
    try {
      return { ...claimed, template: await this.get(claimed.key) } as ClaimedJob;
    } catch (error) {
      await this.fail(claimed, 'OBJECT_UNAVAILABLE', true);
      throw error;
    }
  }
  async succeed(job: ClaimedJob, result: RenderResult) {
    if (result.bytes.length > reserve) throw new ServiceError(413, 'LIMIT_EXCEEDED');
    const key = await this.put(result.bytes);
    await this.transaction(async (c) => {
      await c.query(
        "UPDATE jobs SET status='succeeded',result_key=$1,format=$2,payload='{}',reserved_bytes=$3,error=NULL,lease=NULL,lease_until=NULL,updated_at=$4 WHERE id=$5 AND status='running' AND lease=$6",
        [key, result.format, result.bytes.length, this.now(), job.id, job.lease],
      );
      await this.events(c);
    });
  }
  async fail(job: ClaimedJob, code: string, retryable: boolean) {
    await this.transaction(async (c) => {
      const retry = retryable && job.attempts < 2 && this.now() + 2000 < job.deadline;
      await c.query(
        `UPDATE jobs SET status=$1,error=$2,available_at=$3,updated_at=$4,lease=NULL,lease_until=NULL,payload=CASE WHEN $5 THEN payload ELSE '{}' END,reserved_bytes=CASE WHEN $5 THEN reserved_bytes ELSE 0 END WHERE id=$6 AND status='running' AND lease=$7`,
        [retry ? 'queued' : 'failed', code, this.now() + 2000, this.now(), retry, job.id, job.lease],
      );
      await this.events(c);
    });
  }
  async cleanup() {
    await this.pool.query("DELETE FROM jobs WHERE status IN('succeeded','failed') AND updated_at<$1", [
      this.now() - 7 * 86400000,
    ]);
    await this.pool.query('DELETE FROM webhooks WHERE created_at<$1', [this.now() - 7 * 86400000]);
    // Objects are immutable; the one-hour staging grace exceeds all upload and transaction deadlines.
    const orphaned = await this.pool.query(
      'SELECT key FROM objects o WHERE created_at<$1 AND NOT EXISTS(SELECT 1 FROM versions v WHERE v.object_key=o.key) AND NOT EXISTS(SELECT 1 FROM jobs j WHERE j.result_key=o.key) LIMIT 10',
      [this.now() - 3600000],
    );
    for (const { key } of orphaned.rows) {
      await this.objects.delete(key);
      await this.pool.query(
        'DELETE FROM objects o WHERE key=$1 AND NOT EXISTS(SELECT 1 FROM versions v WHERE v.object_key=o.key) AND NOT EXISTS(SELECT 1 FROM jobs j WHERE j.result_key=o.key)',
        [key],
      );
    }
  }
}
