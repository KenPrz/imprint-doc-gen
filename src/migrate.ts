import Database from 'better-sqlite3';
import { randomUUID } from 'node:crypto';
import { digest } from './object-store.ts';
import type { PostgresStore } from './postgres-store.ts';
// Input must be an offline backup, never the live service file. Destination must be empty.
export async function migrateSqlite(snapshot: string, target: PostgresStore) {
  const source = new Database(snapshot, { readonly: true, fileMustExist: true });
  const uploaded: string[] = [];
  try {
    source.exec('BEGIN');
    if (
      source.pragma('user_version', { simple: true }) !== 1 ||
      source.pragma('integrity_check', { simple: true }) !== 'ok' ||
      (source.pragma('foreign_key_check') as unknown[]).length
    )
      throw new Error('Source schema or integrity check failed');
    if ((source.prepare("SELECT COUNT(*) AS n FROM jobs WHERE status='running'").get() as { n: number }).n)
      throw new Error('Source contains running jobs; stop and drain workers before backing up');
    return await target.transaction(async (client) => {
      await client.query('SELECT pg_advisory_xact_lock(774302)');
      await client.query(
        'LOCK TABLE tenants,templates,versions,jobs,objects,webhooks IN ACCESS EXCLUSIVE MODE',
      );
      for (const table of ['templates', 'versions', 'jobs', 'objects', 'webhooks'])
        if (Number((await client.query(`SELECT COUNT(*) AS n FROM ${table}`)).rows[0].n))
          throw new Error('Migration destination is not empty');
      const tenants = await client.query("SELECT * FROM tenants WHERE id<>'default' OR key_hash IS NOT NULL");
      if (tenants.rowCount) throw new Error('Migration destination has tenant credentials');
      const insert = async (table: string, row: Record<string, unknown>) => {
        const keys = Object.keys(row);
        if (keys.some((k) => !/^[_a-z]+$/.test(k))) throw new Error('Unexpected source column');
        await client.query(
          `INSERT INTO ${table} (${keys.map((k) => '"' + k + '"').join(',')}) VALUES (${keys.map((_, n) => '$' + (n + 1)).join(',')})`,
          Object.values(row),
        );
      };
      let totalBytes = 0;
      const upload = async (bytes: Buffer) => {
        if (!Buffer.isBuffer(bytes) || bytes.length > 32 * 1024 * 1024)
          throw new Error('Invalid source blob');
        totalBytes += bytes.length;
        if (totalBytes > target.totalQuota) throw new Error('Migration exceeds destination storage quota');
        const key = randomUUID(),
          hash = digest(bytes);
        uploaded.push(key);
        await target.objects.put(key, bytes);
        const verified = await target.objects.get(key, bytes.length, hash);
        if (!verified.equals(bytes)) throw new Error('Migration object verification failed');
        await insert('objects', { key, bytes: bytes.length, digest: hash, created_at: Date.now() });
        return key;
      };
      const counts: Record<string, number> = {};
      await client.query("DELETE FROM tenants WHERE id='default'");
      for (const table of ['tenants', 'templates', 'versions', 'jobs']) {
        counts[table] = 0;
        for (const raw of source.prepare(`SELECT * FROM ${table}`).iterate()) {
          const row = { ...(raw as Record<string, unknown>) };
          if (table === 'versions') {
            row.object_key = await upload(row.bytes as Buffer);
            delete row.bytes;
          }
          if (table === 'jobs') {
            row.result_key = row.result === null ? null : await upload(row.result as Buffer);
            delete row.result;
          }
          await insert(table, row);
          counts[table]++;
        }
        if (Number((await client.query(`SELECT COUNT(*) AS n FROM ${table}`)).rows[0].n) !== counts[table])
          throw new Error('Migration row count mismatch');
      }
      // Verify ownership and published references before making the metadata visible.
      const invalid =
        await client.query(`SELECT t.id FROM templates t LEFT JOIN versions v ON v.id=t.published_version AND v.template_id=t.id WHERE t.published_version IS NOT NULL AND v.id IS NULL
        UNION ALL SELECT j.id FROM jobs j JOIN templates t ON t.id=j.template_id JOIN versions v ON v.id=j.version_id WHERE j.tenant_id<>t.tenant_id OR v.template_id<>t.id`);
      if (invalid.rowCount) throw new Error('Source ownership or publication references are inconsistent');
      const quota = await client.query(
        `SELECT t.tenant_id FROM templates t JOIN versions v ON v.template_id=t.id JOIN objects o ON o.key=v.object_key GROUP BY t.tenant_id HAVING SUM(o.bytes)+COALESCE((SELECT SUM(j.reserved_bytes) FROM jobs j WHERE j.tenant_id=t.tenant_id),0)>$1`,
        [target.tenantQuota],
      );
      const usage = await client.query(
        'SELECT COALESCE((SELECT SUM(o.bytes) FROM versions v JOIN objects o ON o.key=v.object_key),0)+COALESCE((SELECT SUM(reserved_bytes) FROM jobs),0) AS n',
      );
      if (Number(usage.rows[0].n) > target.totalQuota)
        throw new Error('Migration exceeds total reserved storage quota');
      if (quota.rowCount) throw new Error('Migration exceeds tenant storage quota');
      return { counts, verifiedObjects: uploaded.length, verifiedBytes: totalBytes };
    });
  } catch (error) {
    const cleanup = await Promise.allSettled(
      uploaded.map(async (key) => {
        // A lost COMMIT acknowledgement must never delete a committed object.
        const row = await target.pool.query('SELECT key FROM objects WHERE key=$1', [key]);
        if (!row.rowCount) await target.objects.delete(key);
      }),
    );
    if (cleanup.some((r) => r.status === 'rejected'))
      throw new AggregateError([error], 'Migration rolled back; some unreferenced objects require cleanup');
    throw error;
  } finally {
    source.close();
  }
}
