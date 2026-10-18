import test from 'node:test';
import assert from 'node:assert/strict';
import pg from 'pg';
import { randomUUID } from 'node:crypto';
import { mkdtemp, rm } from 'node:fs/promises';
import { tmpdir } from 'node:os';
import { join } from 'node:path';
import { S3Client, CreateBucketCommand, ListObjectsV2Command } from '@aws-sdk/client-s3';
import { PostgresStore } from '../../src/postgres-store.ts';
import { S3Objects } from '../../src/object-store.ts';
import { StandaloneStore } from '../../src/store.ts';
import { migrateSqlite } from '../../src/migrate.ts';
import { encodeTextTemplate } from '../../src/text-template.ts';
import { backupSqlite } from '../../src/backup.ts';
const url = process.env.TEST_DATABASE_URL;
if (!url) throw Error('TEST_DATABASE_URL must name disposable infrastructure');
test('Offline migration verifies blobs, rolls back failed writes and preserves tenant/job/version identities', async () => {
  const dbname = 'migration_' + randomUUID().replaceAll('-', ''),
    admin = new pg.Pool({ connectionString: url });
  await admin.query(`CREATE DATABASE "${dbname}"`);
  const targetUrl = new URL(url!);
  targetUrl.pathname = '/' + dbname;
  const endpoint = process.env.S3_ENDPOINT ?? 'http://127.0.0.1:18400',
    bucket = 'migration-' + randomUUID();
  const client = new S3Client({ region: 'us-east-1', endpoint, forcePathStyle: true });
  await client.send(new CreateBucketCommand({ Bucket: bucket }));
  const objects = new S3Objects({ bucket, endpoint }),
    target = new PostgresStore(targetUrl.toString(), objects);
  const dir = await mkdtemp(join(tmpdir(), 'imprint-migration-'));
  const source = new StandaloneStore(dir);
  try {
    await target.initialize();
    const owner = source.createTenant('Migrated tenant');
    const template = encodeTextTemplate('txt', 'Hello {d.name}');
    const v = source.upload(owner.id, 'Greeting', template, 'txt');
    source.publish(owner.id, v.templateId, v.versionId);
    const input = { templateId: v.templateId, data: { name: 'World' } };
    const done = source.enqueue(owner.id, input, 'completed');
    const claim = source.claim()!;
    source.succeed(claim, { bytes: Buffer.from('Hello World'), format: 'txt' });
    const queued = source.enqueue(owner.id, input, 'pending');
    const snapshot = join(dir, 'backup.sqlite');
    await backupSqlite(dir, snapshot);
    const get = objects.get.bind(objects);
    let fail = true;
    objects.get = async (...args) => {
      if (fail) throw Error('injected read failure');
      return get(...args);
    };
    await assert.rejects(migrateSqlite(snapshot, target), /injected read failure/);
    assert.equal(Number((await target.pool.query('SELECT count(*) AS n FROM versions')).rows[0].n), 0);
    assert.equal((await target.listTenants()).length, 1);
    assert.equal((await client.send(new ListObjectsV2Command({ Bucket: bucket }))).Contents?.length ?? 0, 0);
    fail = false;
    const report = await migrateSqlite(snapshot, target);
    assert.equal(report.verifiedObjects, 2);
    assert.equal(report.counts.jobs, 2);
    assert.equal(await target.authenticate(owner.apiKey), owner.id);
    assert.deepEqual((await target.result(owner.id, done.job.id)).bytes, Buffer.from('Hello World'));
    assert.equal((await target.enqueue(owner.id, input, 'pending')).job.id, queued.job.id);
    const next = await target.claim();
    assert.equal(next!.id, queued.job.id);
    assert(next!.template.equals(template));
    await assert.rejects(migrateSqlite(snapshot, target), /not empty/);
    assert.deepEqual((await target.result(owner.id, done.job.id)).bytes, Buffer.from('Hello World'));
    const running = source.claim();
    assert(running);
    const active = join(dir, 'active.sqlite');
    await backupSqlite(dir, active);
    await assert.rejects(migrateSqlite(active, target), /running jobs/);
  } finally {
    source.close();
    await target.close();
    client.destroy();
    await admin.query(`DROP DATABASE "${dbname}" WITH (FORCE)`);
    await admin.end();
    await rm(dir, { recursive: true, force: true });
  }
});
