import { spawn } from 'node:child_process';
import { once } from 'node:events';
import { fileURLToPath } from 'node:url';
import test from 'node:test';
import assert from 'node:assert/strict';
import { randomUUID, createHmac } from 'node:crypto';
import { createServer } from 'node:http';
import { S3Client, CreateBucketCommand, DeleteObjectCommand } from '@aws-sdk/client-s3';
import { PostgresStore } from '../../src/postgres-store.ts';
import { S3Objects } from '../../src/object-store.ts';
import { WebhookDispatcher } from '../../src/webhooks.ts';
import { createApp } from '../../src/app.ts';
import { fixture, sample } from '../../scripts/formats/fixtures.ts';
const url = process.env.TEST_DATABASE_URL;
if (!url) throw new Error('TEST_DATABASE_URL must name a disposable integration database');
const endpoint = process.env.S3_ENDPOINT ?? 'http://127.0.0.1:18400',
  bucket = `imprint-test-${randomUUID()}`;
const client = new S3Client({ region: 'us-east-1', endpoint, forcePathStyle: true });
await client.send(new CreateBucketCommand({ Bucket: bucket }));
let now = Date.now();
const open = () => new PostgresStore(url!, new S3Objects({ bucket, endpoint }), { now: () => now });
test('PostgreSQL/S3: tenants, versions, concurrent claims, leases, durable webhooks, cleanup, independent API/worker', async () => {
  const first = open(),
    second = open();
  await first.initialize();
  await second.initialize();
  const owner = await first.createTenant('Cluster example'),
    other = await first.createTenant('Isolated tenant');
  try {
    assert.equal(await second.authenticate(owner.apiKey), owner.id);
    const bytes = await fixture('odt', 1),
      uploaded = await first.upload(owner.id, 'Example', bytes, 'odt');
    await first.publish(owner.id, uploaded.templateId, uploaded.versionId);
    await assert.rejects(second.template(other.id, uploaded.templateId), { code: 'NOT_FOUND' });
    const input = { templateId: uploaded.templateId, data: sample(1) },
      one = await first.enqueue(owner.id, input, 'same');
    assert.equal((await second.enqueue(owner.id, input, 'same')).job.id, one.job.id);
    await assert.rejects(second.enqueue(owner.id, { ...input, data: { different: true } }, 'same'), {
      code: 'IDEMPOTENCY_CONFLICT',
    });
    const two = await second.enqueue(owner.id, input),
      claims = await Promise.all([first.claim(), second.claim()]);
    assert(claims[0] && claims[1]);
    assert.notEqual(claims[0].id, claims[1].id);
    assert(claims[0].template.equals(bytes));
    // Expired attempts are reclaimed and fenced from publishing stale results.
    now += 151000;
    const reclaimed = await first.claim();
    assert(reclaimed);
    assert.equal(reclaimed.attempts, 2);
    await second.succeed(
      claims.find((j) => j!.id === reclaimed.id)!,
      { bytes: Buffer.from('stale'), format: 'odt' },
    );
    assert.equal((await first.job(owner.id, reclaimed.id)).status, 'running');
    await first.succeed(reclaimed, { bytes, format: 'odt' });
    assert((await second.result(owner.id, reclaimed.id)).bytes.equals(bytes));
    await assert.rejects(second.result(other.id, reclaimed.id), { code: 'NOT_FOUND' });
    const remaining = await second.claim();
    assert(remaining);
    await second.fail(remaining, 'TEST_FAILURE', false);
    // Retry the exact event ID and payload, verifying its HMAC at the receiving application.
    const received: { id: string; body: string }[] = [];
    let reject = true;
    const secret = 'signed-webhook-integration-secret-32';
    const server = createServer(async (req, res) => {
      const chunks: Buffer[] = [];
      for await (const part of req) chunks.push(Buffer.from(part));
      const body = Buffer.concat(chunks).toString();
      assert.equal(
        req.headers['x-imprint-signature'],
        `sha256=${createHmac('sha256', secret).update(`${req.headers['x-imprint-timestamp']}.${body}`).digest('hex')}`,
      );
      received.push({ id: String(req.headers['x-imprint-event-id']), body });
      res.writeHead(reject ? 503 : 204);
      res.end();
    });
    await new Promise<void>((resolve) => server.listen(0, '127.0.0.1', resolve));
    try {
      const address = server.address() as { port: number },
        dispatcher = new WebhookDispatcher(first, {
          url: `http://127.0.0.1:${address.port}`,
          secret,
          allowHttp: true,
        });
      await dispatcher.tick();
      reject = false;
      now += 5000;
      await dispatcher.tick();
      assert.equal(received.length, 2);
      assert.deepEqual(received[0], received[1]);
      assert(!received[0]!.body.includes('customer'));
      await dispatcher.tick();
      assert.equal(
        Number(
          (await first.pool.query('SELECT COUNT(*) AS n FROM webhooks WHERE delivered_at IS NOT NULL'))
            .rows[0].n,
        ),
        2,
      );
    } finally {
      await new Promise<void>((resolve) => server.close(() => resolve()));
    }
    // Exercise the actual independent worker entry point in another OS process.
    now = Date.now();
    const api = createApp({ apiKey: 'cluster-admin-key-with-32-characters', store: second, runQueue: false });
    const worker = spawn(process.execPath, [fileURLToPath(new URL('../../src/worker.ts', import.meta.url))], {
      env: {
        ...process.env,
        DATABASE_URL: url!,
        S3_ENDPOINT: endpoint,
        S3_BUCKET: bucket,
        IMPRINT_DATA_DIR: '',
        WEBHOOK_URL: '',
      },
      stdio: ['ignore', 'pipe', 'pipe'],
    });
    let workerErrors = '';
    worker.stderr.on('data', (chunk) => {
      workerErrors += String(chunk).slice(0, 10000);
    });
    const exited = once(worker, 'exit');
    try {
      const headers = { authorization: `Bearer ${owner.apiKey}` };
      const submitted = await api.inject({ method: 'POST', url: '/v1/jobs', headers, payload: input });
      assert.equal(submitted.statusCode, 202, submitted.body);
      let state;
      for (let n = 0; n < 100; n++) {
        state = (await api.inject({ url: `/v1/jobs/${submitted.json().id}`, headers })).json();
        if (state.status === 'succeeded') break;
        await new Promise((r) => setTimeout(r, 100));
      }
      assert.equal(state.status, 'succeeded');
      const download = await api.inject({ url: state.downloadUrl, headers });
      assert.equal(download.statusCode, 200);
      assert(download.rawPayload.subarray(0, 2).equals(Buffer.from('PK')));
      assert.equal(workerErrors, '');
    } finally {
      worker.kill('SIGTERM');
      const result = await Promise.race([
        exited,
        new Promise((_, reject) => {
          const timer = setTimeout(() => {
            worker.kill('SIGKILL');
            reject(new Error('Worker failed to drain'));
          }, 10000);
          timer.unref();
        }),
      ]);
      assert.equal((result as unknown[])[0], 0);
      await api.close();
    }
    const key = (await first.pool.query('SELECT result_key FROM jobs WHERE id=$1', [reclaimed.id])).rows[0]
      .result_key;
    await client.send(new DeleteObjectCommand({ Bucket: bucket, Key: `imprint/${key}` }));
    await assert.rejects(first.result(owner.id, reclaimed.id), { code: 'OBJECT_UNAVAILABLE' });
    now += 8 * 86400000;
    await first.cleanup();
    assert.equal((await first.pool.query('SELECT COUNT(*) AS n FROM jobs')).rows[0].n, '0');
    await first.deleteTemplate(owner.id, uploaded.templateId);
    now += 3600001;
    await first.cleanup();
    assert.equal((await first.pool.query('SELECT COUNT(*) AS n FROM objects')).rows[0].n, '0');
  } finally {
    await first.close();
    await second.close().catch(() => {});
    client.destroy();
  }
});
