import test from 'node:test';
import assert from 'node:assert/strict';
import { mkdtemp, readFile, rm } from 'node:fs/promises';
import { tmpdir } from 'node:os';
import { join } from 'node:path';
import { StandaloneStore, LEASE_MS } from '../src/store.ts';
import { ServiceError } from '../src/service-error.ts';
import { createApp } from '../src/app.ts';
import { RenderService } from '../src/render-service.ts';
import { DurableQueue } from '../src/queue.ts';
import { RenderError } from '../src/errors.ts';

const key = 'standalone-test-administrator-key';
const template = await readFile(new URL('../corpus/docx/level-001/template.docx', import.meta.url));
const data = JSON.parse(await readFile(new URL('../corpus/docx/level-001/input.json', import.meta.url), 'utf8'));
const code = (value: string) => (error: unknown) => error instanceof ServiceError && error.code === value;
async function fixture(run: (directory: string) => Promise<void>) {
  const directory = await mkdtemp(join(tmpdir(), 'imprint-service-'));
  try { await run(directory); } finally { await rm(directory, { recursive: true, force: true }); }
}
function version(store: StandaloneStore, tenant = 'default') {
  const uploaded = store.upload(tenant, 'Invoice', template, 'docx');
  store.publish(tenant, uploaded.templateId, uploaded.versionId);
  return uploaded;
}

test('store pins a version at submission and supports publish/rollback without changing queued jobs', async () => fixture(async directory => {
  const store = new StandaloneStore(directory);
  try {
    const first = version(store);
    const job = store.enqueue('default', { templateId: first.templateId, data }).job;
    const second = store.upload('default', '', template, 'docx', first.templateId);
    store.publish('default', first.templateId, second.versionId);
    assert.equal(store.job('default', job.id).version_id, first.versionId);
    assert.equal(store.enqueue('default', { templateId: first.templateId, data }).job.version_id, second.versionId);
    store.publish('default', first.templateId, first.versionId);
    assert.equal(store.versions('default', first.templateId).length, 2);
    assert.throws(() => store.deleteTemplate('default', first.templateId), code('TEMPLATE_IN_USE'));
  } finally { store.close(); }
}));

test('store commits templates and queued jobs across reopening; expired leases recover and reject stale completion', async () => fixture(async directory => {
  let now = Date.now();
  let store = new StandaloneStore(directory, { now: () => now });
  const first = version(store), submitted = store.enqueue('default', { templateId: first.templateId, data }).job;
  const old = store.claim()!;
  assert.equal(old.attempts, 1);
  store.close();
  store = new StandaloneStore(directory, { now: () => now });
  try {
    assert.equal(store.job('default', submitted.id).status, 'running');
    assert.equal(store.claim(), undefined);
    now += LEASE_MS + 1;
    const recovered = store.claim()!;
    assert.equal(recovered.id, old.id); assert.equal(recovered.attempts, 2);
    store.succeed(old, { bytes: Buffer.from('stale'), format: 'docx' });
    assert.equal(store.job('default', old.id).status, 'running');
    store.succeed(recovered, { bytes: template, format: 'docx' });
    assert.deepEqual(store.result('default', old.id).bytes, template);
    store.close(); store = new StandaloneStore(directory);
    assert.deepEqual(store.result('default', old.id).bytes, template);
  } finally { store.close(); }
}));

test('idempotency is tenant scoped and a conflicting body cannot silently reuse an old result', async () => fixture(async directory => {
  const store = new StandaloneStore(directory);
  try {
    const first = version(store), input = { templateId: first.templateId, data };
    const job = store.enqueue('default', input, 'order-42');
    assert.equal(job.reused, false);
    assert.equal(store.enqueue('default', input, 'order-42').job.id, job.job.id);
    assert.equal(store.enqueue('default', input, 'order-42').reused, true);
    assert.throws(() => store.enqueue('default', { ...input, output: 'pdf' }, 'order-42'), code('IDEMPOTENCY_CONFLICT'));
    const other = store.createTenant('Other'), second = version(store, other.id);
    assert.equal(store.enqueue(other.id, { templateId: second.templateId, data }, 'order-42').reused, false);
  } finally { store.close(); }
}));

test('tenant boundaries protect versions, publication, job status, downloads and deletion', async () => fixture(async directory => {
  const store = new StandaloneStore(directory);
  try {
    const other = store.createTenant('Other'), first = version(store);
    const job = store.enqueue('default', { templateId: first.templateId, data }).job;
    for (const run of [
      () => store.template(other.id, first.templateId), () => store.versions(other.id, first.templateId),
      () => store.publish(other.id, first.templateId, first.versionId), () => store.deleteTemplate(other.id, first.templateId),
      () => store.upload(other.id, '', template, 'docx', first.templateId),
      () => store.enqueue(other.id, { templateId: first.templateId, data }),
      () => store.job(other.id, job.id), () => store.result(other.id, job.id), () => store.deleteJob(other.id, job.id),
    ]) assert.throws(run, code('NOT_FOUND'));
  } finally { store.close(); }
}));

test('queued work requires a published or explicitly selected owned version and matching output format', async () => fixture(async directory => {
  const store = new StandaloneStore(directory);
  try {
    const first = store.upload('default', 'Draft', template, 'docx'), second = version(store);
    assert.throws(() => store.enqueue('default', { templateId: first.templateId, data }), code('TEMPLATE_NOT_PUBLISHED'));
    assert.throws(() => store.enqueue('default', { templateId: first.templateId, versionId: second.versionId, data }), code('NOT_FOUND'));
    assert.throws(() => store.enqueue('default', { templateId: second.templateId, data, output: 'xlsx' }), code('OUTPUT_FORMAT_MISMATCH'));
    assert.equal(store.enqueue('default', { ...first, data }).job.version_id, first.versionId);
  } finally { store.close(); }
}));

test('transient failures retry at most twice; deterministic failures, queue deadlines and interrupted final attempts are terminal', async () => fixture(async directory => {
  let now = Date.now(); const store = new StandaloneStore(directory, { now: () => now });
  try {
    const first = version(store), input = { templateId: first.templateId, data };
    const job = store.enqueue('default', input).job;
    store.fail(store.claim()!, 'CONVERSION_FAILED', true);
    assert.equal(store.claim(), undefined); now += 2001;
    const retry = store.claim()!; assert.equal(retry.attempts, 2);
    store.fail(retry, 'CONVERSION_FAILED', true);
    assert.equal(store.job('default', job.id).status, 'failed');
    const bad = store.enqueue('default', input).job;
    store.fail(store.claim()!, 'MISSING_DATA', false);
    assert.equal(store.job('default', bad.id).attempts, 1);
    const expired = store.enqueue('default', input).job;
    now += 86_400_001; assert.equal(store.claim(), undefined);
    assert.equal(store.job('default', expired.id).error, 'JOB_EXPIRED');
    const killed = store.enqueue('default', input).job;
    store.claim(); now += LEASE_MS + 1; store.claim(); now += LEASE_MS + 1;
    assert.equal(store.claim(), undefined); assert.equal(store.job('default', killed.id).status, 'failed');
  } finally { store.close(); }
}));

test('quota reserves output space before accepting jobs, bounds queue size and releases capacity on deletion', async () => fixture(async directory => {
  const store = new StandaloneStore(directory, { tenantQuota: template.length + 33 * 1024 * 1024 });
  try {
    const first = version(store), input = { templateId: first.templateId, data };
    const job = store.enqueue('default', input).job;
    assert.throws(() => store.enqueue('default', input), code('STORAGE_QUOTA'));
    assert.throws(() => store.deleteJob('default', job.id), code('JOB_ACTIVE'));
    store.succeed(store.claim()!, { bytes: template, format: 'docx' });
    store.deleteJob('default', job.id); store.deleteTemplate('default', first.templateId);
    assert.deepEqual(store.listTemplates('default'), []);
  } finally { store.close(); }
}));

test('terminal job retention removes inputs/results and permits deleting the template', async () => fixture(async directory => {
  let now = Date.now(); const store = new StandaloneStore(directory, { now: () => now });
  try {
    const first = version(store), job = store.enqueue('default', { templateId: first.templateId, data }).job;
    store.succeed(store.claim()!, { bytes: template, format: 'docx' });
    assert.deepEqual(store.db.prepare('SELECT payload FROM jobs WHERE id=?').get(job.id), { payload: '{}' });
    now += 7 * 86_400_000 + 1; store.cleanup();
    assert.throws(() => store.result('default', job.id), code('NOT_FOUND'));
    store.deleteTemplate('default', first.templateId);
  } finally { store.close(); }
}));

test('rotated and revoked keys immediately lose access, while key hashes persist across restart', async () => fixture(async directory => {
  let store = new StandaloneStore(directory);
  const tenant = store.createTenant('Example application');
  assert.equal(store.authenticate(tenant.apiKey), tenant.id);
  assert(!JSON.stringify(store.db.prepare('SELECT * FROM tenants').all()).includes(tenant.apiKey));
  store.close(); store = new StandaloneStore(directory);
  try {
    assert.equal(store.authenticate(tenant.apiKey), tenant.id);
    const rotated = store.rotateKey(tenant.id);
    assert.equal(store.authenticate(tenant.apiKey), undefined);
    assert.equal(store.authenticate(rotated.apiKey!), tenant.id);
    store.rotateKey(tenant.id, true); assert.equal(store.authenticate(rotated.apiKey!), undefined);
  } finally { store.close(); }
}));

test('HTTP service validates uploads, enforces admin boundaries and renders a persistent queued job after API restart', async () => fixture(async directory => {
  let app = createApp({ apiKey: key, dataDir: directory, runQueue: false });
  const call = (method: 'GET' | 'POST' | 'DELETE', url: string, payload?: object, apiKey = key) => app.inject({ method, url, payload, headers: { authorization: `Bearer ${apiKey}` } });
  try {
    assert.equal((await app.inject({ method: 'POST', url: '/v1/tenants', payload: { name: 'Example application' } })).statusCode, 401);
    const owner = (await call('POST', '/v1/tenants', { name: 'Example application' })).json();
    assert.equal((await call('POST', '/v1/tenants', { name: 'Escalation' }, owner.apiKey)).statusCode, 403);
    assert.equal((await call('POST', '/v1/templates', { name: 'Broken', template: Buffer.from('not a zip').toString('base64') }, owner.apiKey)).statusCode, 422);
    const upload = await call('POST', '/v1/templates', { name: 'Invoice', template: template.toString('base64') }, owner.apiKey);
    assert.equal(upload.statusCode, 201, upload.body);
    const first = upload.json();
    assert.equal((await call('GET', `/v1/templates/${first.templateId}`)).statusCode, 404);
    assert.equal((await call('POST', `/v1/templates/${first.templateId}/publish`, { versionId: first.versionId }, owner.apiKey)).statusCode, 200);
    const metadata = (await call('GET', `/v1/templates/${first.templateId}`, undefined, owner.apiKey)).json();
    assert.equal(metadata.publishedVersion, first.versionId); assert.equal(metadata.versions[0].templateId, first.templateId);
    assert.equal(metadata.published_version, undefined);
    const response = await call('POST', '/v1/jobs', { templateId: first.templateId, data }, owner.apiKey);
    assert.equal(response.statusCode, 202, response.body); const job = response.json();
    assert.equal((await call('GET', `/v1/jobs/${job.id}/result`, undefined, owner.apiKey)).statusCode, 409);
    await app.close();
    app = createApp({ apiKey: key, dataDir: directory });
    let status;
    for (let n = 0; n < 100; n++) {
      status = (await call('GET', `/v1/jobs/${job.id}`, undefined, owner.apiKey)).json();
      if (['succeeded', 'failed'].includes(status.status)) break;
      await new Promise(resolve => setTimeout(resolve, 50));
    }
    assert.equal(status.status, 'succeeded', JSON.stringify(status));
    assert.equal((await call('GET', status.downloadUrl)).statusCode, 404);
    const result = await call('GET', status.downloadUrl, undefined, owner.apiKey);
    assert.equal(result.statusCode, 200); assert.equal(result.rawPayload.subarray(0, 2).toString(), 'PK');
    assert.equal(result.headers['cache-control'], 'no-store');
    assert.equal((await call('DELETE', `/v1/jobs/${job.id}`, undefined, owner.apiKey)).statusCode, 204);
    assert.equal((await call('DELETE', `/v1/templates/${first.templateId}`, undefined, owner.apiKey)).statusCode, 204);
  } finally { await app.close(); }
}));

test('queue drains in-flight work on shutdown, preserving real render errors without retries', async () => fixture(async directory => {
  const store = new StandaloneStore(directory), first = version(store);
  const job = store.enqueue('default', { templateId: first.templateId, data: { title: {} } }).job;
  const errors: unknown[] = [], queue = new DurableQueue(store, new RenderService(), error => errors.push(error));
  try {
    await queue.tick(); await queue.stop();
    assert.deepEqual(errors, []);
    const result = store.job('default', job.id);
    assert.equal(result.status, 'failed'); assert.equal(result.error, 'INVALID_DATA'); assert.equal(result.attempts, 1);
  } finally { store.close(); }
}));

test('synchronous and queued work share one admission limit', async () => {
  const renderer = new RenderService(1);
  const rendering = renderer.render(template, { data });
  await assert.rejects(renderer.render(template, { data }), code('BUSY'));
  await rendering;
  assert.equal(renderer.available, true);
  await assert.rejects(renderer.render(template, { data: { title: {} } }), error => error instanceof RenderError && error.code === 'INVALID_DATA');
  assert.equal(renderer.available, true);
});

test('independent database connections cannot claim the same active job', async () => fixture(async directory => {
  const a = new StandaloneStore(directory), b = new StandaloneStore(directory);
  try {
    const first = version(a), job = a.enqueue('default', { templateId: first.templateId, data }).job;
    const claimed = a.claim()!;
    assert.equal(claimed.id, job.id); assert.equal(b.claim(), undefined);
    b.fail({ ...claimed, lease: 'stale-token' }, 'CONVERSION_FAILED', true);
    assert.equal(a.job('default', job.id).status, 'running');
    a.succeed(claimed, { bytes: template, format: 'docx' });
    assert.equal(b.job('default', job.id).status, 'succeeded');
  } finally { b.close(); a.close(); }
}));

test('global quota and per-tenant queue bounds reject work before acceptance', async () => fixture(async directory => {
  const store = new StandaloneStore(directory, { tenantQuota: 8 * 1024 ** 3, totalQuota: 16 * 1024 ** 3 });
  try {
    const first = version(store), input = { templateId: first.templateId, data };
    for (let n = 0; n < 100; n++) store.enqueue('default', input);
    assert.throws(() => store.enqueue('default', input), code('QUEUE_FULL'));
    assert.equal((store.db.prepare('SELECT COUNT(*) AS n FROM jobs').get() as { n: number }).n, 100);
  } finally { store.close(); }
  const limited = new StandaloneStore(directory, { totalQuota: 1 });
  try { assert.throws(() => version(limited, limited.createTenant('Other').id), code('STORAGE_QUOTA')); }
  finally { limited.close(); }
}));

test('version uploads reserve admission across asynchronous ownership checks and object writes', async () => fixture(async directory => {
  const store = new StandaloneStore(directory), first = version(store);
  let releaseLookup!: () => void, enteredLookup!: () => void, releaseWrite!: () => void, enteredWrite!: () => void;
  const lookupGate = new Promise<void>(resolve => { releaseLookup = resolve; });
  const lookupStarted = new Promise<void>(resolve => { enteredLookup = resolve; });
  const writeGate = new Promise<void>(resolve => { releaseWrite = resolve; });
  const writeStarted = new Promise<void>(resolve => { enteredWrite = resolve; });
  const adapter = new Proxy(store, { get(target, property) {
    if (property === 'template') return async (...args: Parameters<StandaloneStore['template']>) => { enteredLookup(); await lookupGate; return target.template(...args); };
    if (property === 'upload') return async (...args: Parameters<StandaloneStore['upload']>) => { enteredWrite(); await writeGate; return target.upload(...args); };
    const value = Reflect.get(target, property); return typeof value === 'function' ? value.bind(target) : value;
  } }) as unknown as import('../src/store-interface.ts').ServiceStore;
  const app = createApp({ apiKey: key, store: adapter, concurrency: 1, runQueue: false });
  const request = { method: 'POST' as const, url: `/v1/templates/${first.templateId}/versions`, headers: { authorization: `Bearer ${key}` }, payload: { template: template.toString('base64') } };
  try {
    const pending = app.inject(request).then(response => response);
    await lookupStarted; assert.equal((await app.inject(request)).statusCode, 503);
    releaseLookup(); await writeStarted; assert.equal((await app.inject(request)).statusCode, 503);
    releaseWrite(); assert.equal((await pending).statusCode, 201);
  } finally { releaseLookup(); releaseWrite(); await app.close(); }
}));

test('asynchronous queue claims reserve the same slots as direct rendering', async () => fixture(async directory => {
  const store = new StandaloneStore(directory), first = version(store);
  const submitted = store.enqueue('default', { templateId: first.templateId, data }).job;
  let release!: () => void, entered!: () => void;
  const gate = new Promise<void>(resolve => { release = resolve; }), started = new Promise<void>(resolve => { entered = resolve; });
  const adapter = new Proxy(store, { get(target, property) {
    if (property === 'claim') return async () => { entered(); await gate; return target.claim(); };
    const value = Reflect.get(target, property); return typeof value === 'function' ? value.bind(target) : value;
  } }) as unknown as import('../src/store-interface.ts').ServiceStore;
  const errors: unknown[] = [], renderer = new RenderService(1), queue = new DurableQueue(adapter, renderer, error => errors.push(error));
  try {
    const ticking = queue.tick(); await started;
    await assert.rejects(renderer.render(template, { data }), code('BUSY'));
    release(); await ticking; await queue.stop();
    assert.equal(store.job('default', submitted.id).status, 'succeeded'); assert.deepEqual(errors, []);
  } finally { release(); await queue.stop(); store.close(); }
}));
