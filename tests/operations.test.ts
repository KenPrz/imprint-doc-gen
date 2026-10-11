import test from 'node:test';
import assert from 'node:assert/strict';
import { mkdtemp, rm } from 'node:fs/promises';
import { tmpdir } from 'node:os';
import { join } from 'node:path';
import { StandaloneStore } from '../src/store.ts';
import { backupSqlite, restoreSqlite } from '../src/backup.ts';
import { createApp } from '../src/app.ts';

test('Online backup restores committed WAL data and refuses populated destinations', async () => {
  const directory = await mkdtemp(join(tmpdir(), 'imprint-backup-')),
    source = join(directory, 'source'),
    target = join(directory, 'target'),
    path = join(directory, 'backup.sqlite');
  const store = new StandaloneStore(source);
  try {
    const tenant = store.createTenant('Example');
    await backupSqlite(source, path);
    store.rotateKey(tenant.id);
    await restoreSqlite(path, target);
    const restored = new StandaloneStore(target);
    try {
      assert.equal(restored.authenticate(tenant.apiKey), tenant.id);
    } finally {
      restored.close();
    }
    await assert.rejects(restoreSqlite(path, target), /empty directory/);
    await assert.rejects(backupSqlite(source, path), /already exists/);
  } finally {
    store.close();
    await rm(directory, { recursive: true, force: true });
  }
});
test('Metrics require administrator credentials and do not include request data', async () => {
  const key = 'metrics-integration-test-key',
    app = createApp({ apiKey: key });
  try {
    assert.equal((await app.inject('/metrics')).statusCode, 401);
    await app.inject('/not-a-route-private-name');
    const r = await app.inject({ url: '/metrics', headers: { authorization: `Bearer ${key}` } });
    assert.equal(r.statusCode, 200);
    assert.match(r.body, /imprint_http_requests_total\{status="4xx"\} 2/);
    assert(!r.body.includes('private-name'));
    assert.match(r.body, /imprint_render_capacity 2/);
  } finally {
    await app.close();
  }
});
