import test from 'node:test';
import assert from 'node:assert/strict';
import { createApp } from '../src/app.ts';
import { encodeTextTemplate } from '../src/text-template.ts';
import { PDFDocument } from 'pdf-lib';
const apiKey = 'assurance-api-tests-secret-0123456789';
const headers = { authorization: `Bearer ${apiKey}` };
const template = encodeTextTemplate('txt', '{d.name}').toString('base64');
const b64 = (text: string) => Buffer.from(text).toString('base64');
const definition = {
  title: 'Example',
  columns: [
    { key: 'name', label: 'Name', type: 'string' },
    { key: 'amount', label: 'Amount', type: 'number', sum: true },
  ],
  rows: [
    { name: 'A', amount: 1 },
    { name: 'B', amount: 2 },
  ],
};

test('Assurance HTTP operations require auth and reject malformed and extra properties', async () => {
  const app = createApp({ apiKey });
  try {
    assert.equal(
      (await app.inject({ method: 'POST', url: '/v1/assurance', payload: { operation: 'contract' } }))
        .statusCode,
      401,
    );
    for (const payload of [
      { operation: 'unknown' },
      { operation: 'report', definition, extra: true },
      { operation: 'report', definition, filters: null },
      { operation: 'contract', schema: { type: 'string' }, data: 'A', assertions: null },
      { operation: 'contract', schema: { type: 'object' } },
      { operation: 'trace', template: 'AB==', data: {} },
      { operation: 'trace', template, data: {}, options: { script: 'run-me' } },
      { operation: 'trace', template, data: {}, includeValues: 'true' },
      { operation: 'diff', before: 'YQ==\n', after: b64('a'), format: 'txt' },
      { operation: 'delivery', artifact: b64('x'), validator: 'arbitrary-code' },
    ]) {
      const response = await app.inject({ method: 'POST', url: '/v1/assurance', headers, payload });
      assert.equal(response.statusCode, 422, response.body);
    }
    const over = await app.inject({
      method: 'POST',
      url: '/v1/assurance',
      headers,
      payload: { operation: 'report', filler: 'x'.repeat(12 * 1024 * 1024) },
    });
    assert.equal(over.statusCode, 413);
  } finally {
    await app.close();
  }
});

test('Assurance HTTP contract, execution trace, regression, diff and report work end to end', async () => {
  const app = createApp({ apiKey });
  const call = async (payload: object) => {
    const response = await app.inject({ method: 'POST', url: '/v1/assurance', headers, payload });
    assert.equal(response.statusCode, 200, response.body);
    assert.match(response.headers['content-type']!, /application\/json/);
    assert.equal(response.headers['cache-control'], 'no-store');
    return response.json();
  };
  try {
    const contract = await call({
      operation: 'contract',
      schema: {
        type: 'object',
        properties: { name: { type: 'string' } },
        required: ['name'],
        additionalProperties: false,
      },
      data: { name: 'A' },
    });
    assert.equal(contract.validation.valid, true);
    assert.match(contract.typescript, /name/);
    const invalid = await call({ operation: 'contract', schema: { type: 'string' }, data: 12 });
    assert.equal(invalid.validation.valid, false);
    const trace = await call({ operation: 'trace', template, data: { name: 'Sensitive name' } });
    assert.equal(trace.ok, true);
    assert.equal(Buffer.from(trace.artifact, 'base64').toString(), 'Sensitive name');
    assert.equal(trace.trace.phase, 'execution');
    assert(!JSON.stringify(trace.trace).includes('Sensitive name'));
    const suite = await call({
      operation: 'regression',
      template,
      cases: [
        { name: 'baseline', data: { name: 'A' }, approved: b64('A') },
        { name: 'empty', data: { name: '' }, approved: '' },
      ],
    });
    assert.equal(suite.passed, true);
    const changed = await call({
      operation: 'regression',
      template,
      cases: [{ name: 'baseline', data: { name: 'B' }, approved: b64('A') }],
    });
    assert.equal(changed.passed, false);
    const diff = await call({ operation: 'diff', before: '', after: b64('B'), format: 'txt' });
    assert.equal(diff.equal, false);
    const report = await call({
      operation: 'report',
      definition,
      filters: [{ key: 'amount', op: 'gte', value: 2 }],
    });
    assert.equal(report.snapshot.rows.length, 1);
    assert.equal(report.snapshot.totals.amount, 2);
    assert.match(report.html, /Download selected CSV/);
  } finally {
    await app.close();
  }
});

test('Assurance HTTP migrations, component impact, reviewed artifacts and replay round trip', async () => {
  const app = createApp({ apiKey });
  const call = async (payload: object) => {
    const response = await app.inject({ method: 'POST', url: '/v1/assurance', headers, payload });
    assert.equal(response.statusCode, 200, response.body);
    return response.json();
  };
  try {
    const migration = await call({
      operation: 'migrate',
      sourceType: 'carbone-text',
      source: '{d.name:upperCase}',
      stringFields: ['name'],
      locale: 'en-US',
      data: { name: 'alice' },
    });
    assert.equal(Buffer.from(migration.artifact, 'base64').toString(), 'ALICE');
    const unsafe = await app.inject({
      method: 'POST',
      url: '/v1/assurance',
      headers,
      payload: {
        operation: 'migrate',
        sourceType: 'carbone-text',
        source: '{d.name:upperCase}',
        stringFields: ['name'],
        locale: 'en-US',
        data: { name: 12 },
      },
    });
    assert.equal(unsafe.statusCode, 422);
    const expansion = await call({
      operation: 'components',
      components: [{ id: 'address', version: '1', source: 'Street' }],
      source: '{{>address@1}}',
      impact: { id: 'address', version: '1', templates: { invoice: '{{>address@1}}', other: 'Text' } },
    });
    assert.equal(expansion.source, 'Street');
    assert.deepEqual(expansion.impactedTemplates, ['invoice']);
    const review = await call({
      operation: 'review',
      base: template,
      candidate: template,
      cases: [{ name: 'baseline', data: { name: 'A' }, approved: b64('A') }],
    });
    assert.equal(review.review.passed, true);
    assert.equal(review.artifact, template);
    const rejected = await call({
      operation: 'review',
      base: template,
      candidate: template,
      cases: [{ name: 'baseline', data: { name: 'B' }, approved: b64('A') }],
    });
    assert.equal(rejected.review.passed, false);
    assert.equal(rejected.artifact, undefined);
    const bundle = await call({ operation: 'replay-create', template, data: { name: 'Replay' } });
    const replay = await call({ operation: 'replay', bundle });
    assert.equal(Buffer.from(replay.artifact, 'base64').toString(), 'Replay');
    bundle.payload.data.name = 'Tampered';
    const bad = await app.inject({
      method: 'POST',
      url: '/v1/assurance',
      headers,
      payload: { operation: 'replay', bundle },
    });
    assert.equal(bad.statusCode, 422);
  } finally {
    await app.close();
  }
});

test('Assurance HTTP PDF structural delivery quarantines invalid artifacts', async () => {
  const app = createApp({ apiKey });
  try {
    const pdf = await PDFDocument.create();
    pdf.addPage();
    const valid = await app.inject({
      method: 'POST',
      url: '/v1/assurance',
      headers,
      payload: { operation: 'delivery', artifact: Buffer.from(await pdf.save()).toString('base64') },
    });
    assert.equal(valid.statusCode, 200, valid.body);
    assert.equal(valid.json().evidence.status, 'approved');
    const invalid = await app.inject({
      method: 'POST',
      url: '/v1/assurance',
      headers,
      payload: { operation: 'delivery', artifact: b64('Not a PDF') },
    });
    assert.equal(invalid.statusCode, 200, invalid.body);
    assert.equal(invalid.json().evidence.status, 'quarantined');
    assert.equal(invalid.json().artifact, undefined);
  } finally {
    await app.close();
  }
});

test('Assurance HTTP shares renderer capacity and releases it after completion', async () => {
  const app = createApp({ apiKey, concurrency: 1 });
  try {
    const request = {
      method: 'POST' as const,
      url: '/v1/assurance',
      headers,
      payload: { operation: 'report', definition },
    };
    const responses = await Promise.all([app.inject(request), app.inject(request), app.inject(request)]);
    assert.equal(responses.filter((r) => r.statusCode === 200).length, 1);
    const busy = responses.filter((r) => r.statusCode === 503);
    assert.equal(busy.length, 2);
    assert(busy.every((r) => r.headers['retry-after'] === '1'));
    assert.equal((await app.inject(request)).statusCode, 200);
  } finally {
    await app.close();
  }
});
