import { test } from 'node:test';
import assert from 'node:assert/strict';
import { PDFDocument } from 'pdf-lib';
import { validateDelivery, pdfStructureValidator, artifactDigest } from '../src/delivery.ts';
import {
  createReportSnapshot,
  verifyReportSnapshot,
  reportSnapshotHtml,
  renderReportSnapshot,
} from '../src/report-view.ts';
import { encodeTextTemplate } from '../src/text-template.ts';
const definition = {
  title: 'Sales',
  columns: [
    { key: 'name', label: 'Name', type: 'string' as const },
    { key: 'amount', label: 'Amount', type: 'number' as const, sum: true },
  ],
  rows: [
    { name: 'A', amount: 1 },
    { name: 'B', amount: 2 },
    { name: 'C', amount: null },
  ],
};
test('delivery requires explicit policy, quarantines false, error, timeout and invalid results', async () => {
  const bytes = Buffer.from('test');
  await assert.rejects(validateDelivery(bytes, []));
  for (const validate of [
    async () => ({ passed: false, details: 'bad' }),
    async () => {
      throw new Error('secret');
    },
    async () => ({ passed: 'yes', details: 'bad' }) as any,
    () => new Promise<any>(() => {}),
  ]) {
    const result = await validateDelivery(bytes, [{ id: 'test', version: '1', validate }], 5);
    assert.equal(result.evidence.status, 'quarantined');
    assert.equal(result.artifact, undefined);
    assert.equal(result.evidence.sha256, artifactDigest(bytes));
    assert.ok(!JSON.stringify(result).includes('secret'));
  }
});
test('delivery snapshots inputs and isolates validator mutations', async () => {
  const bytes = Buffer.from('safe');
  const validator = {
    id: 'first',
    version: '1',
    async validate(b: Uint8Array) {
      b.fill(0);
      bytes.fill(0);
      validator.id = 'changed';
      return { passed: true, details: 'ok' };
    },
  };
  const result = await validateDelivery(bytes, [
    validator,
    {
      id: 'second',
      version: '1',
      async validate(b) {
        assert.equal(Buffer.from(b).toString(), 'safe');
        return { passed: true, details: 'ok' };
      },
    },
  ]);
  assert.equal(result.artifact?.toString(), 'safe');
  assert.equal(result.evidence.checks[0].id, 'first');
  await assert.rejects(validateDelivery(Buffer.from('x'), [validator, validator]));
});
test('PDF structural validator accepts real PDF and rejects wrong format, malformed and over-limit', async () => {
  const pdf = await PDFDocument.create();
  pdf.addPage();
  pdf.addPage();
  const bytes = await pdf.save();
  assert.equal((await validateDelivery(bytes, [pdfStructureValidator()])).evidence.status, 'approved');
  for (const b of [bytes, Buffer.from('wrong'), Buffer.from('%PDF-broken')])
    assert.equal((await validateDelivery(b, [pdfStructureValidator(1)])).evidence.status, 'quarantined');
});
test('report filters and exports use exactly the selected snapshot with verified totals', async () => {
  const s = createReportSnapshot(definition, [{ key: 'amount', op: 'gte', value: 2 }]);
  assert.equal(s.rows.length, 1);
  assert.equal(s.totals.amount, 2);
  const template = encodeTextTemplate(
    'txt',
    '{d.report.title}\n{d.report.rows[i].name}: {d.report.rows[i].amount}\n{d.report.rows[i+1]}\nTotal: {d.report.totals.amount}',
  );
  const text = (await renderReportSnapshot(template, s)).toString();
  assert.match(text, /B: 2/);
  assert.doesNotMatch(text, /A: 1/);
  assert.match(text, /Total: 2/);
  s.totals.amount = 99;
  assert.throws(() => verifyReportSnapshot(s));
});
test('report nullable values, empty selection, booleans and precision', () => {
  assert.equal(
    createReportSnapshot(definition, [{ key: 'amount', op: 'eq', value: null }]).rows[0].name,
    'C',
  );
  assert.equal(createReportSnapshot(definition, [{ key: 'amount', op: 'gte', value: 5 }]).totals.amount, 0);
  assert.equal(
    createReportSnapshot(definition, [{ key: 'name', op: 'contains', value: 'A' }]).rows.length,
    1,
  );
  const s = createReportSnapshot(
    {
      title: 'Bool',
      columns: [{ key: 'ok', label: 'OK', type: 'boolean' }],
      rows: [{ ok: true }, { ok: false }],
    },
    [{ key: 'ok', op: 'eq', value: false }],
  );
  assert.equal(s.rows.length, 1);
  const decimal = createReportSnapshot({
    ...definition,
    rows: [
      { name: 'a', amount: 0.1 },
      { name: 'b', amount: 0.2 },
    ],
  });
  assert.ok(Math.abs(decimal.totals.amount - 0.3) < 1e-15);
});
test('report rejects invalid definitions, unsafe data, filters, overflow and tampering', () => {
  for (const d of [
    { ...definition, columns: [] },
    { ...definition, columns: [definition.columns[0], definition.columns[0]] },
    { ...definition, rows: [{ name: 'x', amount: '1' }] },
    { ...definition, rows: [{ name: 'x' }] },
    { ...definition, rows: [{ name: 'x', amount: 1, extra: true }] },
    { ...definition, rows: [{ name: 'x', amount: Infinity }] },
    {
      ...definition,
      rows: [
        { name: 'x', amount: Number.MAX_VALUE },
        { name: 'y', amount: Number.MAX_VALUE },
      ],
    },
  ])
    assert.throws(() => createReportSnapshot(d as any));
  for (const f of [
    { key: 'missing', op: 'eq', value: 2 },
    { key: 'name', op: 'gte', value: 'a' },
    { key: 'amount', op: 'eq', value: '2' },
    { key: 'amount', op: 'contains', value: 2 },
    { key: 'amount', op: 'gte', value: null },
  ])
    assert.throws(() => createReportSnapshot(definition, [f as any]));
  const s = createReportSnapshot(definition);
  s.rows[0].name = 'changed';
  assert.throws(() => verifyReportSnapshot(s));
  assert.throws(() => createReportSnapshot(JSON.parse('{"__proto__":{}}')));
});
test('HTML escapes hostile titles/labels/data and embeds a constrained script policy', () => {
  const payload = '</script><script>alert(1)</script><img src=x onerror=alert(2)>';
  const html = reportSnapshotHtml(
    createReportSnapshot({
      title: payload,
      columns: [{ key: 'x', label: payload, type: 'string' }],
      rows: [{ x: payload }],
    }),
  );
  assert.ok(!html.includes(payload));
  assert.ok(html.includes('\\u003c/script'));
  assert.match(html, /Content-Security-Policy/);
  assert.match(html, /script-src 'sha256-/);
  assert.ok(!html.includes('innerHTML'));
});
