import test from 'node:test';
import assert from 'node:assert/strict';
import { validateDelivery, type ArtifactValidator } from '../src/delivery.ts';
import { reviewTemplateCandidate } from '../src/template-review.ts';
import { runRegressionSuite } from '../src/regression.ts';
import { encodeTextTemplate } from '../src/text-template.ts';
import { createReportSnapshot, reportSnapshotHtml, renderReportSnapshot } from '../src/report-view.ts';

test('Delivery rejects coerced non-string validator identities', async () => {
  for (const id of [undefined, null, 123, { toString: () => 'plausible' }]) {
    const validator = {
      id,
      version: '1',
      async validate() {
        return { passed: true, details: 'ok' };
      },
    } as unknown as ArtifactValidator;
    await assert.rejects(validateDelivery(Buffer.from('artifact'), [validator]), /identity/);
  }
});
test('Review rejects accessors in options without invoking them', async () => {
  const template = encodeTextTemplate('txt', '{d.name}');
  let accessed = false;
  const options = Object.defineProperty({}, 'locale', {
    enumerable: true,
    get() {
      accessed = true;
      return 'en-US';
    },
  });
  await assert.rejects(
    reviewTemplateCandidate(template, template, [
      { name: 'case', data: { name: 'Expected' }, approved: Buffer.from('Expected'), options },
    ]),
  );
  assert.equal(accessed, false);
});
test('Direct regression suite snapshots approved cases before yielding', async () => {
  const template = encodeTextTemplate('txt', '{d.name}');
  const data = { name: 'Expected' };
  const pending = runRegressionSuite(template, [
    { name: 'immutable', data, approved: Buffer.from('Expected') },
  ]);
  data.name = 'Changed after call';
  assert.equal((await pending).passed, true);
});
test('Report snapshot HTML escapes script terminators and exports same filtered rows', async () => {
  const snapshot = createReportSnapshot(
    {
      title: '</script><script>alert(1)</script>',
      columns: [
        { key: 'name', label: '<img src=x onerror=alert(1)>', type: 'string' },
        { key: 'amount', label: 'Amount', type: 'number', sum: true },
      ],
      rows: [
        { name: '</script>\u2028<script>alert(1)</script>', amount: 2 },
        { name: 'excluded', amount: 3 },
      ],
    },
    [{ key: 'amount', op: 'lte', value: 2 }],
  );
  const html = reportSnapshotHtml(snapshot);
  assert(!html.includes('<script>alert(1)</script>'));
  assert(!html.includes('<img src=x'));
  const text = await renderReportSnapshot(
    encodeTextTemplate(
      'txt',
      '{d.report.rows[i].name}\n{d.report.rows[i+1]}\nTotal {d.report.totals.amount}',
    ),
    snapshot,
  );
  assert(!text.toString().includes('excluded'));
  assert(text.toString().includes('Total 2'));
});
