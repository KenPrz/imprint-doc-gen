import assert from 'node:assert/strict';
import { mkdir, writeFile } from 'node:fs/promises';
import {
  ComponentRegistry,
  encodeTextTemplate,
  validateContract,
  evaluateAssertions,
  renderWithTrace,
  runRegressionSuite,
  createReplayBundle,
  replayBundle,
  createReportSnapshot,
  reportSnapshotHtml,
  reviewTemplateCandidate,
  acceptTemplateCandidate,
  validateDelivery,
  artifactDigest,
} from '../../src/index.ts';
const output = 'artifacts/assurance/demo';
await mkdir(output, { recursive: true });
const registry = new ComponentRegistry([{ id: 'heading', version: 'v1', source: 'Invoice {d.id}' }]);
const expansion = registry.expand('{{> heading@v1}}\nTotal: {d.total}');
const template = encodeTextTemplate('txt', expansion.source),
  data = { id: 'INV-001', total: 12.5 };
assert.equal(
  validateContract(
    {
      type: 'object',
      additionalProperties: false,
      required: ['id', 'total'],
      properties: { id: { type: 'string', minLength: 1 }, total: { type: 'number', minimum: 0 } },
    },
    data,
  ).valid,
  true,
);
assert.equal(
  evaluateAssertions({ ...data, lines: [{ amount: 10 }, { amount: 2.5 }] }, [
    {
      name: 'balance',
      left: { sum: '/lines', field: '/amount' },
      operator: 'equals',
      right: { path: '/total' },
    },
  ]).valid,
  true,
);
// Approved expectation is authored independently, not generated from the candidate.
const approved = Buffer.from('Invoice INV-001\nTotal: 12.5');
const cases = [{ name: 'balanced invoice', data, approved }];
const regression = await runRegressionSuite(template, cases);
assert.equal(regression.passed, true);
const rendered = await renderWithTrace(template, data);
if (!rendered.ok) throw new Error(rendered.error.code);
const bundle = await createReplayBundle(template, data);
assert.deepEqual(await replayBundle(bundle), approved);
const review = await reviewTemplateCandidate(template, template, cases);
assert.deepEqual(acceptTemplateCandidate(review, template), template);
const delivery = await validateDelivery(rendered.value, [
  {
    id: 'approved-invoice',
    version: '1',
    async validate(bytes) {
      return {
        passed: artifactDigest(bytes) === artifactDigest(approved),
        details: 'Exact customer fixture',
      };
    },
  },
]);
assert.equal(delivery.evidence.status, 'approved');
const snapshot = createReportSnapshot({
  title: 'Invoice lines',
  columns: [
    { key: 'description', label: 'Description', type: 'string' },
    { key: 'amount', label: 'Amount', type: 'number', sum: true },
  ],
  rows: [
    { description: 'Service', amount: 10 },
    { description: 'Extra', amount: 2.5 },
  ],
});
await writeFile(`${output}/report.html`, reportSnapshotHtml(snapshot), { mode: 0o600 });
await writeFile(`${output}/invoice.txt`, delivery.artifact!, { mode: 0o600 });
await writeFile(
  `${output}/evidence.json`,
  JSON.stringify(
    { components: expansion.lock, regression, trace: rendered.trace, delivery: delivery.evidence },
    null,
    2,
  ),
  { mode: 0o600 },
);
await writeFile(`${output}/replay.json`, JSON.stringify(bundle, null, 2), { mode: 0o600 });
console.log(
  `Verified contract → components → render/trace → regression → replay → review → delivery; evidence in ${output}`,
);
