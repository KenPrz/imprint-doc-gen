import test from 'node:test';
import assert from 'node:assert/strict';
import { render, compile } from '../src/index.ts';
import { readPackage, writePackage } from '../src/package.ts';
import { syntheticCases } from '../scripts/pms/synthetic-cases.ts';
import { syntheticData, syntheticInvoice } from '../scripts/pms/synthetic-data.ts';
import { checkSynthetic } from '../scripts/pms/synthetic-check.ts';
const { cases } = await syntheticCases();
const compiled = new Map<string, Awaited<ReturnType<typeof compile>>>();
for (const item of cases) test(`Synthetic PMS ${item.id}: data, structure and image placement`, async () => {
  if (!compiled.has(item.source)) compiled.set(item.source, await compile(item.template));
  const bytes = await compiled.get(item.source)!.render(item.data, { attachments: item.attachments });
  await checkSynthetic(item.template, bytes, item.data, item.format, item.expectedImage);
});
test('Synthetic data is deterministic, preserves source samples and reconciles folio/invoice arithmetic', () => {
  const input = { property_name: 'Original', guest_name: 'Original', rows: [{ name: 'Original' }] }, before = structuredClone(input);
  assert.deepEqual(syntheticData('demo', input, 'normal'), syntheticData('demo', input, 'normal'));
  assert.notDeepEqual(syntheticData('demo', input, 'normal', 1), syntheticData('demo', input, 'normal', 2));
  assert.deepEqual(input, before);
  for (const scenario of ['normal', 'empty', 'multilingual', 'large'] as const) {
    const invoice = syntheticInvoice(scenario), total = invoice.items.reduce((sum: number, item: { quantity: number; price: number }) => sum + item.quantity * item.price, 0);
    assert.equal(Number(invoice.total.replaceAll(',', '')), total);
    const folio = cases.find(item => item.report === 'folio_statement' && item.scenario === scenario)!.data;
    assert.equal(folio.transactions.reduce((sum: number, row: { amount_raw: string }) => sum + Number(row.amount_raw), 0), 0);
    assert.equal(folio.total_charges, folio.total_payments); assert.equal(folio.balance, '0.00');
  }
});
test('Synthetic checker rejects incorrect guest data and missing supplied images', async () => {
  const item = cases.find(item => item.id === 'registration_form-normal')!;
  const good = await render(item.template, item.data, { attachments: item.attachments });
  const parts = await readPackage(good), document = 'word/document.xml';
  parts.set(document, Buffer.from(parts.get(document)!.toString().replace(item.data.property_name, 'WRONG PROPERTY')));
  await assert.rejects(checkSynthetic(item.template, await writePackage(parts), item.data, item.format, item.expectedImage), /Missing inserted value/);
  await assert.rejects(checkSynthetic(item.template, good, item.data, item.format, Buffer.from('not the supplied picture')), /Expected supplied image/);
});
test('Synthetic checker rejects a wrong calculated invoice total', async () => {
  const item = cases.find(item => item.id === 'invoice-xlsx-empty')!;
  const parts = await readPackage(await render(item.template, item.data, { attachments: item.attachments }));
  const path = 'xl/worksheets/sheet1.xml';
  const source = parts.get(path)!.toString(), wrong = source.replace(/(<f>SUM\([^<]+<\/f><v>)0(<\/v>)/, '$1999$2');
  assert.notEqual(source, wrong, 'The mutation must alter the total cache'); parts.set(path, Buffer.from(wrong));
  await assert.rejects(checkSynthetic(item.template, await writePackage(parts), item.data, item.format, item.expectedImage), /Invoice total/);
});
