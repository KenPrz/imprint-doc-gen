import test from 'node:test';
import assert from 'node:assert/strict';
import { render, compile } from '../src/index.ts';
import { readPackage, writePackage } from '../src/package.ts';
import { syntheticCases } from '../scripts/synthetic/synthetic-cases.ts';
import { syntheticInvoice } from '../scripts/synthetic/synthetic-data.ts';
import { checkSynthetic } from '../scripts/synthetic/synthetic-check.ts';
const { cases } = await syntheticCases();
const compiled = new Map<string, Awaited<ReturnType<typeof compile>>>();
for (const item of cases) test(`Synthetic invoice ${item.id}: data, structure and image placement`, async () => {
  if (!compiled.has(item.source)) compiled.set(item.source, await compile(item.template));
  const bytes = await compiled.get(item.source)!.render(item.data, { attachments: item.attachments });
  await checkSynthetic(item.template, bytes, item.data, item.format, item.expectedImage);
});
test('Synthetic data is deterministic and reconciles invoice arithmetic', () => {
  assert.deepEqual(syntheticInvoice('normal'), syntheticInvoice('normal'));
  assert.notDeepEqual(syntheticInvoice('normal', 1), syntheticInvoice('normal', 2));
  for (const scenario of ['normal', 'empty', 'multilingual', 'large'] as const) {
    const invoice = syntheticInvoice(scenario), total = invoice.items.reduce((sum: number, item: { quantity: number; price: number }) => sum + item.quantity * item.price, 0);
    assert.equal(Number(invoice.total.replaceAll(',', '')), total);
  }
});
test('Synthetic checker rejects incorrect customer data and missing supplied images', async () => {
  const item = cases.find(item => item.id === 'signature-docx-normal')!;
  const good = await render(item.template, item.data, { attachments: item.attachments });
  const parts = await readPackage(good), document = 'word/document.xml';
  parts.set(document, Buffer.from(parts.get(document)!.toString().replace(item.data.company, 'WRONG COMPANY')));
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
