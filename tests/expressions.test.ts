import test from 'node:test';
import assert from 'node:assert/strict';
import { readFile } from 'node:fs/promises';
import { readPackage, writePackage } from '../src/package.ts';
import { render } from '../src/index.ts';
import {
  parseExpression,
  formatValue,
  filterArray,
  expressionSettings,
  renderContext,
} from '../src/expressions.ts';
import { escapeXml, parseXml, descendants, ns } from '../src/xml.ts';
async function template(text: string) {
  const parts = await readPackage(await readFile('examples/invoice/invoice.docx'));
  parts.set(
    'word/document.xml',
    Buffer.from(
      parts
        .get('word/document.xml')!
        .toString()
        .replace(
          /<w:body>[\s\S]*?<w:sectPr/,
          `<w:body><w:p><w:r><w:t>${escapeXml(text)}</w:t></w:r></w:p><w:sectPr`,
        ),
    ),
  );
  return writePackage(parts);
}
async function text(bytes: Uint8Array) {
  const parts = await readPackage(bytes);
  return descendants(parseXml(parts.get('word/document.xml')!.toString(), 'test'), ns.word, 't')
    .map((n) => n.textContent)
    .join('');
}
test('Formatting uses explicit locale and remains isolated across concurrent renders', async () => {
  const input = await template(
    '{d.amount:number(2)} | {d.enabled:if("Yes","No")} | {d.missing:default("Unknown")}',
  );
  const [a, b] = await Promise.all([
    render(input, { amount: 1234.5, enabled: true }, { locale: 'en-US' }),
    render(input, { amount: 1234.5, enabled: false }, { locale: 'de-DE' }),
  ]);
  assert.equal(await text(a), '1,234.50 | Yes | Unknown');
  assert.equal(await text(b), '1.234,50 | No | Unknown');
});
test('Dates reject ambiguous values and implement both spreadsheet epochs', () => {
  assert.equal(formatValue('1900-03-01', [{ name: 'dateSerial', args: ['1900'] }]), 61);
  assert.equal(formatValue('1904-01-01', [{ name: 'dateSerial', args: ['1904'] }]), 0);
  assert.throws(() => formatValue('2026-02-30', [{ name: 'date', args: [] }]), /Invalid date/);
  assert.throws(() => formatValue('01/02/2026', [{ name: 'date', args: [] }]), /ISO/);
  assert.equal(
    renderContext.run(expressionSettings('en-US', 'UTC'), () =>
      formatValue(12.5, [{ name: 'currency', args: ['USD'] }]),
    ),
    '$12.50',
  );
});
test('Filter grammar evaluates owned fields without executing code', () => {
  const parsed = parseExpression('d.items:where("active",true)[i].name:upper()', 'test');
  assert.equal(parsed.path, 'd.items[i].name');
  assert.deepEqual(filterArray([{ active: false }, { active: true }], parsed.filters, 'test'), [
    { active: true },
  ]);
  for (const expression of [
    'd.x:number(process.exit())',
    'd.x:default({})',
    'd.x:number(100)',
    'd.items:where("__proto__",true)[i].name',
  ])
    assert.throws(() => parseExpression(expression, 'test'));
});
test('DOCX filtered repetition excludes items before expansion', async () => {
  const parts = await readPackage(await readFile('examples/invoice/invoice.docx'));
  const xml = parts
    .get('word/document.xml')!
    .toString()
    .replaceAll('{d.items[i].description}', '{d.items:where("active",true)[i].description}');
  parts.set('word/document.xml', Buffer.from(xml));
  const result = await render(await writePackage(parts), {
    company: 'Example',
    items: [
      { active: true, description: 'KEEP' },
      { active: false, description: 'DROP' },
    ],
  });
  assert((await text(result)).includes('KEEP'));
  assert(!(await text(result)).includes('DROP'));
});
test('formatter string literals can contain braces without introducing template expressions', async () => {
  const { standardBindings, standardLookup } = await import('../src/standard-syntax.ts');
  const tag = '{d.value:default("{literal}")}';
  const binding = standardBindings(tag, 'test')[0]!;
  assert.equal(binding.end, tag.length);
  assert.equal(standardLookup([{}], binding.path, 'test', binding.formatters), '{literal}');
});
