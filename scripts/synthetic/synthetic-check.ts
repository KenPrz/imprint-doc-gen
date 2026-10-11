import assert from 'node:assert/strict';
import { readPackage } from '../../src/package.ts';
import { parseXml, descendants, ns } from '../../src/xml.ts';
import { validatePackage, snapshot as officeSnapshot } from '../corpus/inspect.ts';
import { snapshot } from './snapshot.ts';
import type { Data } from './synthetic-data.ts';

// Independent data walker and structural assertions; no compiler plans, path
// resolver or loop inference from src/standard-syntax.ts are used by this checker.
function valuesAt(data: unknown, path: string): unknown[] {
  let values = [data];
  for (const segment of path.split('.')) {
    const repeated = segment.endsWith('[i]'), key = repeated ? segment.slice(0, -3) : segment;
    values = values.flatMap(value => {
      const child = value && typeof value === 'object' ? (value as Data)[key] : undefined;
      return repeated ? Array.isArray(child) ? child : [] : [child];
    });
  }
  return values;
}
const wordText = (node: Parameters<typeof descendants>[0]) => descendants(node, ns.word, 't').map(n => n.textContent).join('');
export async function checkSynthetic(template: Uint8Array, output: Uint8Array, data: Data, format: 'docx' | 'xlsx', expectedImage?: Buffer) {
  const parts = await readPackage(output); validatePackage(parts);
  if (expectedImage) assert([...parts].some(([name, bytes]) => /\/media\//.test(name) && bytes.equals(expectedImage)), 'Expected supplied image bytes in output package');
  if (format === 'xlsx') {
    const sheet = officeSnapshot(parts, 'xlsx')['xl/worksheets/sheet1.xml'] as { cells: { ref: string; value: string; formula?: string }[] };
    const cells = new Map(sheet.cells.map(cell => [cell.ref, cell]));
    assert.equal(cells.get('A1')?.value, data.company);
    let total = 0;
    for (const [index, item] of data.items.entries()) {
      assert.equal(cells.get(`A${11 + index}`)?.value, item.description);
      assert.equal(Number(cells.get(`D${11 + index}`)?.value), item.quantity * item.price);
      total += item.quantity * item.price;
    }
    assert.equal(Number(cells.get(`D${11 + data.items.length}`)?.value), total, 'Invoice total, including the empty-list case');
    assert(!JSON.stringify(sheet).includes('{d.'));
    return { text: sheet.cells.map(c => c.value).join('\n'), tables: 1, repeatedItems: data.items.length };
  }
  const before = await readPackage(template), actual = await snapshot(output), original = await snapshot(template);
  let text = '', tables = 0, repeatedItems = 0;
  for (const [part, story] of Object.entries(actual)) {
    const rendered = story as { paragraphs: string[]; tables: string[][][][]; pictures: string[][] };
    const source = before.get(part)!;
    const document = parseXml(source.toString(), part), content = rendered.paragraphs.join('\n');
    text += content + '\n'; tables += rendered.tables.length;
    assert(!/\{\{|\{d[.[]/.test(content), `Unresolved text binding in ${part}`);
    assert(!/\{d\./.test(parts.get(part)!.toString()), `Unresolved image binding in ${part}`);
    assert.deepEqual(rendered.pictures, (original[part] as { pictures: string[][] }).pictures, `Picture dimensions changed in ${part}`);
    const tags = new Set(descendants(document, ns.word, 'p').flatMap(p => [...wordText(p).matchAll(/\{d\.([^{}]+)\}/g)].map(match => match[1]!)));
    for (const tag of tags) if (!tag.includes('[i+1]')) for (const value of valuesAt(data, tag)) {
      if (value === null || value === undefined || typeof value === 'object' || value === '') continue;
      assert(content.includes(String(value).replace(/\r\n?/g, '\n')), `Missing inserted value at ${tag}: ${String(value).slice(0, 90)}`);
    }
    const expectedRows: number[] = [];
    for (const table of descendants(document, ns.word, 'tbl')) {
      let staticRows = 0;
      const paths: string[] = [];
      for (const row of descendants(table, ns.word, 'tr')) {
        const rowText = wordText(row);
        if (rowText.includes('[i+1]')) continue;
        const marker = /\{d\.([^{}]+)\[i\]\./.exec(rowText);
        if (marker) paths.push(marker[1]!); else staticRows++;
      }
      const lengths = paths.map(path => valuesAt(data, path).map(value => Array.isArray(value) ? value.length : 0));
      const copies = paths.some(path => path.includes('[i]')) ? Math.max(0, ...lengths.map(a => a.length)) : 1;
      for (let n = 0; n < copies; n++) { const repeated = lengths.reduce((count, list) => count + (list[n] ?? 0), 0); repeatedItems += repeated; expectedRows.push(staticRows + repeated); }
    }
    assert.deepEqual(rendered.tables.map(table => table.length), expectedRows, `Repeated table structure in ${part}`);
  }
  return { text, tables, repeatedItems };
}
