// Independent LibreOffice comparison input. Run --prepare, round-trip the file
// with a separate LibreOffice profile, then run --compare <round-trip.xlsx>.
import assert from 'node:assert/strict';
import { readFile, writeFile, mkdir } from 'node:fs/promises';
import { resolve } from 'node:path';
import { buildXlsx } from './corpus/xlsx.ts';
import { inputData } from './corpus/recipes.ts';
import { pack, unpack, NS, esc, parse, elements } from './corpus/common.ts';
import { render } from '../src/index.ts';

const dir = resolve('corpus/results/calculation');
const formulas = [
  ["SUM('Report'!B40:B40)", "MIN('Report'!B40:B40)", "MAX('Report'!B40:B40)", "AVERAGE('Report'!B40:B40)", "COUNT('Report'!B40:B40)"],
  ['2^3^2', '-2^2', '2^-2', 'IF(1<2,"yes",1/0)', 'SUM("3",TRUE,2)'],
  ['2^0', '2^-1', '(-1)^0.5', '"TRUE"&"FALSE"', 'IF("TRUE",1,2)'],
  ['COUNT(1,1/0)', "SUM('Report'!A40:A40)", "'Report'!E40=\"\"", "COUNT('Report'!E40:E40)", '1/0'],
];
if (process.argv.includes('--prepare')) {
  await mkdir(dir, { recursive: true });
  const data = inputData(60), parts = buildXlsx(60, data, false);
  const rows = formulas.map((values, i) => `<row r="${i + 3}">${values.map((f, j) => `<c r="${String.fromCharCode(65 + j)}${i + 3}"><f>${esc(f)}</f><v>999</v></c>`).join('')}</row>`).join('');
  parts.set('xl/worksheets/sheet2.xml', Buffer.from(parts.get('xl/worksheets/sheet2.xml')!.toString().replace('</sheetData>', rows + '</sheetData>').replace('A1:B2', `A1:E${formulas.length + 2}`)));
  const attachments = { red: await readFile('corpus/attachments/red.png'), blue: await readFile('corpus/attachments/blue.png') };
  await writeFile(resolve(dir, 'imprint-calculated.xlsx'), await render(await pack(parts), data, { attachments }));
  console.log(resolve(dir, 'imprint-calculated.xlsx'));
} else {
  const index = process.argv.indexOf('--compare'), path = process.argv[index + 1];
  assert(index >= 0 && path, 'Expected --prepare or --compare <LibreOffice round-trip.xlsx>');
  const snapshot = async (path: string) => {
    const parts = await unpack(await readFile(path));
    const result: Record<string, { type: string; value: string }> = {};
    for (const name of ['xl/worksheets/sheet1.xml', 'xl/worksheets/sheet2.xml']) {
      const doc = parse(parts.get(name)!.toString());
      for (const cell of elements(doc, NS.s, 'c')) if (elements(cell, NS.s, 'f').length) result[`${name}:${cell.getAttribute('r')}`] = { type: cell.getAttribute('t') || 'n', value: elements(cell, NS.s, 'v')[0]?.textContent ?? '' };
    }
    return result;
  };
  const actual = await snapshot(resolve(dir, 'imprint-calculated.xlsx')), roundTrip = await snapshot(path);
  // LibreOffice may synthesize formulas for static boolean cells. Compare every
  // formula we emitted, without treating those extra formulas as a missing cache.
  const independent = Object.fromEntries(Object.keys(actual).map(key => [key, roundTrip[key]]));
  await writeFile(resolve(dir, 'formula-comparison.json'), JSON.stringify({ actual, independent }, null, 2) + '\n');
  assert.deepEqual(actual, independent);
  console.log(`${Object.keys(actual).length} formula caches match independent LibreOffice recalculation`);
}
