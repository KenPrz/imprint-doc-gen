import { spawnSync } from 'node:child_process';
import { readFile, mkdir, writeFile } from 'node:fs/promises';
import { resolve } from 'node:path';
import { pathToFileURL } from 'node:url';
import assert from 'node:assert/strict';
import { readPackage } from '../src/package.ts';
import { ns, parseXml, descendants } from '../src/xml.ts';
const dir = resolve('artifacts/feature-expansion');
const run = (command: string, args: string[]) => {
  const result = spawnSync(command, args, { encoding: 'utf8', timeout: 120000 });
  if (result.error || result.status !== 0)
    throw new Error(`${command} failed: ${result.error?.message ?? result.stderr}`);
  return result.stdout;
};
await mkdir(dir, { recursive: true });
await mkdir(resolve(dir, 'pdf'), { recursive: true });
await mkdir(resolve(dir, 'recalculated'), { recursive: true });
run(process.execPath, ['--test', 'tests/feature-expansion.test.ts']);
run(process.execPath, ['scripts/feature-office-check.ts']);
run(process.execPath, ['scripts/check-xlsx-calculation.ts', '--prepare']);
const profile = '-env:UserInstallation=' + pathToFileURL(resolve(dir, 'libreoffice-profile')).href;
const office = (format: string, output: string, input: string) =>
  run('libreoffice', [profile, '--headless', '--convert-to', format, '--outdir', output, input]);
for (const file of ['repeated-notes.docx', 'repeated-comments.docx', 'inline.docx', 'chart.xlsx'])
  office('pdf', resolve(dir, 'pdf'), resolve(dir, file));
const notes = run('pdftotext', [resolve(dir, 'pdf/repeated-notes.pdf'), '-']);
assert.equal(notes.match(/First note/g)?.length, 2, notes);
assert.equal(notes.match(/Second note/g)?.length, 2, notes);
const chart = run('pdftotext', [resolve(dir, 'pdf/chart.pdf'), '-']);
for (const word of ['January', 'February', 'Revenue', 'Costs']) assert(chart.includes(word), chart);
const inline = run('pdftotext', [resolve(dir, 'pdf/inline.pdf'), '-']);
assert(inline.includes('Before') && inline.includes('After'));
const images = run('pdfimages', ['-list', resolve(dir, 'pdf/inline.pdf')]);
assert.match(images, /image/);
for (const file of ['inline', 'repeated-notes', 'chart'])
  run('pdftoppm', [
    '-scale-to',
    '1100',
    '-singlefile',
    '-png',
    resolve(dir, `pdf/${file}.pdf`),
    resolve(dir, file),
  ]);
office('xlsx', resolve(dir, 'recalculated'), resolve(dir, 'formulas.ods'));
const parts = await readPackage(await readFile(resolve(dir, 'recalculated/formulas.xlsx')));
const sheet = parseXml(parts.get('xl/worksheets/sheet1.xml')!.toString(), 'sheet');
const values = new Map(
  descendants(sheet, ns.spreadsheet, 'c').map((c) => [
    c.getAttribute('r'),
    Number(descendants(c, ns.spreadsheet, 'v')[0]?.textContent),
  ]),
);
assert.deepEqual(
  ['A1', 'B1', 'A2', 'B2', 'A3'].map((k) => values.get(k)),
  [2, 4, 3, 6, 5],
);
office('xlsx', resolve(dir, 'recalculated'), resolve('corpus/results/calculation/imprint-calculated.xlsx'));
const comparison = run(process.execPath, [
  'scripts/check-xlsx-calculation.ts',
  '--compare',
  resolve(dir, 'recalculated/imprint-calculated.xlsx'),
]);
const report = {
  converter: run('libreoffice', ['--version']).trim(),
  wordNoteText: true,
  inlineImage: true,
  literalChart: true,
  odsCalculatedValues: [2, 4, 3, 6, 5],
  xlsxComparison: comparison.trim(),
};
await writeFile(resolve(dir, 'office-validation.json'), JSON.stringify(report, null, 2) + '\n');
console.log(JSON.stringify(report));
