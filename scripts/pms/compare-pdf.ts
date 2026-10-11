import assert from 'node:assert/strict';
import { readFile, writeFile, mkdir, readdir, rm } from 'node:fs/promises';
import { execFile } from 'node:child_process';
import { promisify } from 'node:util';
import sharp from 'sharp';
import { render } from '../../src/index.ts';
import { convertToPdf } from '../../src/converter.ts';
const execute = promisify(execFile);
const endpoint = process.env.GOTENBERG_URL;
if (!endpoint) throw new Error('Set GOTENBERG_URL for the independent PDF comparison');
const reports = ['confirmation_letter', 'folio_statement', 'trial_balance', 'task_sheet', 'registration_form'];
const results: { report: string; pages: number; changedPixels: number }[] = [];
for (const name of reports) {
  const directory = `corpus/results/pms/visual/${name}`;
  await rm(directory, { recursive: true, force: true });
  await mkdir(directory, { recursive: true });
  const template = await readFile(`tests/fixtures/pms/${name}-template.docx`), data = JSON.parse(await readFile(`tests/fixtures/pms/${name}.json`, 'utf8'));
  const documents = { imprint: await render(template, data), carbone: await readFile(`corpus/results/pms/oracle/${name}-template.docx`) };
  for (const [engine, bytes] of Object.entries(documents)) {
    const pdf = await convertToPdf(bytes, endpoint);
    await writeFile(`${directory}/${engine}.pdf`, pdf);
    await execute('pdftoppm', ['-png', '-r', '96', `${directory}/${engine}.pdf`, `${directory}/${engine}`]);
  }
  const files = await readdir(directory), actual = files.filter(n => /^imprint-\d+\.png$/.test(n)).sort(), expected = files.filter(n => /^carbone-\d+\.png$/.test(n)).sort();
  assert.equal(actual.length, expected.length, `${name}: page count`);
  let changedPixels = 0;
  for (let i = 0; i < actual.length; i++) {
    const a = await sharp(`${directory}/${actual[i]}`).raw().toBuffer({ resolveWithObject: true });
    const b = await sharp(`${directory}/${expected[i]}`).raw().toBuffer({ resolveWithObject: true });
    assert.deepEqual(a.info, b.info, `${name}: page dimensions`);
    for (let p = 0; p < a.data.length; p += a.info.channels) if (!a.data.subarray(p, p + a.info.channels).equals(b.data.subarray(p, p + b.info.channels))) changedPixels++;
  }
  results.push({ report: name, pages: actual.length, changedPixels });
  console.log(JSON.stringify(results.at(-1)));
}
await writeFile('corpus/results/pms/visual-comparison.json', JSON.stringify({ dpi: 96, results }, null, 2) + '\n');
assert(results.every(result => result.changedPixels === 0), 'PDF appearance differs; inspect the page images');
