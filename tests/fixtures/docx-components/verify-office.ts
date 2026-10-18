/** Independent acceptance: node tests/fixtures/docx-components/verify-office.ts /absolute/output/directory */
import assert from 'node:assert/strict';
import { createHash } from 'node:crypto';
import { execFileSync } from 'node:child_process';
import { mkdir, readFile, writeFile } from 'node:fs/promises';
import { resolve } from 'node:path';
import { pathToFileURL } from 'node:url';
import { assembleDocxComponents } from '../../../src/docx-components.ts';
const output = resolve(process.argv[2] ?? 'artifacts/docx-components');
await mkdir(output, { recursive: true });
const host = await readFile(new URL('./host.docx', import.meta.url));
const document = await readFile(new URL('./clause.docx', import.meta.url));
const result = await assembleDocxComponents(host, [
  { id: 'clause', version: '1', document, sha256: createHash('sha256').update(document).digest('hex') },
]);
const assembled = resolve(output, 'assembled.docx');
await writeFile(assembled, result.document);
execFileSync(
  'libreoffice',
  [
    `-env:UserInstallation=${pathToFileURL(resolve(output, 'lo-profile')).href}`,
    '--headless',
    '--convert-to',
    'pdf',
    '--outdir',
    output,
    assembled,
  ],
  { timeout: 60000, stdio: 'pipe' },
);
const pdf = resolve(output, 'assembled.pdf');
const text = execFileSync('pdftotext', ['-layout', pdf, '-'], { encoding: 'utf8' });
for (const expected of [
  /Warranty clause/g,
  /Service warranty covers defects\./g,
  /1\.\s+First requirement/g,
  /2\.\s+Second requirement/g,
  /Parts\s+90/g,
])
  assert.equal([...text.matchAll(expected)].length, 2, `Expected two occurrences: ${expected}`);
assert.ok(!/3\.\s+First requirement|4\.\s+Second requirement/.test(text), 'Component lists must restart');
const images = execFileSync('pdfimages', ['-list', pdf], { encoding: 'utf8' });
assert.equal(
  images.split('\n').filter((line) => /^\s*\d+\s+\d+\s+image\s/.test(line)).length,
  2,
  'Both embedded PNG images must survive conversion',
);
execFileSync('pdftoppm', ['-scale-to', '1000', '-png', '-singlefile', pdf, resolve(output, 'page')], {
  timeout: 60000,
});
await writeFile(
  resolve(output, 'validation.json'),
  JSON.stringify(
    {
      converter: execFileSync('libreoffice', ['--version'], { encoding: 'utf8' }).trim(),
      checks: [
        'two styled clauses',
        'two numbered lists restart at 1',
        'two populated tables',
        'two PNG images',
      ],
      lock: result.lock,
    },
    null,
    2,
  ) + '\n',
);
console.log(`DOCX component Office acceptance passed: ${output}`);
