import { readFile, writeFile, mkdir } from 'node:fs/promises';
import { resolve, join } from 'node:path';
import { fileURLToPath } from 'node:url';
import { unpack, type Format } from './common.ts';
import { compare } from './inspect.ts';

const corpus = resolve(fileURLToPath(new URL('../../corpus', import.meta.url)));
const actual = process.argv[2] ? resolve(process.argv[2]) : join(corpus, 'actual');
if (actual === corpus || actual.includes('/expected')) throw new Error('Supply an independent engine-output directory');
const manifest = JSON.parse(await readFile(join(corpus, 'manifest.json'), 'utf8'));
let attempts: { id: string; templateSha256: string; status: string; message?: string }[] = [];
if (actual === join(corpus, 'actual')) {
  try { attempts = JSON.parse(await readFile(join(corpus, 'results/renderer-attempts.json'), 'utf8')).results; }
  catch (error) { if ((error as NodeJS.ErrnoException).code !== 'ENOENT') throw error; }
}
const results: { id: string; format: Format; level: number; status: string; error?: string }[] = [];
for (const entry of manifest.cases) {
  const base = { id: entry.id, format: entry.format as Format, level: entry.level };
  let output: Buffer;
  try { output = await readFile(join(actual, entry.format, `level-${String(entry.level).padStart(3, '0')}.${entry.format}`)); }
  catch (error) {
    if ((error as NodeJS.ErrnoException).code !== 'ENOENT') throw error;
    const attempt = attempts.find(r => r.id === entry.id && r.templateSha256 === entry.sha256.template);
    results.push({ ...base, status: attempt?.status === 'unsupported' ? 'unsupported' : attempt?.status === 'failed' ? 'failed' : 'not-run', error: attempt?.message }); continue;
  }
  try {
    const spec = JSON.parse(await readFile(join(corpus, entry.path, 'case.json'), 'utf8'));
    const expected = await unpack(await readFile(join(corpus, entry.path, spec.expected)));
    compare(await unpack(output), expected, entry.format, spec.preservedParts);
    results.push({ ...base, status: 'passed' });
  } catch (error) { results.push({ ...base, status: 'failed', error: String(error).slice(0, 3000) }); }
}
const summary = Object.fromEntries((['docx', 'xlsx'] as Format[]).map(format => {
  const rows = results.filter(r => r.format === format); let highestContiguousLevel = 0;
  for (const row of rows) { if (row.status !== 'passed') break; highestContiguousLevel = row.level; }
  return [format, { passed: rows.filter(r => r.status === 'passed').length, failed: rows.filter(r => r.status === 'failed').length, unsupported: rows.filter(r => r.status === 'unsupported').length, notRun: rows.filter(r => r.status === 'not-run').length, highestContiguousLevel }];
}));
await mkdir(join(corpus, 'results'), { recursive: true });
await writeFile(join(corpus, 'results/engine-score.json'), JSON.stringify({ kind: 'supplied-output-comparison', actualDirectory: actual, visualValidation: 'not-run', summary, results }, null, 2) + '\n');
console.log(JSON.stringify(summary, null, 2));
if (results.some(r => r.status === 'failed')) process.exitCode = 1;
else if (results.some(r => r.status === 'not-run' || r.status === 'unsupported')) process.exitCode = 2;
