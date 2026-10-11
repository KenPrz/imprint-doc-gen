import { readFile, writeFile, mkdir } from 'node:fs/promises';
import { resolve, join } from 'node:path';
import { fileURLToPath } from 'node:url';
import assert from 'node:assert/strict';
import { hash, unpack, type Format } from './common.ts';
import { validatePackage, assertGolden, compare } from './inspect.ts';
import type { Data } from './recipes.ts';

const root = resolve(fileURLToPath(new URL('../../corpus', import.meta.url)));
const manifest = JSON.parse(await readFile(join(root, 'manifest.json'), 'utf8'));
const results = [];
for (const entry of manifest.cases as { path: string; id: string; format: Format; level: number }[]) {
  try {
    const dir = join(root, entry.path);
    const spec = JSON.parse(await readFile(join(dir, 'case.json'), 'utf8'));
    const data: Data = JSON.parse(await readFile(join(dir, 'input.json'), 'utf8'));
    const templateBytes = await readFile(join(dir, spec.template)), expectedBytes = await readFile(join(dir, spec.expected));
    assert.equal(hash(templateBytes), spec.sha256.template); assert.equal(hash(expectedBytes), spec.sha256.expected);
    const template = await unpack(templateBytes), expected = await unpack(expectedBytes);
    validatePackage(template); validatePackage(expected); assertGolden(expected, entry.format, entry.level, data);
    compare(expected, expected, entry.format, spec.preservedParts);
    assert.throws(() => compare(template, expected, entry.format, spec.preservedParts), 'Unchanged template must fail');
    for (const name of spec.preservedParts) assert.equal(hash(template.get(name)!), hash(expected.get(name)!));
    results.push({ id: entry.id, status: 'passed' });
  } catch (error) { results.push({ id: entry.id, status: 'failed', error: String(error) }); }
}
assert.equal(results.length, 200);
const failed = results.filter(r => r.status === 'failed');
const report = { kind: 'fixture-validation', checkedDocuments: 400, cases: results.length, passed: results.length - failed.length, failed: failed.length, engineTestsRun: 0, officeApplicationTestsRun: 0, results };
await mkdir(join(root, 'results'), { recursive: true });
await writeFile(join(root, 'results/fixture-validation.json'), JSON.stringify(report, null, 2) + '\n');
console.log(`${report.passed}/200 fixture cases passed; 400 document packages checked. Engine and Office validation: not run.`);
if (failed.length) { console.error(JSON.stringify(failed.slice(0, 5), null, 2)); process.exitCode = 1; }
