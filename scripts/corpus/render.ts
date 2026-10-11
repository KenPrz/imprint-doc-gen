import { readFile, writeFile, mkdir, rm } from 'node:fs/promises';
import { resolve, join } from 'node:path';
import { fileURLToPath } from 'node:url';
import { performance } from 'node:perf_hooks';
import { createHash } from 'node:crypto';
import { render, RenderError } from '../../src/index.ts';

const root = resolve(fileURLToPath(new URL('../../corpus', import.meta.url)));
const manifest = JSON.parse(await readFile(join(root, 'manifest.json'), 'utf8'));
const results = [];
for (const entry of manifest.cases) {
  const dir = join(root, entry.path), output = join(root, 'actual', entry.format, `level-${String(entry.level).padStart(3, '0')}.${entry.format}`);
  await mkdir(join(root, 'actual', entry.format), { recursive: true });
  await rm(output, { force: true }); // Remove only this case's previous generated output; stale passes must not survive.
  const start = performance.now();
  const templateBytes = await readFile(join(dir, `template.${entry.format}`));
  const templateSha256 = createHash('sha256').update(templateBytes).digest('hex');
  try {
    const spec = JSON.parse(await readFile(join(dir, 'case.json'), 'utf8'));
    const data = JSON.parse(await readFile(join(dir, 'input.json'), 'utf8'));
    const attachments: Record<string, Uint8Array> = {};
    for (const [key, path] of Object.entries(spec.attachments as Record<string, string>)) attachments[key] = await readFile(resolve(dir, path));
    // Deliberately reads no expected output. The separate scorer owns the oracle.
    const bytes = await render(templateBytes, data, { attachments });
    await writeFile(output, bytes);
    results.push({ id: entry.id, templateSha256, status: 'rendered', bytes: bytes.length, elapsedMs: Math.round(performance.now() - start) });
  } catch (error) {
    results.push({ id: entry.id, templateSha256, status: error instanceof RenderError && error.code.startsWith('UNSUPPORTED') ? 'unsupported' : 'failed', code: error instanceof RenderError ? error.code : 'ERROR', message: (error as Error).message });
  }
  if (entry.level % 25 === 0) console.log(`Attempted ${entry.format} levels 1–${entry.level}`);
}
await mkdir(join(root, 'results'), { recursive: true });
await writeFile(join(root, 'results/renderer-attempts.json'), JSON.stringify({ engine: 'imprint-doc-gen', milestone: 1, results }, null, 2) + '\n');
console.log(JSON.stringify({ rendered: results.filter(r => r.status === 'rendered').length, unsupported: results.filter(r => r.status === 'unsupported').length, failed: results.filter(r => r.status === 'failed').length }));
if (results.some(r => r.status === 'failed')) process.exitCode = 1;
