import assert from 'node:assert/strict';
import { readFile, mkdir, writeFile } from 'node:fs/promises';
import { fileURLToPath } from 'node:url';

// Run against a separate Compose project with GOTENBERG_API_TIMEOUT=5s.
// The unsupported reference document exercises the converter directly, not the renderer.
const root = fileURLToPath(new URL('../', import.meta.url));
const endpoint = process.env.GOTENBERG_URL ?? 'http://gotenberg:3000';
const results = [];
for (let cycle = 1; cycle <= 2; cycle++) {
  for (const [level, shouldFail] of [[93, true], [1, false]] as const) {
    const form = new FormData();
    const bytes = await readFile(`${root}corpus/docx/level-${String(level).padStart(3, '0')}/expected.docx`);
    form.append('files', new Blob([bytes]), `level-${level}.docx`);
    const started = performance.now();
    const response = await fetch(new URL('/forms/libreoffice/convert', endpoint), { method: 'POST', body: form, signal: AbortSignal.timeout(20_000) });
    const output = Buffer.from(await response.arrayBuffer());
    const elapsedMs = Math.round(performance.now() - started);
    console.log(JSON.stringify({ cycle, level, status: response.status, elapsedMs, error: response.ok ? undefined : output.toString() }));
    if (shouldFail) assert.equal(response.status, 503, 'Stress case must hit the 5-second conversion deadline');
    else { assert.equal(response.status, 200, output.toString()); assert.equal(output.subarray(0, 5).toString(), '%PDF-'); }
    assert(elapsedMs < 15_000, 'Failure and recovery must stay bounded');
    results.push({ cycle, level, status: response.status, elapsedMs });
  }
}
const apiEndpoint = process.env.SMOKE_API_URL ?? 'http://app:3000';
const key = process.env.IMPRINT_API_KEY;
assert(key, 'IMPRINT_API_KEY is required for the API recovery check');
const apiResults = [];
for (const level of [93, 1]) {
  const template = await readFile(`${root}corpus/docx/level-${String(level).padStart(3, '0')}/expected.docx`);
  const started = performance.now();
  const response: Response = await fetch(new URL('/v1/render', apiEndpoint), {
    method: 'POST', headers: { authorization: `Bearer ${key}`, 'content-type': 'application/json' },
    body: JSON.stringify({ template: template.toString('base64'), data: {}, output: 'pdf' }), signal: AbortSignal.timeout(20_000),
  });
  const elapsedMs = Math.round(performance.now() - started);
  if (level === 93) {
    assert.equal(response.status, 504);
    assert.match(response.headers.get('content-type')!, /application\/json/);
    assert.equal(response.headers.get('content-disposition'), null);
    assert.equal((await response.json() as { error: string }).error, 'RENDER_TIMEOUT');
  } else {
    assert.equal(response.status, 200, await (response.ok ? Promise.resolve('') : response.text()));
    assert.equal(Buffer.from(await response.arrayBuffer()).subarray(0, 5).toString(), '%PDF-');
  }
  apiResults.push({ level, status: response.status, elapsedMs });
  console.log(JSON.stringify({ api: true, level, status: response.status, elapsedMs }));
}
await mkdir(`${root}corpus/results`, { recursive: true });
await writeFile(`${root}corpus/results/converter-recovery.json`, JSON.stringify({ timeout: '5s', results, apiResults }, null, 2) + '\n');
console.log('Two timed-out documents were each followed immediately by a successful conversion.');
