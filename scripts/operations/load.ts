import assert from 'node:assert/strict';
import { readFile, mkdir, writeFile } from 'node:fs/promises';
import { cpus, totalmem, platform, arch } from 'node:os';
const endpoint = process.env.IMPRINT_URL,
  key = process.env.IMPRINT_API_KEY;
assert(endpoint && key, 'Set IMPRINT_URL and IMPRINT_API_KEY for an isolated test service');
const requests = Number(process.env.LOAD_REQUESTS ?? 30),
  concurrency = Number(process.env.LOAD_CONCURRENCY ?? 2),
  output = process.env.LOAD_FORMAT ?? 'docx';
assert(Number.isInteger(requests) && requests >= 1 && requests <= 500);
assert(Number.isInteger(concurrency) && concurrency >= 1 && concurrency <= 32);
assert(['docx', 'xlsx', 'pdf'].includes(output));
const template = await readFile(`examples/invoice/invoice.${output === 'xlsx' ? 'xlsx' : 'docx'}`),
  data = JSON.parse(await readFile('examples/invoice/data.json', 'utf8'));
const body = JSON.stringify({ template: template.toString('base64'), data, output }),
  results: { status: number; ms: number; bytes: number }[] = [];
let next = 0,
  peakApiRss = 0,
  sampling = false;
const sample = async () => {
  if (sampling) return;
  sampling = true;
  try {
    const r = await fetch(new URL('/metrics', endpoint), {
      headers: { authorization: `Bearer ${key}` },
      signal: AbortSignal.timeout(2000),
    });
    if (r.ok)
      peakApiRss = Math.max(
        peakApiRss,
        Number(/imprint_process_resident_bytes (\d+)/.exec(await r.text())?.[1] ?? 0),
      );
  } catch {
    /* Record unavailable metrics as zero. */
  } finally {
    sampling = false;
  }
};
await sample();
const timer = setInterval(() => void sample(), 500),
  started = performance.now();
try {
  await Promise.all(
    Array.from({ length: concurrency }, async () => {
      while (next++ < requests) {
        const start = performance.now();
        try {
          const r = await fetch(new URL('/v1/render', endpoint), {
            method: 'POST',
            headers: { authorization: `Bearer ${key}`, 'content-type': 'application/json' },
            body,
            signal: AbortSignal.timeout(100000),
          });
          const bytes = Buffer.from(await r.arrayBuffer());
          const valid =
            !r.ok ||
            bytes.subarray(0, output === 'pdf' ? 5 : 2).toString() === (output === 'pdf' ? '%PDF-' : 'PK');
          results.push({
            status: valid ? r.status : 0,
            ms: performance.now() - start,
            bytes: r.ok ? bytes.length : 0,
          });
        } catch {
          results.push({ status: 0, ms: performance.now() - start, bytes: 0 });
        }
      }
    }),
  );
} finally {
  clearInterval(timer);
}
await sample();
const elapsed = (performance.now() - started) / 1000,
  durations = results
    .filter((r) => r.status === 200)
    .map((r) => r.ms)
    .sort((a, b) => a - b);
const percentile = (p: number) => (durations.length ? durations[Math.ceil(durations.length * p) - 1] : null);
const report = {
  requests,
  concurrency,
  output,
  elapsedSeconds: elapsed,
  successful: durations.length,
  rejected: results.filter((r) => r.status === 503).length,
  failures: results.filter((r) => r.status !== 200 && r.status !== 503).length,
  successfulPerSecond: durations.length / elapsed,
  latencyMs: { p50: percentile(0.5), p95: percentile(0.95), p99: percentile(0.99) },
  peakApiRssBytes: peakApiRss,
  runner: {
    node: process.version,
    platform: platform(),
    arch: arch(),
    cpus: cpus().length,
    memoryBytes: totalmem(),
  },
  results,
};
await mkdir('artifacts/operations', { recursive: true });
await writeFile('artifacts/operations/load.json', JSON.stringify(report, null, 2));
console.log(JSON.stringify({ ...report, results: undefined }));
if (report.failures || !report.successful) process.exitCode = 1;
