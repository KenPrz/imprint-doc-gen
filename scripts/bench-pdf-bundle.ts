// Times POST /v1/pdf/bundle at the target size: 20 inputs of up to 10 MB (10 phone photos, 10 image-heavy PDFs).
// Usage: node scripts/bench-pdf-bundle.ts  (needs pdfcpu on PATH, Linux for /proc).
// Runs the server as a child process and prints wall time, output size and the server's peak RSS (VmHWM).
import { spawn } from 'node:child_process';
import { readFile } from 'node:fs/promises';
import { PDFDocument } from 'pdf-lib';
import sharp from 'sharp';

const limit = 10 * 1024 * 1024;
// Noise compresses badly, so a 12 MP photo lands close to the 10 MB limit, like a real high-detail shot.
const photo = async (seed: number) => {
  const raw = Buffer.alloc(4000 * 3000 * 3);
  for (let i = 0, x = seed; i < raw.length; i++) raw[i] = (x = (x * 1103515245 + 12345) >>> 0) >>> 24;
  for (let quality = 92; ; quality -= 4) {
    const jpeg = await sharp(raw, { raw: { width: 4000, height: 3000, channels: 3 } })
      .jpeg({ quality })
      .withMetadata({ orientation: 6 })
      .toBuffer();
    if (jpeg.length <= limit) return jpeg;
  }
};
const pdfOf = async (jpeg: Buffer) => {
  const doc = await PDFDocument.create();
  const image = await doc.embedJpg(jpeg);
  doc.addPage([842, 595]).drawImage(image, { x: 0, y: 0, width: 842, height: 595 });
  return Buffer.from(await doc.save());
};

const inputs: Buffer[] = [];
for (let i = 0; i < 10; i++) {
  const jpeg = await photo(i + 1);
  inputs.push(jpeg, await pdfOf(jpeg));
}
const form = new FormData();
for (const flag of ['decrypt', 'flatten', 'stripActiveContent']) form.append(flag, 'true');
for (const input of inputs) form.append('files', new Blob([new Uint8Array(input)]));
const total = inputs.reduce((n, b) => n + b.length, 0);

const apiKey = 'b'.repeat(32);
const port = 3000 + Math.floor(Math.random() * 1000) + 5000;
const server = spawn(process.execPath, ['src/server.ts'], {
  env: { ...process.env, IMPRINT_API_KEY: apiKey, PORT: String(port), IMPRINT_CONCURRENCY: '1' },
  stdio: 'ignore',
});
const address = `http://127.0.0.1:${port}`;
for (let i = 0; ; i++) {
  if ((await fetch(`${address}/health`).catch(() => undefined))?.ok) break;
  if (i > 100) throw new Error('Server did not start');
  await new Promise((r) => setTimeout(r, 100));
}
const started = performance.now();
try {
  const response = await fetch(`${address}/v1/pdf/bundle`, {
    method: 'POST',
    headers: { authorization: `Bearer ${apiKey}` },
    body: form,
  });
  const body = Buffer.from(await response.arrayBuffer());
  if (!response.ok) throw new Error(`${response.status} ${body.toString()}`);
  const seconds = (performance.now() - started) / 1000;
  const status = await readFile(`/proc/${server.pid}/status`, 'utf8');
  const peakKiB = Number(/VmHWM:\s+(\d+)/.exec(status)?.[1]);
  console.log(
    JSON.stringify({
      inputs: inputs.length,
      inputMiB: +(total / 2 ** 20).toFixed(1),
      seconds: +seconds.toFixed(1),
      pages: (await PDFDocument.load(body)).getPageCount(),
      outputMiB: +(body.length / 2 ** 20).toFixed(1),
      serverPeakRssMiB: Math.round(peakKiB / 1024),
    }),
  );
} finally {
  server.kill();
}
