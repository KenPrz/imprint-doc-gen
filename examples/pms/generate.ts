import { readFile, mkdir, writeFile } from 'node:fs/promises';
import { randomUUID } from 'node:crypto';
import { ImprintClient } from './client.ts';

const key = process.env.IMPRINT_API_KEY;
if (!key) throw new Error('Set IMPRINT_API_KEY to a tenant key (or the default installation key for this demo)');
const client = new ImprintClient(process.env.IMPRINT_URL ?? 'http://127.0.0.1:8088', key);
const format = process.argv[2] ?? 'pdf';
if (!['docx', 'xlsx', 'pdf'].includes(format)) throw new Error('Usage: node examples/pms/generate.ts [pdf|docx|xlsx]');
const source = format === 'xlsx' ? 'xlsx' : 'docx';
const data = JSON.parse(await readFile(new URL('./booking.json', import.meta.url), 'utf8'));
// Upload/publish only when your template changes. Save these IDs in PMS configuration.
const template = await client.upload('PMS guest invoice', await readFile(new URL(`./invoice.${source}`, import.meta.url)));
await client.publish(template.templateId, template.versionId);
const job = await client.submit({
  templateId: template.templateId, data,
  attachments: { logo: await readFile(new URL('./logo.png', import.meta.url)) },
  output: format as 'docx' | 'xlsx' | 'pdf',
}, `demo-${randomUUID()}`);
// In production save job.id on the booking, and poll it from a background task.
// Use e.g. booking-42-invoice-revision-3 as a stable idempotency key for retries.
console.log(`Submitted ${job.id} with template version ${job.versionId}`);
await client.wait(job.id);
await mkdir('artifacts', { recursive: true });
const path = `artifacts/booking-invoice.${format}`;
await writeFile(path, await client.download(job.id));
console.log(`Saved ${path}`);
