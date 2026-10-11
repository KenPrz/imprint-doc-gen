import assert from 'node:assert/strict';
import { readFile, mkdir, writeFile } from 'node:fs/promises';
import { execFile } from 'node:child_process';
import { promisify } from 'node:util';
import { readPackage } from '../src/package.ts';
import { snapshot } from './pms/snapshot.ts';
import { ImprintClient } from '../examples/pms/client.ts';

const execute = promisify(execFile);
const endpoint = process.env.IMPRINT_URL ?? 'http://127.0.0.1:8091';
const adminKey = process.env.IMPRINT_API_KEY;
const container = process.env.IMPRINT_CONTAINER;
assert(adminKey && container, 'IMPRINT_API_KEY and IMPRINT_CONTAINER are required');
const docker = (...args: string[]) => execute('docker', args, { timeout: 125_000 });
const delay = (ms: number) => new Promise(resolve => setTimeout(resolve, ms));
async function ready() {
  for (let n = 0; n < 120; n++) {
    try { if ((await fetch(`${endpoint}/ready`, { signal: AbortSignal.timeout(1000) })).ok) return; } catch { /* Starting. */ }
    await delay(500);
  }
  throw new Error('Container did not become ready');
}
async function request(path: string, method: string, key: string, body?: object) {
  return fetch(new URL(path, endpoint), { method, headers: { authorization: `Bearer ${key}`, 'content-type': 'application/json' }, body: body ? JSON.stringify(body) : undefined, signal: AbortSignal.timeout(35_000) });
}
await ready();
const ownerResponse = await request('/v1/tenants', 'POST', adminKey, { name: 'PMS integration check' });
assert.equal(ownerResponse.status, 201);
const owner = await ownerResponse.json() as { id: string; apiKey: string };
const client = new ImprintClient(endpoint, owner.apiKey);
const fixture = new URL('../examples/pms/', import.meta.url);
const data = JSON.parse(await readFile(new URL('booking.json', fixture), 'utf8'));
const logo = await readFile(new URL('logo.png', fixture));
let docxTemplate: { templateId: string; versionId: string } | undefined;
const successful: string[] = [];
for (const format of ['docx', 'xlsx', 'pdf'] as const) {
  const uploaded = await client.upload(`Invoice ${format}`, await readFile(new URL(`invoice.${format === 'xlsx' ? 'xlsx' : 'docx'}`, fixture)));
  await client.publish(uploaded.templateId, uploaded.versionId);
  if (format === 'docx') docxTemplate = uploaded;
  const input = { templateId: uploaded.templateId, data, attachments: { logo }, output: format };
  const job = await client.submit(input, `format-${format}`);
  assert.equal((await client.submit(input, `format-${format}`)).id, job.id);
  await client.wait(job.id);
  const bytes = await client.download(job.id);
  assert.equal(bytes.subarray(0, format === 'pdf' ? 5 : 2).toString(), format === 'pdf' ? '%PDF-' : 'PK');
  if (format !== 'pdf') {
    const parts = await readPackage(bytes);
    assert([...parts.values()].some(part => part.equals(logo)), 'The supplied logo must be embedded');
    const xml = parts.get(format === 'docx' ? 'word/document.xml' : 'xl/worksheets/sheet1.xml')!.toString();
    assert(xml.includes('Alex Rivera')); assert(xml.includes('Breakfast')); assert(!xml.includes('{{'));
    if (format === 'xlsx') { assert.match(xml, /<v>8000<\/v>/); assert.match(xml, /<v>8750<\/v>/); }
  }
  assert.equal((await request(`/v1/jobs/${job.id}/result`, 'GET', adminKey)).status, 404);
  successful.push(job.id);
  console.log(`Standalone ${format}: passed`);
}
// Exercise the original PMS request and template without rewriting any tags.
const pmsFixture = new URL('../tests/fixtures/pms/', import.meta.url);
const pmsTemplate = await readFile(new URL('folio_statement-template.docx', pmsFixture));
const pmsData = JSON.parse(await readFile(new URL('folio_statement.json', pmsFixture), 'utf8'));
for (const convertTo of ['docx', 'pdf']) {
  const response = await request('/render/template?download=true', 'POST', owner.apiKey, { template: pmsTemplate.toString('base64'), data: pmsData, convertTo });
  assert.equal(response.status, 200, await response.clone().text());
  const bytes = Buffer.from(await response.arrayBuffer());
  if (convertTo === 'pdf') assert.equal(bytes.subarray(0, 5).toString(), '%PDF-');
  else assert.deepEqual(await snapshot(bytes), JSON.parse(await readFile(new URL('folio_statement.expected.json', pmsFixture), 'utf8')));
  console.log(`Original PMS Carbone request, ${convertTo}: passed`);
}
const pmsUploaded = await client.upload('Original PMS folio', pmsTemplate);
await client.publish(pmsUploaded.templateId, pmsUploaded.versionId);
const pmsJob = await client.submit({ templateId: pmsUploaded.templateId, data: pmsData, output: 'docx' }, 'original-pms-folio');
await client.wait(pmsJob.id);
assert.deepEqual(await snapshot(await client.download(pmsJob.id)), JSON.parse(await readFile(new URL('folio_statement.expected.json', pmsFixture), 'utf8')));
// Freeze the API once work is demonstrably running, then kill the whole container.
// This exercises the persisted lease rather than relying on a graceful close.
const interrupted = await Promise.all(Array.from({ length: 6 }, (_, n) => client.submit({ templateId: docxTemplate!.templateId, data, attachments: { logo }, output: 'pdf' }, `restart-${n}`)));
let running = false;
for (let n = 0; n < 100; n++) {
  const jobs = await Promise.all(interrupted.map(job => client.status(job.id)));
  if (jobs.some(job => job.status === 'running')) { running = true; break; }
  await delay(20);
}
assert(running, 'Expected an active job before forced interruption');
await docker('update', '--restart=no', container);
await docker('kill', '--signal=KILL', container);
await docker('start', container);
await ready();
console.log('Container restarted; waiting for interrupted leases to recover (up to 150 seconds)');
await Promise.all(interrupted.map(job => client.wait(job.id, 200_000)));
for (const job of [...interrupted, ...successful.map(id => ({ id }))]) assert.equal((await client.download(job.id)).subarray(0, job.id === successful[0] || job.id === successful[1] ? 2 : 5).toString(), job.id === successful[0] || job.id === successful[1] ? 'PK' : '%PDF-');
console.log('Queued jobs, active leases, tenant credentials and results survived forced restart');
// A failed converter must terminate the container, not leave a healthy-looking API.
await docker('exec', container, 'node', '-e', "const fs=require('node:fs'); const pid=fs.readdirSync('/proc').find(p=>{try{return fs.readFileSync('/proc/'+p+'/comm','utf8').trim()==='gotenberg'}catch{return false}}); if(!pid)process.exit(2); process.kill(Number(pid),'SIGKILL')");
let exited = false;
for (let n = 0; n < 120; n++) {
  const state = (await docker('inspect', container, '--format', '{{.State.Status}} {{.State.ExitCode}}')).stdout.trim();
  if (state === 'exited 1') { exited = true; break; }
  await delay(500);
}
assert(exited, 'Supervisor must exit nonzero when the converter dies');
await docker('start', container); await ready();
await docker('update', '--restart=unless-stopped', container);
assert.equal((await client.download(successful[2]!)).subarray(0, 5).toString(), '%PDF-');
await mkdir('corpus/results', { recursive: true });
await writeFile('corpus/results/standalone-validation.json', JSON.stringify({ formats: ['docx', 'xlsx', 'pdf'], imageAttachment: true, originalPmsCarboneDocxAndPdf: true, originalPmsDurableJob: true, tenantIsolation: true, idempotency: true, queuedAndRunningRestartRecovery: true, interruptedJobs: interrupted.length, converterFailureExitsContainer: true, resultsPersist: true }, null, 2) + '\n');
console.log('Converter failure supervision: passed');
