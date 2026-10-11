import assert from 'node:assert/strict';
import { mkdir, writeFile, rm } from 'node:fs/promises';
import { resolve, join, relative } from 'node:path';
import { execFile } from 'node:child_process';
import { promisify } from 'node:util';
import { createHash } from 'node:crypto';
import { compile } from '../../src/index.ts';
import { convertToPdf } from '../../src/converter.ts';
import { syntheticCases, type SyntheticCase } from './synthetic-cases.ts';
import { checkSynthetic } from './synthetic-check.ts';
const execute = promisify(execFile), args = process.argv.slice(2);
const pdf = args.includes('--pdf'), api = args.includes('--api'), smoke = args.includes('--smoke');
const outputIndex = args.indexOf('--output'), output = resolve(outputIndex < 0 ? 'artifacts/pms-synthetic' : args[outputIndex + 1]!);
const seed = Number(process.env.PMS_TEST_SEED ?? 20261011);
assert(Number.isSafeInteger(seed), 'PMS_TEST_SEED must be an integer');
if (api) assert(process.env.IMPRINT_URL && process.env.IMPRINT_API_KEY, '--api requires IMPRINT_URL and IMPRINT_API_KEY');
if (pdf && !api) assert(process.env.GOTENBERG_URL, '--pdf requires GOTENBERG_URL (or use --api)');
if (pdf) for (const command of ['pdfinfo', 'pdftotext', 'pdfimages', 'pdftoppm']) await execute(command, ['-v']);
const { cases: allCases, assets, sourceCommit } = await syntheticCases(seed);
const selected = new Set(['registration_form-normal', 'registration_form-multilingual', 'folio_statement-normal', 'task_sheet-normal', 'invoice-docx-normal', 'invoice-xlsx-normal', 'invoice-xlsx-empty']);
const cases = smoke ? allCases.filter(c => selected.has(c.id)) : allCases;
await mkdir(join(output, 'assets'), { recursive: true }); await mkdir(join(output, 'templates'), { recursive: true });
for (const [name, bytes] of Object.entries(assets)) await writeFile(join(output, 'assets', name), bytes);
const json = (value: unknown) => JSON.stringify(value, null, 2) + '\n';
const hash = (bytes: Uint8Array) => createHash('sha256').update(bytes).digest('hex');
const compiled = new Map<string, Awaited<ReturnType<typeof compile>>>();
async function remote(item: SyntheticCase, format: string) {
  const response = await fetch(new URL('/render/template?download=true', process.env.IMPRINT_URL), {
    method: 'POST', headers: { authorization: `Bearer ${process.env.IMPRINT_API_KEY}`, 'content-type': 'application/json' },
    body: JSON.stringify({ template: item.template.toString('base64'), data: item.data, attachments: Object.fromEntries(Object.entries(item.attachments).map(([key, bytes]) => [key, bytes.toString('base64')])), convertTo: format }), signal: AbortSignal.timeout(100000),
  });
  if (!response.ok) throw new Error(`Imprint returned ${response.status}: ${(await response.text()).slice(0, 300)}`);
  return Buffer.from(await response.arrayBuffer());
}
const results: Record<string, any>[] = [];
for (const item of cases) {
  const started = performance.now(), directory = join(output, item.id);
  await mkdir(directory, { recursive: true });
  for (const name of ['report.docx', 'report.xlsx', 'report.pdf', 'preview.png', 'pdf-text.txt']) await rm(join(directory, name), { force: true });
  const templateName = `${item.report}.${item.format}`;
  await writeFile(join(output, 'templates', templateName), item.template);
  await writeFile(join(directory, 'data.json'), json(item.data));
  const attachments: Record<string, string> = {};
  for (const [key, bytes] of Object.entries(item.attachments)) attachments[key] = '../assets/' + Object.keys(assets).find(name => assets[name]!.equals(bytes))!;
  await writeFile(join(directory, 'attachments.json'), json(attachments));
  const result: Record<string, any> = { id: item.id, report: item.report, scenario: item.scenario, format: item.format, source: item.source, templateSha256: hash(item.template), template: `templates/${templateName}`, imageMode: item.imageMode, status: 'failed', pdf: pdf ? 'pending' : 'not-run' };
  try {
    let bytes: Buffer;
    if (api) bytes = await remote(item, item.format);
    else {
      if (!compiled.has(item.source)) compiled.set(item.source, await compile(item.template));
      bytes = Buffer.from(await compiled.get(item.source)!.render(item.data, { attachments: item.attachments }));
    }
    await writeFile(join(directory, `report.${item.format}`), bytes);
    const semantic = await checkSynthetic(item.template, bytes, item.data, item.format, item.expectedImage);
    result.documentBytes = bytes.length; result.documentSha256 = hash(bytes); result.tables = semantic.tables; result.repeatedItems = semantic.repeatedItems; result.semantics = 'passed';
    if (pdf) {
      const pdfBytes = api ? await remote(item, 'pdf') : await convertToPdf(bytes, process.env.GOTENBERG_URL!, `document.${item.format}`);
      const path = join(directory, 'report.pdf'); await writeFile(path, pdfBytes);
      const [info, extracted, images] = await Promise.all([
        execute('pdfinfo', [path], { timeout: 30000 }), execute('pdftotext', ['-layout', path, '-'], { timeout: 30000, maxBuffer: 16 * 1024 * 1024 }), execute('pdfimages', ['-list', path], { timeout: 30000 }),
      ]);
      result.pages = Number(/Pages:\s+(\d+)/.exec(info.stdout)?.[1]); assert(result.pages > 0, 'Expected at least one PDF page');
      assert(!/\{d\.[^}]+\}/.test(extracted.stdout), 'PDF contains unresolved template tags');
      const compactPdf = extracted.stdout.normalize('NFC').replace(/\s+/g, '');
      for (const probe of ['Zoë', '山田', 'François', 'Αλέξανδρος']) if (semantic.text.includes(probe)) assert(compactPdf.includes(probe), `PDF text lost multilingual probe ${probe}`);
      if (semantic.text.includes('Imprint Test Hotel')) assert(extracted.stdout.replace(/\s+/g, ' ').includes('Imprint Test Hotel'), 'PDF text lost the synthetic property name');
      result.pdfImages = images.stdout.split('\n').filter(line => /^\s*\d+\s+\d+\s+image\s/.test(line)).length;
      if (item.expectedImage) assert(result.pdfImages > 0, 'Supplied image disappeared during PDF conversion');
      if (item.expectedImage) assert(images.stdout.split('\n').some(line => /^\s*1\s+\d+\s+image\s/.test(line)), 'Expected the signature/logo on the first PDF page');
      if (item.format === 'xlsx') { const firstPage = extracted.stdout.split('\f')[0]!.replace(/\s+/g, ' '); for (const label of ['Description', 'Quantity', 'Unit price', 'Amount']) assert(firstPage.includes(label), `Invoice column ${label} overflowed the first PDF page`); }
      await writeFile(join(directory, 'pdf-text.txt'), extracted.stdout);
      await execute('pdftoppm', ['-f', '1', '-singlefile', '-scale-to', '1000', '-png', path, join(directory, 'preview')], { timeout: 30000 });
      result.pdf = 'passed'; result.pdfBytes = pdfBytes.length;
    }
    result.status = 'passed';
  } catch (error) { result.error = String(error); if (result.pdf === 'pending') result.pdf = 'failed'; console.error(`${item.id}: ${result.error}`); }
  result.elapsedMs = Math.round(performance.now() - started); results.push(result);
  await writeFile(join(directory, 'result.json'), json(result));
  if (results.length % 8 === 0) console.log(`${results.length}/${cases.length} synthetic cases completed`);
}
const summary = { seed, sourceCommit, mode: api ? 'http-api' : 'sdk', scope: smoke ? 'smoke' : 'full', cases: results.length, passed: results.filter(r => r.status === 'passed').length, failed: results.filter(r => r.status !== 'passed').length, pdfPassed: results.filter(r => r.pdf === 'passed').length, pdfNotRun: results.filter(r => r.pdf === 'not-run').length, pages: results.reduce((n, r) => n + (r.pages ?? 0), 0), imageCasesPassed: results.filter(r => r.status === 'passed' && r.imageMode !== 'no-image-field' && r.imageMode !== 'optional-signature-omitted').length, results };
await writeFile(join(output, 'summary.json'), json(summary));
const esc = (value: unknown) => String(value).replace(/[&<>"']/g, c => ({ '&': '&amp;', '<': '&lt;', '>': '&gt;', '"': '&quot;', "'": '&#39;' }[c]!));
await writeFile(join(output, 'index.html'), `<!doctype html><html lang="en"><meta charset="utf-8"><meta name="viewport" content="width=device-width,initial-scale=1"><title>Imprint synthetic PMS reports</title><style>body{font:16px/1.5 system-ui;background:#f4f7fa;color:#183749;margin:32px}main{max-width:1300px;margin:auto}.grid{display:grid;grid-template-columns:repeat(auto-fit,minmax(270px,1fr));gap:18px}article{background:white;border:1px solid #d3dfe5;padding:16px;border-radius:8px}img{width:100%;height:290px;object-fit:contain;background:#eee}a{color:#00678a}small{display:block}h2{font-size:17px;overflow-wrap:anywhere}.failed{border-color:#b00}nav a{margin-right:16px}</style><main><h1>Synthetic PMS test reports</h1><p>Seed ${seed}. ${summary.passed}/${summary.cases} cases passed. ${summary.pdfPassed} PDFs checked; ${summary.pdfNotRun} not run. ${summary.pages} PDF pages.</p><p>All data and image assets are fictional. The 30 original seeded PMS templates retain their bytes. Invoice examples provide extra logo and XLSX tests. These checks verify rendering, not PMS accounting logic or production integration.</p><nav><a href="summary.json">Machine-readable results</a><a href="assets/signature.png">PNG signature</a><a href="assets/signature.jpg">JPEG signature</a><a href="assets/logo.png">Test logo</a></nav><div class="grid">${results.map(r => `<article class="${r.status}"><h2>${esc(r.id)}</h2><small>${esc(r.status)} · ${esc(r.imageMode)} · ${r.pages ?? '—'} pages</small>${r.pdf === 'passed' ? `<a href="${r.id}/report.pdf"><img loading="lazy" src="${r.id}/preview.png" alt="First page of ${esc(r.id)}"></a>` : ''}<p><a href="${r.id}/report.${r.format}">${r.format.toUpperCase()}</a> · <a href="${r.id}/data.json">Data</a> · <a href="${r.id}/attachments.json">Attachments</a> · <a href="${r.template}">Template</a></p>${r.error ? `<p>${esc(r.error)}</p>` : ''}</article>`).join('')}</div></main></html>`);
console.log(JSON.stringify({ ...summary, results: undefined, output: relative(process.cwd(), output) }));
if (summary.failed) process.exitCode = 1;
