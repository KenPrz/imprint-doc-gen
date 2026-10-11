import { mkdir, writeFile, readFile } from 'node:fs/promises';
import { resolve, join } from 'node:path';
import { fileURLToPath } from 'node:url';
import { pack, png, hash, esc, type Format } from './common.ts';
import { recipe, inputData, bands } from './recipes.ts';
import { buildDocx } from './docx.ts';
import { buildXlsx } from './xlsx.ts';
import { validatePackage, assertGolden } from './inspect.ts';

const root = resolve(fileURLToPath(new URL('../..', import.meta.url)));
const destination = process.argv[2] ? resolve(process.argv[2]) : join(root, 'corpus');
const json = (value: unknown) => JSON.stringify(value, null, 2) + '\n';
await mkdir(join(destination, 'attachments'), { recursive: true });
await writeFile(join(destination, 'attachments/red.png'), png([220, 40, 40]));
await writeFile(join(destination, 'attachments/blue.png'), png([40, 80, 220]));
const cases = [];
for (const format of ['docx', 'xlsx'] as Format[]) {
  for (let level = 1; level <= 100; level++) {
    const r = { ...recipe(format, level), syntax: 'standard' }, data = inputData(level);
    const path = `${format}/level-${String(level).padStart(3, '0')}`;
    const dir = join(destination, path);
    const builder = format === 'docx' ? buildDocx : buildXlsx;
    const template = builder(level, data, false), expected = builder(level, data, true);
    validatePackage(template); validatePackage(expected); assertGolden(expected, format, level, data);
    const inputBytes = await pack(template, r.dimensions[8]);
    const expectedBytes = await pack(expected, r.dimensions[8]);
    await mkdir(dir, { recursive: true });
    const spec = {
      ...r, dialect: 'imprint-v1', expectedOutcome: 'render-success',
      input: 'input.json', template: `template.${format}`, expected: `expected.${format}`,
      attachments: { red: '../../attachments/red.png', blue: '../../attachments/blue.png' },
      preservedParts: [format === 'docx' ? 'word/styles.xml' : 'xl/styles.xml', 'customXml/fixture.xml'],
      sha256: { template: hash(inputBytes), expected: hash(expectedBytes) },
      bytes: { template: inputBytes.length, expected: expectedBytes.length },
      checks: ['ZIP integrity', 'XML well-formedness', 'OPC relationship integrity', 'no unresolved bindings', 'semantic content and layout comparison', 'image bytes and geometry', 'unchanged preservation parts'],
      runtimeStatus: 'not-run', visualStatus: 'not-run',
    };
    await writeFile(join(dir, `template.${format}`), inputBytes);
    await writeFile(join(dir, `expected.${format}`), expectedBytes);
    await writeFile(join(dir, 'input.json'), json(data));
    await writeFile(join(dir, 'case.json'), json(spec));
    cases.push({ id: r.id, format, level, tier: r.tier, step: r.step, title: r.title, path, sha256: spec.sha256 });
  }
  console.log(`Generated and checked 100 ${format.toUpperCase()} template/expected pairs`);
}
await writeFile(join(destination, 'manifest.json'), json({
  schemaVersion: 1, fixtureRevision: 2, dialect: 'imprint-v1', seed: 'imprint-doc-gen-corpus-v1-fixed',
  counts: { docx: 100, xlsx: 100, templates: 200, expectedDocuments: 200 },
  status: { fixturePackageChecks: 'passed', engineValidation: 'not-run', officeApplicationValidation: 'not-run', schemaValidation: 'not-run' },
  limitations: ['Synthetic fixtures; no claim of complete OOXML coverage.', 'Expected files and templates share packaging helpers; independent semantic checks and mutation tests reduce but do not eliminate correlated errors.', 'No Word/Excel/LibreOffice opening, full XSD validation or visual baseline approval has been completed.'],
  cases,
}));
await writeFile(join(destination, 'levels.csv'), 'id,format,level,tier,step,focus,path\n' + cases.map(c => [c.id, c.format, c.level, c.tier, c.step, c.title, c.path].map(v => `"${String(v).replace(/"/g, '""')}"`).join(',')).join('\n') + '\n');
await writeFile(join(destination, 'curriculum.json'), json(bands));
await writeFile(join(destination, 'index.html'), `<!doctype html>
<html lang="en"><meta charset="utf-8"><meta name="viewport" content="width=device-width,initial-scale=1">
<title>Imprint Doc Gen document test corpus</title>
<style>
body{font:16px/1.5 system-ui,sans-serif;color:#192c3d;background:#f5f7fa;margin:0}main{max-width:1120px;margin:auto;padding:40px 24px}h1{font-size:36px;margin-bottom:8px}p{max-width:850px}aside{background:#e4eef9;border-left:4px solid #2864aa;padding:16px;margin:24px 0}label{display:inline-block;margin:8px 24px 16px 0}select,input{font:inherit}table{width:100%;border-collapse:collapse;background:white}th,td{text-align:left;padding:12px;border-bottom:1px solid #dce3ea}th{background:#eaf0f5}a{color:#174e93}small{color:#4a5f71}tr[hidden]{display:none}nav a{margin-right:20px}.links{white-space:nowrap}footer{margin-top:24px}
</style><main>
<h1>Imprint Doc Gen document test corpus</h1>
<p>100 DOCX levels and 100 XLSX levels. Each level includes a template, JSON input and a reference output. Templates use standard {d.path} and [i]/[i+1] bindings. Ten cumulative bands increase structural and packaging complexity.</p>
<aside><strong>These are test fixtures, not engine results.</strong> Reference documents are generated by the fixture builders. Engine execution, full schema checks and Office visual approval remain separate gates.</aside>
<nav><a href="manifest.json">Manifest</a><a href="levels.csv">CSV inventory</a><a href="../docs/corpus-guide.txt">Usage guide</a><a href="../docs/corpus-roadmap.json">Future roadmap</a></nav>
<label>Format <select id="format"><option value="all">Both formats</option><option>docx</option><option>xlsx</option></select></label>
<label>Band <select id="band"><option value="all">All bands</option>${Array.from({ length: 10 }, (_, i) => `<option value="${i + 1}">${i * 10 + 1}–${i * 10 + 10}</option>`).join('')}</select></label>
<label>Up to level <input id="level" type="range" min="1" max="100" value="100"><output id="limit">100</output></label>
<p id="count" aria-live="polite">200 levels shown</p>
<table><thead><tr><th>Case</th><th>Focus</th><th>Files</th></tr></thead><tbody>
${cases.map(c => `<tr data-format="${c.format}" data-band="${c.tier}" data-level="${c.level}"><td><strong>${c.id}</strong><br><small>Band ${c.tier}, step ${c.step}</small></td><td>${esc(c.title)}</td><td class="links"><a href="${c.path}/template.${c.format}">Template</a> · <a href="${c.path}/expected.${c.format}">Expected</a> · <a href="${c.path}/input.json">Data</a> · <a href="${c.path}/case.json">Recipe</a></td></tr>`).join('\n')}
</tbody></table><footer><small>Package/content checks do not establish visual fidelity. See the manifest and usage guide for limitations.</small></footer></main>
<script>
const format=document.querySelector('#format'),band=document.querySelector('#band'),level=document.querySelector('#level');
function filter(){let count=0;document.querySelectorAll('tbody tr').forEach(row=>{row.hidden=(format.value!=='all'&&row.dataset.format!==format.value)||(band.value!=='all'&&row.dataset.band!==band.value)||Number(row.dataset.level)>Number(level.value);if(!row.hidden)count++});document.querySelector('#limit').value=level.value;document.querySelector('#count').textContent=count+' levels shown'}
[format,band,level].forEach(control=>control.addEventListener('input',filter));
</script></html>`);
// Prevent accidental generation into the repository root being mistaken for package installation.
await readFile(join(root, 'package.json'));
console.log(`Corpus written to ${destination}`);
