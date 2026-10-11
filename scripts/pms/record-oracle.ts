// Explicit reference acquisition only: CI consumes checked-in oracle snapshots.
import { readFile, readdir, mkdir, writeFile } from 'node:fs/promises';
import { withArrayLength } from './variants.ts';
import { snapshot } from './snapshot.ts';
const endpoint = process.env.CARBONE_ORACLE_URL;
if (!endpoint) throw new Error('Set CARBONE_ORACLE_URL to your existing Carbone service to record independent references');
await mkdir('corpus/results/pms/oracle', { recursive: true });
for (const file of (await readdir('tests/fixtures/pms')).filter(file => file.endsWith('-template.docx')).sort()) {
  const name = file.replace('-template.docx', '');
  const data = JSON.parse(await readFile(`tests/fixtures/pms/${name}.json`, 'utf8'));
  const response = await fetch(new URL('/render/template?download=true', endpoint), { method: 'POST', headers: { 'content-type': 'application/json' },
    body: JSON.stringify({ template: (await readFile(`tests/fixtures/pms/${file}`)).toString('base64'), data, convertTo: 'docx' }), signal: AbortSignal.timeout(30_000) });
  if (!response.ok) throw new Error(`Carbone returned ${response.status} for ${file}`);
  const bytes = Buffer.from(await response.arrayBuffer());
  await writeFile(`corpus/results/pms/oracle/${file}`, bytes);
  await writeFile(`tests/fixtures/pms/${name}.expected.json`, JSON.stringify(await snapshot(bytes), null, 2) + '\n');
  console.log(`Recorded independent oracle: ${file}`);
}

for (const name of ['folio_statement', 'task_sheet', 'trial_balance']) {
  const data = JSON.parse(await readFile(`tests/fixtures/pms/${name}.json`, 'utf8'));
  const template = (await readFile(`tests/fixtures/pms/${name}-template.docx`)).toString('base64');
  const variants: Record<string, unknown> = {};
  for (const count of [0, 1, 3]) {
    const response = await fetch(new URL('/render/template?download=true', endpoint), { method: 'POST', headers: { 'content-type': 'application/json' }, body: JSON.stringify({ template, data: withArrayLength(data, count), convertTo: 'docx' }), signal: AbortSignal.timeout(30_000) });
    if (!response.ok) throw new Error(`Carbone returned ${response.status} for ${name} variant ${count}`);
    variants[String(count)] = await snapshot(new Uint8Array(await response.arrayBuffer()));
  }
  await writeFile(`tests/fixtures/pms/${name}.variants.expected.json`, JSON.stringify(variants, null, 2) + '\n');
}
