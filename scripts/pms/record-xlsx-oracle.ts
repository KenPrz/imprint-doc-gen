import { readFile, writeFile } from 'node:fs/promises';
import { withArrayLength } from './variants.ts';
import { xlsxContent } from './xlsx-content.ts';
const endpoint = process.env.CARBONE_ORACLE_URL;
if (!endpoint) throw new Error('Set CARBONE_ORACLE_URL to an authorized existing installation');
const template = (await readFile('examples/pms/invoice.xlsx')).toString('base64');
const data = JSON.parse(await readFile('examples/pms/booking.json', 'utf8'));
const variants: Record<string, string[]> = {};
for (const count of [0, 1, 3]) {
  const response = await fetch(new URL('/render/template?download=true', endpoint), { method: 'POST', headers: { 'content-type': 'application/json' }, body: JSON.stringify({ template, data: withArrayLength(data, count), convertTo: 'xlsx' }), signal: AbortSignal.timeout(30000) });
  if (!response.ok) throw new Error(`Carbone returned ${response.status}`);
  variants[String(count)] = await xlsxContent(new Uint8Array(await response.arrayBuffer()));
}
await writeFile('tests/fixtures/carbone-xlsx-invoice.expected.json', JSON.stringify(variants, null, 2) + '\n');
