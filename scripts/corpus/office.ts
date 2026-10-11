import { readFile, writeFile, mkdir } from 'node:fs/promises';
import { resolve, join } from 'node:path';
import { fileURLToPath } from 'node:url';
import { convertToPdf } from '../../src/converter.ts';

const root = resolve(fileURLToPath(new URL('../../corpus', import.meta.url)));
const endpoint = process.env.GOTENBERG_URL ?? 'http://gotenberg:3000';
const levels = process.argv.includes('--all') ? Array.from({ length: 100 }, (_, i) => i + 1) : [1, 10, 11, 20, 21, 30, 40, 50, 60, 70, 80, 90, 100];
const onlyFormat = process.argv.find(arg => arg.startsWith('--format='))?.split('=')[1];
if (onlyFormat && !['docx', 'xlsx'].includes(onlyFormat)) throw new Error('Expected --format=docx or --format=xlsx');
const reportPath = join(root, 'results/office-validation.json');
const previous = onlyFormat ? JSON.parse(await readFile(reportPath, 'utf8').catch(() => '{\"results\":[]}')).results : [];
const results: { id: string; status: string; pdfBytes?: number; error?: string }[] = previous.filter((r: { id: string }) => !r.id.startsWith(`${onlyFormat}-`));
await mkdir(join(root, 'results/office'), { recursive: true });
for (const format of onlyFormat ? [onlyFormat] : ['docx', 'xlsx']) for (const level of levels) for (const kind of ['template', 'expected']) {
  const id = `${format}-${String(level).padStart(3, '0')}-${kind}`;
  try {
    const input = await readFile(join(root, format, `level-${String(level).padStart(3, '0')}`, `${kind}.${format}`));
    const pdf = await convertToPdf(input, endpoint, `${id}.${format}`);
    await writeFile(join(root, 'results/office', `${id}.pdf`), pdf);
    results.push({ id, status: 'converted', pdfBytes: pdf.length });
  } catch (error) { results.push({ id, status: 'failed', error: (error as Error).message }); }
  console.log(`${id}: ${results.at(-1)!.status}`);
  await writeFile(reportPath, JSON.stringify({ converter: 'Gotenberg 8.37.0 / LibreOffice', scope: onlyFormat ?? (levels.length === 100 ? 'all' : 'representative'), interpretation: 'Conversion success does not prove visual equivalence or repair-free Office opening.', results }, null, 2) + '\n');
}
await writeFile(join(root, 'results/office-validation.json'), JSON.stringify({ converter: 'Gotenberg 8.37.0 / LibreOffice', scope: levels.length === 100 ? 'all' : 'representative levels', interpretation: 'Successful PDF conversion establishes that LibreOffice could process the input; it does not prove repair-free Office opening or visual equivalence.', results }, null, 2) + '\n');
if (results.some(r => r.status === 'failed')) process.exitCode = 1;
