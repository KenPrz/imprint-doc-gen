import assert from 'node:assert/strict';
import { readFile, writeFile, mkdir } from 'node:fs/promises';
import { resolve, join } from 'node:path';
import { fileURLToPath } from 'node:url';
import sharp from 'sharp';
import { convertToPdf } from '../../src/converter.ts';

const root = resolve(fileURLToPath(new URL('../../corpus', import.meta.url)));
const format = process.argv.find(arg => arg.startsWith('--format='))?.slice(9) ?? 'docx';
assert(['docx', 'xlsx'].includes(format), 'Expected --format=docx or --format=xlsx');
const destination = join(root, format === 'docx' ? 'results/visual' : 'results/visual-xlsx');
const levels = (process.argv.find(arg => arg.startsWith('--levels='))?.slice(9) ?? '1,20,50').split(',').map(Number);
assert(levels.length && levels.every(level => Number.isInteger(level) && level >= 1 && level <= 100), 'Expected levels 1..100');
await mkdir(destination, { recursive: true });
if (!process.argv.includes('--compare')) {
  for (const level of levels) {
    const id = String(level).padStart(3, '0');
    const { readdir, rm } = await import('node:fs/promises');
    for (const file of await readdir(destination)) if (file.startsWith(`${id}-`) && /\.(png|pdf)$/.test(file)) await rm(join(destination, file));
    const actual = await readFile(join(root, 'actual', format, `level-${id}.${format}`));
    await writeFile(join(destination, `${id}-actual.pdf`), await convertToPdf(actual, process.env.GOTENBERG_URL ?? 'http://gotenberg:3000', `document.${format}`));
    await writeFile(join(destination, `${id}-expected.pdf`), await convertToPdf(await readFile(join(root, format, `level-${id}`, `expected.${format}`)), process.env.GOTENBERG_URL ?? 'http://gotenberg:3000', `document.${format}`));
  }
  console.log(`Converted actual and reference ${format.toUpperCase()} levels ${levels.join(', ')}. Rasterize every PDF page with pdftoppm -r 96 -png, then run this script with --compare and the same --format and --levels.`);
} else {
  const { readdir } = await import('node:fs/promises');
  const files = await readdir(destination);
  const results = [];
  for (const level of levels) {
    const id = String(level).padStart(3, '0');
    const pages = (kind: string) => files.filter(f => f.startsWith(`${id}-${kind}-`) && f.endsWith('.png')).sort();
    const actual = pages('actual'), expected = pages('expected');
    assert(actual.length > 0 && actual.length === expected.length, `Page count mismatch for ${id}`);
    for (let page = 0; page < actual.length; page++) {
      const a = await sharp(join(destination, actual[page]!)).ensureAlpha().raw().toBuffer({ resolveWithObject: true });
      const b = await sharp(join(destination, expected[page]!)).ensureAlpha().raw().toBuffer({ resolveWithObject: true });
      assert.deepEqual(a.info, b.info, `Page geometry mismatch for ${id}`);
      let differentPixels = 0;
      for (let i = 0; i < a.data.length; i += 4) if (a.data[i] !== b.data[i] || a.data[i + 1] !== b.data[i + 1] || a.data[i + 2] !== b.data[i + 2] || a.data[i + 3] !== b.data[i + 3]) differentPixels++;
      results.push({ level, page: page + 1, pixels: a.info.width * a.info.height, differentPixels });
    }
  }
  await writeFile(join(root, format === 'docx' ? 'results/visual-validation.json' : 'results/visual-xlsx-validation.json'), JSON.stringify({ scope: `${format.toUpperCase()} levels ${levels.join(', ')}; actual versus expected PDF rasterized at 96 DPI using the same converter and fonts`, results }, null, 2) + '\n');
  console.log(JSON.stringify(results));
  if (results.some(r => r.differentPixels)) process.exitCode = 1;
}
