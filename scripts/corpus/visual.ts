import assert from 'node:assert/strict';
import { readFile, writeFile, mkdir } from 'node:fs/promises';
import { resolve, join } from 'node:path';
import { fileURLToPath } from 'node:url';
import sharp from 'sharp';
import { convertToPdf } from '../../src/converter.ts';

const root = resolve(fileURLToPath(new URL('../../corpus', import.meta.url)));
const destination = join(root, 'results/visual');
const levels = [1, 20, 50];
await mkdir(destination, { recursive: true });
if (!process.argv.includes('--compare')) {
  for (const level of levels) {
    const id = String(level).padStart(3, '0');
    const actual = await readFile(join(root, 'actual/docx', `level-${id}.docx`));
    await writeFile(join(destination, `${id}-actual.pdf`), await convertToPdf(actual, process.env.GOTENBERG_URL ?? 'http://gotenberg:3000'));
    await writeFile(join(destination, `${id}-expected.pdf`), await readFile(join(root, 'results/office', `docx-${id}-expected.pdf`)));
  }
  console.log('Converted actual DOCX levels 1, 20, 50. Rasterize every PDF page with pdftoppm -r 96 -png, then run this script with --compare.');
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
      for (let i = 0; i < a.data.length; i += 4) if (!a.data.subarray(i, i + 4).equals(b.data.subarray(i, i + 4))) differentPixels++;
      results.push({ level, page: page + 1, pixels: a.info.width * a.info.height, differentPixels });
    }
  }
  await writeFile(join(root, 'results/visual-validation.json'), JSON.stringify({ scope: 'DOCX levels 1, 20, 50; actual versus expected PDF rasterized at 96 DPI using the same converter and fonts', results }, null, 2) + '\n');
  console.log(JSON.stringify(results));
  if (results.some(r => r.differentPixels)) process.exitCode = 1;
}
