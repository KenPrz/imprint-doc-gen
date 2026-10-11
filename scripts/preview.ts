import { readFile, writeFile } from 'node:fs/promises';
import { resolve } from 'node:path';
import { RenderService } from '../src/render-service.ts';
const [template, data, target] = process.argv.slice(2);
if (!template || !data || !target || resolve(template) === resolve(target) || !process.env.GOTENBERG_URL)
  throw new Error('Usage: GOTENBERG_URL=... node scripts/preview.ts <template> <data.json> <new-output.pdf>');
const result = await new RenderService(1, process.env.GOTENBERG_URL).render(await readFile(template), {
  data: JSON.parse(await readFile(data, 'utf8')),
  output: 'pdf',
});
await writeFile(target, result.bytes, { flag: 'wx' });
console.log(`Preview: ${target}`);
