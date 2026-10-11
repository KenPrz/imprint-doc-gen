import assert from 'node:assert/strict';
import { readFile } from 'node:fs/promises';
import { fileURLToPath } from 'node:url';
import { resolve, join } from 'node:path';

const root = resolve(fileURLToPath(new URL('..', import.meta.url)));
const endpoint = process.env.SMOKE_API_URL ?? 'http://127.0.0.1:3000';
const key = process.env.IMPRINT_API_KEY;
assert(key, 'IMPRINT_API_KEY is required');
assert.equal((await fetch(new URL('/health', endpoint))).status, 200);
assert.equal((await fetch(new URL('/v1/render', endpoint), { method: 'POST', headers: { 'content-type': 'application/json' }, body: '{}' })).status, 401);
const attachments = { red: (await readFile(join(root, 'corpus/attachments/red.png'))).toString('base64'), blue: (await readFile(join(root, 'corpus/attachments/blue.png'))).toString('base64') };
for (const level of [20, 61, 100]) {
  const dir = join(root, `corpus/docx/level-${String(level).padStart(3, '0')}`);
  const input = JSON.parse(await readFile(join(dir, 'input.json'), 'utf8'));
  const template = (await readFile(join(dir, 'template.docx'))).toString('base64');
  const variants = level === 61 ? [input, { ...input, items: [], groups: [] }] : [input];
  for (const [variant, data] of variants.entries()) for (const output of level === 100 ? ['docx'] : ['docx', 'pdf']) {
    const response: Response = await fetch(new URL('/v1/render', endpoint), {
      method: 'POST', headers: { authorization: `Bearer ${key}`, 'content-type': 'application/json' },
      body: JSON.stringify({ template, data, attachments, output }), signal: AbortSignal.timeout(90_000),
    });
    assert.equal(response.status, 200, await (response.status !== 200 ? response.text() : Promise.resolve('')));
    const bytes = Buffer.from(await response.arrayBuffer());
    assert(bytes.length > 100); assert(bytes.subarray(0, output === 'pdf' ? 5 : 2).equals(Buffer.from(output === 'pdf' ? '%PDF-' : 'PK')));
    console.log(`${output.toUpperCase()} level ${level}${variant ? ' empty loops' : ''} API render passed (${bytes.length} bytes)`);
  }
}
console.log('Health, authentication, scalar/image and nested-loop DOCX/PDF smoke checks passed.');
