import test from 'node:test';
import assert from 'node:assert/strict';
import { readFile } from 'node:fs/promises';
import { inspectTemplate } from '../src/inspect.ts';
import { readPackage, writePackage } from '../src/package.ts';
import { createApp } from '../src/app.ts';
test('Inspection finds fields, repeats and pictures without generating a document', async () => {
  const input = await readFile('examples/invoice/invoice.docx'),
    result = await inspectTemplate(input);
  assert(result.valid);
  assert(result.fields.some((f) => f.path === 'company'));
  assert(result.fields.some((f) => f.kind === 'image' && f.path === 'logo'));
  assert(result.fields.some((f) => f.kind === 'array'));
  const parts = await readPackage(input);
  parts.set(
    'word/document.xml',
    Buffer.from(parts.get('word/document.xml')!.toString().replace('{d.company}', '{d.company:required()}')),
  );
  assert((await inspectTemplate(await writePackage(parts))).fields.some((f) => f.required));
  parts.set(
    'word/document.xml',
    Buffer.from(parts.get('word/document.xml')!.toString().replace('required()', 'execute()')),
  );
  const bad = await inspectTemplate(await writePackage(parts));
  assert(!bad.valid);
  assert.match(bad.diagnostics[0]!.location!, /paragraph/);
});
test('Inspection API uses the bounded worker and enforces authentication', async () => {
  const key = 'inspection-test-administrator-key',
    app = createApp({ apiKey: key });
  try {
    const payload = { template: (await readFile('examples/invoice/invoice.xlsx')).toString('base64') };
    assert.equal((await app.inject({ method: 'POST', url: '/v1/inspect', payload })).statusCode, 401);
    const r = await app.inject({
      method: 'POST',
      url: '/v1/inspect',
      payload,
      headers: { authorization: `Bearer ${key}` },
    });
    assert.equal(r.statusCode, 200);
    assert(r.json().valid);
    assert(r.json().fields.some((f: any) => f.location === 'cell A1'));
  } finally {
    await app.close();
  }
});
