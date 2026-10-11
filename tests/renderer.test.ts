import test from 'node:test';
import assert from 'node:assert/strict';
import { render, compile, RenderError } from '../src/index.ts';
import { ns, parseXml, descendants } from '../src/xml.ts';
import { readPackage, writePackage, limits } from '../src/package.ts';
import { png } from '../scripts/corpus/common.ts';
import { createApp } from '../src/app.ts';
import { renderInWorker } from '../src/jobs.ts';

async function template(body: string, extra: Record<string, string | Buffer> = {}) {
  const parts = new Map<string, Buffer>(Object.entries({
    '[Content_Types].xml': `<Types xmlns="${ns.contentTypes}"><Default Extension="xml" ContentType="application/xml"/><Default Extension="rels" ContentType="application/vnd.openxmlformats-package.relationships+xml"/><Default Extension="png" ContentType="image/png"/><Override PartName="/word/document.xml" ContentType="application/vnd.openxmlformats-officedocument.wordprocessingml.document.main+xml"/></Types>`,
    '_rels/.rels': `<Relationships xmlns="${ns.packageRelationship}"><Relationship Id="main" Type="${ns.relationship}/officeDocument" Target="word/document.xml"/></Relationships>`,
    'word/document.xml': `<?xml version="1.0"?><word:document xmlns:word="${ns.word}" xmlns:r="${ns.relationship}"><word:body>${body}<word:sectPr/></word:body></word:document>`,
    'customXml/keep.xml': '<keep xmlns="urn:independent-fixture">preserve byte for byte</keep>', ...extra,
  }).map(([name, value]) => [name, Buffer.from(value)]));
  return writePackage(parts);
}
const paragraph = (text: string) => `<word:p><word:r><word:t>${text}</word:t></word:r></word:p>`;
async function mainText(bytes: Uint8Array) {
  const parts = await readPackage(bytes);
  return descendants(parseXml(parts.get('word/document.xml')!.toString(), 'test'), ns.word, 't').map(n => n.textContent).join('');
}
test('source-preserving replacement joins fragmented placeholders and keeps surrounding formatting', async () => {
  const input = await template('<word:p><word:r><word:rPr><word:b/></word:rPr><word:t>Dear {{cus</word:t></word:r><word:r><word:rPr><word:i/></word:rPr><word:t>tomer.name}} — done</word:t></word:r></word:p>');
  const output = await render(input, { customer: { name: '  Zoë & 山田 <💡>  ' } });
  assert.equal(await mainText(output), 'Dear   Zoë & 山田 <💡>   — done');
  const parts = await readPackage(output), source = parts.get('word/document.xml')!.toString();
  assert(source.includes('<word:rPr><word:b/></word:rPr>'));
  assert(source.includes('<word:rPr><word:i/></word:rPr>'));
  assert(source.includes('xml:space="preserve"'));
  assert.deepEqual(parts.get('customXml/keep.xml'), (await readPackage(input)).get('customXml/keep.xml'));
});
test('multiple placeholders, false, zero and explicit null are distinct from missing data', async () => {
  const input = await template(paragraph('{{zero}}|{{flag}}|{{empty}}|{{zero}}'));
  assert.equal(await mainText(await render(input, { zero: 0, flag: false, empty: null })), '0|false||0');
  await assert.rejects(render(input, { zero: 0, flag: false }), (e: unknown) => e instanceof RenderError && e.code === 'MISSING_DATA');
});
test('compiled templates are reusable concurrently without data leaking between renders', async () => {
  const compiled = await compile(await template(paragraph('{{name}}')));
  const results = await Promise.all(['A', 'B', 'C'].map(name => compiled.render({ name })));
  assert.deepEqual(await Promise.all(results.map(mainText)), ['A', 'B', 'C']);
});
test('malformed tags, arbitrary expressions and unclosed loops produce explicit errors', async () => {
  for (const [tag, code] of [['{{name', 'TEMPLATE_SYNTAX'], ['{{name + 1}}', 'TEMPLATE_EXPRESSION'], ['{{constructor.name}}', 'TEMPLATE_EXPRESSION'], ['{{#each items}}', 'TEMPLATE_SYNTAX']]) {
    await assert.rejects(compile(await template(paragraph(tag!))), (e: unknown) => e instanceof RenderError && e.code === code);
  }
});
test('accessor data is rejected without executing the accessor', async () => {
  let accessed = false;
  const data = Object.defineProperty({}, 'name', { get() { accessed = true; return 'bad'; }, enumerable: true });
  await assert.rejects(render(await template(paragraph('{{name}}')), data), (e: unknown) => e instanceof RenderError && e.code === 'INVALID_DATA');
  assert.equal(accessed, false);
});
test('CDATA text can contain a placeholder and escaped values stay literal', async () => {
  const input = await template(paragraph('<![CDATA[{{name}} &amp;]]>'));
  assert.equal(await mainText(await render(input, { name: '& <literal>' })), '& <literal> &amp;');
});
test('milestone limitations are reported for multiline text and oversized packages', async () => {
  await assert.rejects(render(await template(paragraph('{{name}}')), { name: 'a\nb' }), (e: unknown) => e instanceof RenderError && e.code === 'UNSUPPORTED_DATA');
  await assert.rejects(compile(Buffer.alloc(limits.archive + 1)), (e: unknown) => e instanceof RenderError && e.code === 'LIMIT_EXCEEDED');
});
test('DTD and archive traversal are rejected', async () => {
  await assert.rejects(compile(await template(paragraph('{{name}}'), { 'word/document.xml': '<!DOCTYPE x [<!ENTITY x "bad">]><x>&x;</x>' })), (e: unknown) => e instanceof RenderError && e.code === 'INVALID_XML');
  await assert.rejects(compile(await template(paragraph('okay'), { '../escape.xml': '<x/>' })), (e: unknown) => e instanceof RenderError && e.code === 'PACKAGE_PATH');
});
test('strict OOXML is rejected explicitly rather than returned with unresolved placeholders', async () => {
  const parts = await readPackage(await template(paragraph('{{name}}')));
  parts.set('word/document.xml', Buffer.from(parts.get('word/document.xml')!.toString().replace(ns.word, 'http://purl.oclc.org/ooxml/wordprocessingml/main')));
  await assert.rejects(compile(await writePackage(parts)), (e: unknown) => e instanceof RenderError && e.code === 'UNSUPPORTED_FORMAT');
});
test('changing one image leaves an unbound picture sharing the original relationship unchanged', async () => {
  const image = (id: number, description: string) => `<word:p><word:r><word:drawing><wp:inline xmlns:wp="${ns.wordDrawing}" xmlns:a="${ns.drawing}"><wp:extent cx="914400" cy="609600"/><wp:docPr id="${id}" name="Image ${id}" descr="${description}"/><a:graphic><a:graphicData uri="urn:test"><a:blip r:embed="shared"/></a:graphicData></a:graphic></wp:inline></word:drawing></word:r></word:p>`;
  const gray = png([100, 100, 100]), red = png([255, 0, 0]);
  const input = await template(image(1, '{{image:logo}}') + image(2, 'Keep this image'), {
    'word/_rels/document.xml.rels': `<Relationships xmlns="${ns.packageRelationship}"><Relationship Id="shared" Type="${ns.relationship}/image" Target="media/original.png"/></Relationships>`,
    'word/media/original.png': gray,
  });
  const output = await readPackage(await render(input, {}, { attachments: { logo: red } }));
  assert.deepEqual(output.get('word/media/original.png'), gray);
  const document = parseXml(output.get('word/document.xml')!.toString(), 'test');
  const blips = descendants(document, ns.drawing, 'blip');
  assert.notEqual(blips[0]!.getAttributeNS(ns.relationship, 'embed'), 'shared');
  assert.equal(blips[1]!.getAttributeNS(ns.relationship, 'embed'), 'shared');
  assert.equal(descendants(document, ns.wordDrawing, 'extent')[0]!.getAttribute('cx'), '914400');
  assert([...output.entries()].some(([name, bytes]) => name.includes('imprint-') && bytes.equals(red)));
  await assert.rejects(render(input, {}), (e: unknown) => e instanceof RenderError && ['MISSING_DATA', 'MISSING_ATTACHMENT'].includes(e.code));
  await assert.rejects(render(input, {}, { attachments: { logo: Buffer.from('not an image') } }), (e: unknown) => e instanceof RenderError && e.code === 'INVALID_IMAGE');
});
test('worker produces a real render and enforces a wall-clock deadline', async () => {
  const input = await template(paragraph('{{name}}'));
  assert.equal(await mainText(await renderInWorker(input, { name: 'worker' }, {})), 'worker');
  await assert.rejects(renderInWorker(input, { name: 'late' }, {}, 1), (e: unknown) => e instanceof RenderError && e.code === 'RENDER_TIMEOUT');
});
test('API validates authentication, schema, base64 and successful binary output', async () => {
  const key = 'independent-integration-test-key';
  const app = createApp({ apiKey: key });
  try {
    assert.equal((await app.inject({ url: '/health' })).statusCode, 200);
    assert.equal((await app.inject({ method: 'POST', url: '/v1/render', payload: {} })).statusCode, 401);
    assert.equal((await app.inject({ method: 'POST', url: '/v1/render', headers: { authorization: `Bearer ${key}` }, payload: {} })).statusCode, 400);
    const response = await app.inject({ method: 'POST', url: '/v1/render', headers: { authorization: `Bearer ${key}` }, payload: { template: (await template(paragraph('{{name}}'))).toString('base64'), data: { name: 'API' } } });
    assert.equal(response.statusCode, 200, response.body);
    assert.equal(await mainText(response.rawPayload), 'API');
    assert.equal((await app.inject({ method: 'POST', url: '/v1/render', headers: { authorization: `Bearer ${key}` }, payload: { template: '???', data: {} } })).statusCode, 422);
  } finally { await app.close(); }
});
test('production configuration cannot silently start without an API key', () => {
  assert.throws(() => createApp({ apiKey: '' }), /IMPRINT_API_KEY/);
});

test('PDF converter errors return JSON with 502 instead of inheriting the PDF content type', async () => {
  const { createServer } = await import('node:http');
  const server = createServer((_request, response) => { response.writeHead(503); response.end('busy'); });
  await new Promise<void>(resolve => server.listen(0, '127.0.0.1', resolve));
  const address = server.address() as { port: number };
  const key = 'independent-integration-test-key';
  const app = createApp({ apiKey: key, converterUrl: `http://127.0.0.1:${address.port}` });
  try {
    const response = await app.inject({ method: 'POST', url: '/v1/render', headers: { authorization: `Bearer ${key}` }, payload: { template: (await template(paragraph('{{name}}'))).toString('base64'), data: { name: 'PDF' }, output: 'pdf' } });
    assert.equal(response.statusCode, 502, response.body);
    assert.match(response.headers['content-type']!, /application\/json/);
    assert.equal(response.json().error, 'CONVERSION_FAILED');
  } finally { await app.close(); await new Promise<void>(resolve => server.close(() => resolve())); }
});
