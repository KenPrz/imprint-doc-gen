import test from 'node:test';
import assert from 'node:assert/strict';
import { compile, render, RenderError } from '../src/index.ts';
import { readPackage, writePackage, resolvePart, relsPath } from '../src/package.ts';
import { ns, parseXml, descendants } from '../src/xml.ts';
import { png } from '../scripts/corpus/common.ts';
import { createApp } from '../src/app.ts';

const p = (text: string) => `<w:p><w:r><w:t>${text}</w:t></w:r></w:p>`;
const cell = (body: string) => `<w:tc><w:tcPr><w:tcW w:w="2000" w:type="dxa"/></w:tcPr>${body}</w:tc>`;
const row = (body: string) => `<w:tbl><w:tblPr/><w:tblGrid><w:gridCol w:w="2000"/></w:tblGrid><w:tr>${cell(body)}</w:tr></w:tbl>`;
const block = (body: string, path = 'items') => p(`{{#each ${path}}}`) + body + p('{{/each}}');
const picture = (id: number, binding = '{{image:image}}') => `<w:p><w:r><w:drawing><wp:inline><wp:extent cx="914400" cy="609600"/><wp:docPr id="${id}" name="picture" descr="${binding}"/><a:graphic><a:graphicData uri="urn:test"><a:blip r:embed="shared"/></a:graphicData></a:graphic></wp:inline></w:drawing></w:r></w:p>`;
const namespaces = `xmlns:w="${ns.word}" xmlns:r="${ns.relationship}" xmlns:wp="${ns.wordDrawing}" xmlns:a="${ns.drawing}"`;
async function template(body: string, extra: Record<string, string | Buffer> = {}) {
  const entries: Record<string, string | Buffer> = {
    '[Content_Types].xml': `<Types xmlns="${ns.contentTypes}"><Default Extension="xml" ContentType="application/xml"/><Default Extension="rels" ContentType="application/vnd.openxmlformats-package.relationships+xml"/><Default Extension="png" ContentType="image/png"/><Override PartName="/word/document.xml" ContentType="application/vnd.openxmlformats-officedocument.wordprocessingml.document.main+xml"/></Types>`,
    '_rels/.rels': `<Relationships xmlns="${ns.packageRelationship}"><Relationship Id="main" Type="${ns.relationship}/officeDocument" Target="word/document.xml"/></Relationships>`,
    'word/document.xml': `<w:document ${namespaces}><w:body>${body}<w:sectPr/></w:body></w:document>`,
    'word/_rels/document.xml.rels': `<Relationships xmlns="${ns.packageRelationship}"><Relationship Id="shared" Type="${ns.relationship}/image" Target="media/original.png"/></Relationships>`,
    'word/media/original.png': png([80, 80, 80]), 'customXml/keep.xml': '<keep>exact bytes</keep>', ...extra,
  };
  return writePackage(new Map(Object.entries(entries).map(([path, bytes]) => [path, Buffer.from(bytes)])));
}
async function document(bytes: Uint8Array, part = 'word/document.xml') {
  return parseXml((await readPackage(bytes)).get(part)!.toString(), part);
}
async function texts(bytes: Uint8Array, part?: string) {
  return descendants(await document(bytes, part), ns.word, 'p').map(p => descendants(p, ns.word, 't').map(t => t.textContent).join(''));
}
const errorCode = (code: string) => (error: unknown) => error instanceof RenderError && error.code === code;

test('block loops remove marker paragraphs, preserve surrounding content and empty arrays', async () => {
  const input = await template(p('before') + block(p('{{name}}')) + p('after'));
  assert.deepEqual(await texts(await render(input, { items: [{ name: 'A' }, { name: 'B' }] })), ['before', 'A', 'B', 'after']);
  assert.deepEqual(await texts(await render(input, { items: [] })), ['before', 'after']);
});

test('row loops span cells, retain row/cell properties and handle fragmented markers', async () => {
  const input = await template('<w:tbl><w:tblPr/><w:tblGrid><w:gridCol w:w="2000"/><w:gridCol w:w="2000"/></w:tblGrid><w:tr><w:trPr><w:cantSplit/></w:trPr>' +
    cell('<w:p><w:r><w:t>{{#ea</w:t></w:r><w:r><w:t>ch items}}{{na</w:t></w:r><w:r><w:rPr><w:b/></w:rPr><w:t>me}}</w:t></w:r></w:p>') + cell(p('{{quantity}}{{/each}}')) + '</w:tr></w:tbl>');
  const output = await render(input, { items: [{ name: 'A', quantity: 0 }, { name: 'B', quantity: false }] });
  assert.deepEqual(await texts(output), ['A', '0', 'B', 'false']);
  const xml = await document(output);
  assert.equal(descendants(xml, ns.word, 'tr').length, 2);
  assert.equal(descendants(xml, ns.word, 'cantSplit').length, 2);
  assert.equal(descendants(xml, ns.word, 'tcW').length, 4);
  assert.equal(descendants(await document(await render(input, { items: [] })), ns.word, 'tr').length, 0);
});

test('nested scopes require explicit parent paths and support grandparent paths', async () => {
  const input = await template(block(p('{{name}}') + block(p('{{../../title}} / {{../name}} / {{name}}')), 'groups'));
  const data = { title: 'root', groups: [{ name: 'A', items: [{ name: 'one' }, { name: 'two' }] }, { name: 'B', items: [] }] };
  assert.deepEqual(await texts(await render(input, data)), ['A', 'root / A / one', 'root / A / two', 'B']);
  await assert.rejects(render(await template(block(p('{{title}}'))), { title: 'root', items: [{}] }), errorCode('MISSING_DATA'));
  await assert.rejects(render(await template(p('{{../title}}')), { title: 'root' }), errorCode('MISSING_DATA'));
});

test('paragraph blocks can nest inside a repeated row', async () => {
  const input = await template(row(p('{{#each groups}}{{name}}') + block(p('{{../name}}:{{name}}')) + p('{{/each}}')));
  const output = await render(input, { groups: [{ name: 'A', items: [{ name: 'one' }, { name: 'two' }] }, { name: 'B', items: [] }] });
  assert.deepEqual(await texts(output), ['A', 'A:one', 'A:two', '', 'B', '']);
});

test('per-item image keys win over global attachment keys and unbound media stays unchanged', async () => {
  const input = await template(picture(1, 'Unbound original') + block(p('{{name}}') + picture(2)));
  const red = png([255, 0, 0]), blue = png([0, 0, 255]);
  const output = await render(input, { items: [{ name: 'A', image: 'red' }, { name: 'B', image: 'blue' }, { name: 'C', image: 'red' }] }, { attachments: { red, blue, image: png([0, 255, 0]) } });
  const parts = await readPackage(output), xml = await document(output);
  const relationships = descendants(parseXml(parts.get(relsPath('word/document.xml'))!.toString(), 'rels'), ns.packageRelationship, 'Relationship');
  const images = descendants(xml, ns.drawing, 'blip').map(blip => {
    const id = blip.getAttributeNS(ns.relationship, 'embed');
    const link = relationships.find(r => r.getAttribute('Id') === id)!;
    return parts.get(resolvePart('word/document.xml', link.getAttribute('Target')!));
  });
  assert.deepEqual(images, [png([80, 80, 80]), red, blue, red]);
  const ids = descendants(xml, ns.wordDrawing, 'docPr').map(n => n.getAttribute('id'));
  assert.equal(ids[0], '1'); assert.equal(new Set(ids).size, 4);
  assert.deepEqual(parts.get('customXml/keep.xml'), Buffer.from('<keep>exact bytes</keep>'));
  assert.equal([...parts.keys()].filter(p => p.includes('/imprint-')).length, 2, 'identical image bytes share a media part');
});

test('empty loops need no per-item attachments and a missing attachment fails atomically', async () => {
  const compiled = await compile(await template(block(picture(1))));
  assert.equal(descendants(await document(await compiled.render({ items: [] })), ns.wordDrawing, 'docPr').length, 0);
  await assert.rejects(compiled.render({ items: [{ image: 'absent' }] }), errorCode('MISSING_ATTACHMENT'));
  assert.equal(descendants(await document(await compiled.render({ items: [{ image: 'red' }] }, { attachments: { red: png([255, 0, 0]) } })), ns.wordDrawing, 'docPr').length, 1);
});

test('root attachment lookup remains compatible and parent image keys work inside loops', async () => {
  const input = await template(block(picture(1, '{{image:../logo}}')));
  const output = await render(input, { logo: 'red', items: [{}] }, { attachments: { red: png([255, 0, 0]) } });
  assert.equal(descendants(await document(output), ns.wordDrawing, 'docPr')[0]!.getAttribute('descr'), 'Attachment red');
});

test('repeated static pictures and content controls receive fresh identities', async () => {
  const control = '<w:sdt><w:sdtPr><w:id w:val="8"/></w:sdtPr><w:sdtContent>' + p('{{name}}') + '</w:sdtContent></w:sdt>';
  const output = await render(await template(block(picture(2, 'Static picture') + control)), { items: [{ name: 'A' }, { name: 'B' }] });
  const xml = await document(output);
  assert.equal(new Set(descendants(xml, ns.wordDrawing, 'docPr').map(n => n.getAttribute('id'))).size, 2);
  assert.equal(new Set(descendants(xml, ns.word, 'id').map(n => n.getAttributeNS(ns.word, 'val'))).size, 2);
});

test('compiled loop templates can render concurrently without scope, image or identity leakage', async () => {
  const compiled = await compile(await template(block(p('{{name}}') + picture(1))));
  const attachments = { red: png([255, 0, 0]), blue: png([0, 0, 255]) };
  const outputs = await Promise.all(['red', 'blue', 'red'].map(image => compiled.render({ items: [{ name: image, image }] }, { attachments })));
  assert.deepEqual(await Promise.all(outputs.map(o => texts(o))), [['red', ''], ['blue', ''], ['red', '']]);
  assert.deepEqual(outputs[0], outputs[2]);
});

test('data that looks like template syntax is never parsed again', async () => {
  const output = await render(await template(block(p('{{name}}'))), { items: [{ name: '{{#each evil}} & {{value}}' }] });
  assert.deepEqual(await texts(output), ['{{#each evil}} & {{value}}']);
});

test('loop arrays must contain objects; missing, null, scalar and sparse inputs fail explicitly', async () => {
  const input = await template(block(p('{{name}}')));
  await assert.rejects(render(input, {}), errorCode('MISSING_DATA'));
  for (const items of [null, {}, 0, [null], [1], [[]], new Array(2)]) await assert.rejects(render(input, { items }), errorCode('INVALID_DATA'));
});

test('unbalanced markers and invalid paths fail at compile time even for an empty loop', async () => {
  for (const body of [p('{{/each}}'), p('{{#each items}}'), block(p('{{constructor.value}}')), block(p('{{../__proto__.x}}')), block(p('{{broken'))]) {
    await assert.rejects(compile(await template(body)), (error: unknown) => error instanceof RenderError && ['TEMPLATE_SYNTAX', 'TEMPLATE_EXPRESSION'].includes(error.code));
  }
});

test('inline blocks, multi-row boundaries and duplicate row selections are rejected', async () => {
  const malformed = [
    p('prefix {{#each items}}') + p('{{name}}') + p('{{/each}}'),
    '<w:tbl><w:tblPr/><w:tblGrid/><w:tr>' + cell(p('{{#each items}}')) + '</w:tr><w:tr>' + cell(p('{{/each}}')) + '</w:tr></w:tbl>',
    row(p('{{#each items}}{{#each children}}{{name}}{{/each}}{{/each}}')),
    p('{{#each items}}{{name}}{{/each}}'),
  ];
  for (const body of malformed) await assert.rejects(compile(await template(body)), errorCode('TEMPLATE_STRUCTURE'));
});

test('unsupported reference-bearing structures inside repeats fail instead of being duplicated', async () => {
  for (const content of ['<w:bookmarkStart w:id="1" w:name="bookmark"/>', '<w:vMerge w:val="continue"/>']) {
    await assert.rejects(compile(await template(block(p('value') + content))), errorCode('UNSUPPORTED_TEMPLATE'));
  }
});

test('loop depth and cumulative expansion have explicit limits', async () => {
  let body = p('{{name}}'); for (let i = 0; i < 17; i++) body = block(body);
  await assert.rejects(compile(await template(body)), errorCode('LIMIT_EXCEEDED'));
  const input = await template(block(block(p('constant'), '../items')));
  await assert.rejects(render(input, { items: Array.from({ length: 101 }, () => ({})) }), errorCode('LIMIT_EXCEEDED'));
});

test('repeated replacements cannot exceed the output byte budget', async () => {
  const input = await template(block(p('{{text}}')));
  await assert.rejects(render(input, { items: Array.from({ length: 10 }, () => ({ text: 'x'.repeat(900_000) })) }), errorCode('LIMIT_EXCEEDED'));
});

test('looped header stories keep part-local image relationships and global drawing IDs', async () => {
  const header = `<w:hdr ${namespaces}>${block(p('{{name}}') + picture(1))}</w:hdr>`;
  const input = await template(picture(2, 'Unbound main image'), {
    'word/header1.xml': header,
    'word/_rels/header1.xml.rels': `<Relationships xmlns="${ns.packageRelationship}"><Relationship Id="shared" Type="${ns.relationship}/image" Target="media/original.png"/></Relationships>`,
    'word/_rels/document.xml.rels': `<Relationships xmlns="${ns.packageRelationship}"><Relationship Id="shared" Type="${ns.relationship}/image" Target="media/original.png"/><Relationship Id="header" Type="${ns.relationship}/header" Target="header1.xml"/></Relationships>`,
  });
  const output = await render(input, { items: [{ name: 'A', image: 'red' }, { name: 'B', image: 'blue' }] }, { attachments: { red: png([255, 0, 0]), blue: png([0, 0, 255]) } });
  assert.deepEqual(await texts(output, 'word/header1.xml'), ['A', '', 'B', '']);
  const ids = [await document(output), await document(output, 'word/header1.xml')].flatMap(d => descendants(d, ns.wordDrawing, 'docPr').map(n => n.getAttribute('id')));
  assert.equal(new Set(ids).size, 3);
});

test('HTTP worker renders loops and returns typed data errors', async () => {
  const key = 'imprint-loop-integration-test-key', app = createApp({ apiKey: key });
  const input = await template(block(p('{{name}}')));
  try {
    const send = (data: object) => app.inject({ method: 'POST', url: '/v1/render', headers: { authorization: `Bearer ${key}` }, payload: { template: input.toString('base64'), data } });
    const response = await send({ items: [{ name: 'API A' }, { name: 'API B' }] });
    assert.equal(response.statusCode, 200); assert.deepEqual(await texts(response.rawPayload), ['API A', 'API B']);
    assert.equal((await send({ items: null })).statusCode, 422);
  } finally { await app.close(); }
});

test('an empty block loop retains the required paragraph in its table cell', async () => {
  const output = await render(await template(row(block(p('{{name}}')))), { items: [] });
  const xml = await document(output);
  const cells = descendants(xml, ns.word, 'tc');
  assert.equal(cells.length, 1);
  assert.equal(descendants(cells[0]!, ns.word, 'p').length, 1);
  assert.deepEqual(await texts(output), ['']);
});

test('paragraph identities are regenerated for clones while surrounding paragraphs stay unchanged', async () => {
  const content = '<w:p xmlns:w14="http://schemas.microsoft.com/office/word/2010/wordml" w14:paraId="00000009" w14:textId="00000010"><w:r><w:t>{{name}}</w:t></w:r></w:p>';
  const output = await render(await template(block(content)), { items: [{ name: 'A' }, { name: 'B' }] });
  const paragraphs = descendants(await document(output), ns.word, 'p');
  const ids = paragraphs.flatMap(p => [p.getAttributeNS('http://schemas.microsoft.com/office/word/2010/wordml', 'paraId'), p.getAttributeNS('http://schemas.microsoft.com/office/word/2010/wordml', 'textId')]);
  assert.equal(new Set(ids).size, 4);
  assert(ids.every(id => /^[0-9A-F]{8}$/.test(id!)));
});
