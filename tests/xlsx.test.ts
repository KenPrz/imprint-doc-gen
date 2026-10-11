import test from 'node:test';
import assert from 'node:assert/strict';
import { createServer } from 'node:http';
import sharp from 'sharp';
import { compile, render, RenderError } from '../src/index.ts';
import { createApp } from '../src/app.ts';
import { readPackage, writePackage } from '../src/package.ts';
import { ns, parseXml, descendants } from '../src/xml.ts';
import { png } from '../scripts/corpus/common.ts';

const code = (code: string) => (error: unknown) => error instanceof RenderError && error.code === code;
const rels = (links: string) => `<Relationships xmlns="${ns.packageRelationship}">${links}</Relationships>`;
const rel = (id: string, type: string, target: string) => `<Relationship Id="${id}" Type="${ns.relationship}/${type}" Target="${target}"/>`;
const inline = (ref: string, text: string, attrs = '') => `<s:c r="${ref}" t="inlineStr" s="2" ${attrs}><s:is><s:t>${text}</s:t></s:is></s:c>`;
const rich = (ref: string) => `<s:c r="${ref}" t="inlineStr"><s:is><s:r><s:rPr><s:b/></s:rPr><s:t>Dear {{cus</s:t></s:r><s:r><s:rPr><s:i/></s:rPr><s:t>tomer.name}}!</s:t></s:r></s:is></s:c>`;
async function template(cells: string, shared?: string[], extra: Record<string, string | Buffer> = {}) {
  return writePackage(new Map(Object.entries({
    '[Content_Types].xml': `<Types xmlns="${ns.contentTypes}"><Default Extension="xml" ContentType="application/xml"/><Default Extension="rels" ContentType="application/vnd.openxmlformats-package.relationships+xml"/><Default Extension="png" ContentType="image/png"/><Override PartName="/xl/workbook.xml" ContentType="application/vnd.openxmlformats-officedocument.spreadsheetml.sheet.main+xml"/></Types>`,
    '_rels/.rels': rels(rel('book', 'officeDocument', 'xl/workbook.xml')),
    'xl/workbook.xml': `<s:workbook xmlns:s="${ns.spreadsheet}" xmlns:r="${ns.relationship}"><s:sheets><s:sheet name="Main" sheetId="1" r:id="sheet"/></s:sheets><s:definedNames><s:definedName name="keep">Main!$A$1</s:definedName></s:definedNames><s:calcPr calcMode="manual"/></s:workbook>`,
    'xl/_rels/workbook.xml.rels': rels(rel('sheet', 'worksheet', 'worksheets/main.xml') + (shared ? rel('strings', 'sharedStrings', 'strings/custom.xml') : '')),
    'xl/worksheets/main.xml': `<s:worksheet xmlns:s="${ns.spreadsheet}" xmlns:r="${ns.relationship}"><s:dimension ref="A1:H1"/><s:sheetData><s:row r="1" ht="30" customHeight="1">${cells}</s:row></s:sheetData><s:mergeCells count="1"><s:mergeCell ref="G1:H1"/></s:mergeCells><s:drawing r:id="drawing"/></s:worksheet>`,
    'xl/worksheets/_rels/main.xml.rels': rels(rel('drawing', 'drawing', '../drawings/main.xml')),
    'xl/drawings/main.xml': `<xdr:wsDr xmlns:xdr="${ns.spreadsheetDrawing}"/>`,
    'customXml/keep.xml': '<keep>untouched bytes</keep>',
    ...(shared ? { 'xl/strings/custom.xml': `<s:sst xmlns:s="${ns.spreadsheet}" count="99" uniqueCount="${shared.length}">${shared.map(s => `<s:si>${s}</s:si>`).join('')}</s:sst>` } : {}), ...extra,
  }).map(([name, value]) => [name, Buffer.from(value)])));
}
async function cells(bytes: Uint8Array) {
  const parts = await readPackage(bytes);
  const strings = parts.has('xl/strings/custom.xml') ? descendants(parseXml(parts.get('xl/strings/custom.xml')!.toString(), 'strings'), ns.spreadsheet, 'si') : [];
  return descendants(parseXml(parts.get('xl/worksheets/main.xml')!.toString(), 'sheet'), ns.spreadsheet, 'c').map(c => {
    const type = c.getAttribute('t'), raw = descendants(c, ns.spreadsheet, 'v')[0]?.textContent ?? '';
    const texts = type === 's' ? descendants(strings[Number(raw)]!, ns.spreadsheet, 't') : descendants(c, ns.spreadsheet, 't');
    const text = texts.map(t => t.textContent).join('').replace(/_x([\da-fA-F]{4})_/g, (_, code: string) => String.fromCharCode(parseInt(code, 16)));
    return { ref: c.getAttribute('r'), type, raw, text, style: c.getAttribute('s'), formula: descendants(c, ns.spreadsheet, 'f')[0]?.textContent };
  });
}
const picture = (id: number, description: string, anchor = 'oneCellAnchor') => {
  const from = '<xdr:from><xdr:col>4</xdr:col><xdr:colOff>7</xdr:colOff><xdr:row>2</xdr:row><xdr:rowOff>9</xdr:rowOff></xdr:from>';
  const extent = '<xdr:ext cx="914400" cy="609600"/>';
  const geometry = anchor === 'absoluteAnchor' ? '<xdr:pos x="100" y="200"/>' + extent : anchor === 'twoCellAnchor' ? from + '<xdr:to><xdr:col>6</xdr:col><xdr:colOff>0</xdr:colOff><xdr:row>5</xdr:row><xdr:rowOff>0</xdr:rowOff></xdr:to>' : from + extent;
  return `<xdr:${anchor}>${geometry}<xdr:pic><xdr:nvPicPr><xdr:cNvPr id="${id}" name="Picture ${id}" descr="${description}"/><xdr:cNvPicPr/></xdr:nvPicPr><xdr:blipFill><a:blip r:embed="shared"/><a:srcRect l="100"/><a:stretch><a:fillRect/></a:stretch></xdr:blipFill><xdr:spPr><a:xfrm rot="60000"><a:ext cx="914400" cy="609600"/></a:xfrm></xdr:spPr></xdr:pic><xdr:clientData/></xdr:${anchor}>`;
};
const pictureParts = (pictures: string) => ({
  'xl/drawings/main.xml': `<xdr:wsDr xmlns:xdr="${ns.spreadsheetDrawing}" xmlns:a="${ns.drawing}" xmlns:r="${ns.relationship}">${pictures}</xdr:wsDr>`,
  'xl/drawings/_rels/main.xml.rels': rels(rel('shared', 'image', '../media/original.png')),
  'xl/media/original.png': png([100, 100, 100]),
});

test('XLSX whole-cell bindings retain numbers, booleans and null as typed values', async () => {
  const input = await template(inline('A1', '{{number}}') + inline('B1', '{{flag}}') + inline('C1', '{{empty}}') + inline('D1', '{{text}}') + inline('E1', 'value={{number}}'));
  const compiled = await compile(input); assert.equal(compiled.format, 'xlsx');
  const result = await cells(await compiled.render({ number: 0, flag: false, empty: null, text: '=HYPERLINK("https://example.test")' }));
  assert.deepEqual(result.slice(0, 3).map(c => [c.type, c.raw, c.style]), [['n', '0', '2'], ['b', '0', '2'], ['n', '', '2']]);
  assert.equal(result[3]!.type, 'inlineStr'); assert.equal(result[3]!.text, '=HYPERLINK("https://example.test")');
  assert.equal(result[3]!.formula, undefined); assert.equal(result[4]!.text, 'value=0');
});
test('XLSX split rich-text tags preserve styles, Unicode and surrounding text', async () => {
  const input = await template(rich('A1'));
  const output = await render(input, { customer: { name: '  Zoë & 山田 <💡>  ' } });
  assert.equal((await cells(output))[0]!.text, 'Dear   Zoë & 山田 <💡>  !');
  const parts = await readPackage(output), xml = parts.get('xl/worksheets/main.xml')!.toString();
  assert(xml.includes('<s:rPr><s:b/></s:rPr>')); assert(xml.includes('<s:rPr><s:i/></s:rPr>'));
  assert(xml.includes('<s:t xml:space="preserve">!</s:t>'));
  assert.deepEqual(parts.get('customXml/keep.xml'), (await readPackage(input)).get('customXml/keep.xml'));
});
test('XLSX shared strings stay shared while whole numeric references become native cells', async () => {
  const input = await template('<s:c r="A1" t="s"><s:v>0</s:v></s:c><s:c r="B1" t="s"><s:v>0</s:v></s:c><s:c r="C1" t="s"><s:v>1</s:v></s:c><s:c r="D1" t="s"><s:v>2</s:v></s:c>', ['<s:r><s:rPr><s:b/></s:rPr><s:t>{{na</s:t></s:r><s:r><s:t>me}}</s:t></s:r>', '<s:t>{{count}}</s:t>', '<s:t>untouched</s:t>']);
  const output = await render(input, { name: 'Shared', count: -12.5 });
  const result = await cells(output);
  assert.deepEqual(result.map(c => [c.type, c.type === 's' ? c.text : c.raw]), [['s', 'Shared'], ['s', 'Shared'], ['n', '-12.5'], ['s', 'untouched']]);
  const xml = (await readPackage(output)).get('xl/strings/custom.xml')!.toString();
  assert(xml.includes('count="3"')); assert(xml.includes('uniqueCount="3"')); assert(!xml.includes('{{')); assert(xml.includes('<s:rPr><s:b/></s:rPr>'));
});
test('XLSX text supports line breaks, tabs, literal escape sequences and template-looking input', async () => {
  const value = 'line\r\nnext\t_x0041_ _x005F_ {{untouched}}';
  const input = await template(inline('A1', '{{text}}') + inline('B1', '_x0041_ {{text}}'));
  const output = await render(input, { text: value });
  assert.deepEqual((await cells(output)).map(c => c.text), [value, 'A ' + value]);
  const xml = (await readPackage(output)).get('xl/worksheets/main.xml')!.toString();
  assert(xml.includes('_x005F_x0041_')); assert(xml.includes('_x000D_'));
});
test('XLSX preserves formula XML, style and geometry and requests recalculation', async () => {
  const formula = '<s:c r="B1" s="2"><s:f>A1*2+\'Lookup Data\'!$B$1</s:f><s:v>999</s:v></s:c>';
  const input = await template(inline('A1', '{{count}}') + formula);
  const parts = await readPackage(await render(input, { count: 15 }));
  const xml = parts.get('xl/worksheets/main.xml')!.toString();
  assert(xml.includes(formula)); assert(xml.includes('<s:mergeCell ref="G1:H1"/>')); assert(xml.includes('ht="30" customHeight="1"'));
  const book = parts.get('xl/workbook.xml')!.toString();
  assert(book.includes('calcMode="auto"')); assert(book.includes('fullCalcOnLoad="1"')); assert(book.includes('forceFullCalc="1"')); assert(book.includes('Main!$A$1'));
});
test('XLSX inserts missing calculation properties before workbook extensions', async () => {
  const parts = await readPackage(await template(inline('A1', '{{name}}')));
  parts.set('xl/workbook.xml', Buffer.from(parts.get('xl/workbook.xml')!.toString().replace('<s:calcPr calcMode="manual"/>', '<s:extLst/>')));
  const output = await readPackage(await render(await writePackage(parts), { name: 'ok' }));
  assert.match(output.get('xl/workbook.xml')!.toString(), /<s:calcPr[^>]+\/><s:extLst\/>/);
});
test('XLSX handles hidden sheets and relationship targets outside conventional sheet paths', async () => {
  const parts = await readPackage(await template(inline('A1', '{{name}}')));
  parts.set('xl/workbook.xml', Buffer.from(parts.get('xl/workbook.xml')!.toString().replace('</s:sheets>', '<s:sheet name="Hidden" state="hidden" sheetId="9" r:id="hidden"/></s:sheets>')));
  parts.set('xl/_rels/workbook.xml.rels', Buffer.from(parts.get('xl/_rels/workbook.xml.rels')!.toString().replace('</Relationships>', rel('hidden', 'worksheet', '../data/hidden.xml') + '</Relationships>')));
  parts.set('data/hidden.xml', Buffer.from(`<worksheet xmlns="${ns.spreadsheet}"><sheetData><row r="1"><c r="A1" t="inlineStr"><is><t>{{name}}</t></is></c></row></sheetData></worksheet>`));
  const output = await readPackage(await render(await writePackage(parts), { name: 'hidden value' }));
  assert(output.get('data/hidden.xml')!.toString().includes('hidden value')); assert(output.get('xl/workbook.xml')!.toString().includes('state="hidden"'));
});
test('XLSX replaces only bound pictures and preserves anchors, crop and rotation', async () => {
  const input = await template('', undefined, pictureParts(picture(1, '{{image:customer.logo}}') + picture(2, 'Unbound')));
  const red = png([255, 0, 0]);
  const parts = await readPackage(await render(input, { customer: { logo: 'red' } }, { attachments: { red } }));
  const xml = parts.get('xl/drawings/main.xml')!.toString(), doc = parseXml(xml, 'drawing');
  const blips = descendants(doc, ns.drawing, 'blip');
  assert.notEqual(blips[0]!.getAttributeNS(ns.relationship, 'embed'), 'shared'); assert.equal(blips[1]!.getAttributeNS(ns.relationship, 'embed'), 'shared');
  assert(xml.includes('<xdr:colOff>7</xdr:colOff>')); assert(xml.includes('<a:srcRect l="100"/>')); assert(xml.includes('rot="60000"'));
  assert.deepEqual(parts.get('xl/media/original.png'), png([100, 100, 100]));
  assert.equal([...parts].filter(([name, bytes]) => name.includes('imprint-') && bytes.equals(red)).length, 1);
});
test('XLSX supports all three anchor kinds and deduplicates replacement media', async () => {
  const input = await template('', undefined, pictureParts(['oneCellAnchor', 'twoCellAnchor', 'absoluteAnchor'].map((anchor, i) => picture(i + 1, '{{image:red}}', anchor)).join('')));
  const output = await readPackage(await render(input, {}, { attachments: { red: png([255, 0, 0]) } }));
  assert.equal([...output.keys()].filter(n => n.includes('imprint-')).length, 1);
  const ids = descendants(parseXml(output.get('xl/drawings/main.xml')!.toString(), 'drawing'), ns.drawing, 'blip').map(e => e.getAttributeNS(ns.relationship, 'embed'));
  assert.equal(new Set(ids).size, 3);
});
test('XLSX JPEG attachments are decoded and get matching content types', async () => {
  const jpeg = await sharp(png([20, 60, 100])).jpeg().toBuffer();
  const input = await template('', undefined, pictureParts(picture(1, '{{image:logo}}')));
  const output = await readPackage(await render(input, {}, { attachments: { logo: jpeg } }));
  assert([...output].some(([name, bytes]) => name.endsWith('.jpg') && bytes.equals(jpeg)));
  assert(output.get('[Content_Types].xml')!.toString().includes('ContentType="image/jpeg"'));
});
test('XLSX failed attachments do not poison compiled reuse or concurrent renders', async () => {
  const compiled = await compile(await template(inline('A1', '{{name}}'), undefined, pictureParts(picture(1, '{{image:logo}}'))));
  await assert.rejects(compiled.render({ name: 'missing', logo: 'absent' }), code('MISSING_ATTACHMENT'));
  await assert.rejects(compiled.render({ name: 'invalid' }, { attachments: { logo: Buffer.from('not image') } }), code('INVALID_IMAGE'));
  const outputs = await Promise.all(['A', 'B', 'C'].map(name => compiled.render({ name }, { attachments: { logo: png([0, 0, 255]) } })));
  assert.deepEqual(await Promise.all(outputs.map(async output => (await cells(output))[0]!.text)), ['A', 'B', 'C']);
  assert.deepEqual(outputs[0], await compiled.render({ name: 'A' }, { attachments: { logo: png([0, 0, 255]) } }));
});
test('XLSX rejects malformed row loops, formula templating and malformed or unsafe expressions', async () => {
  for (const [text, expected] of [['{{#each items}}', 'TEMPLATE_STRUCTURE'], ['{{/each}}', 'TEMPLATE_STRUCTURE'], ['{{name', 'TEMPLATE_SYNTAX'], ['{{__proto__.x}}', 'TEMPLATE_EXPRESSION'], ['{{name + 1}}', 'TEMPLATE_EXPRESSION'], ['{{../name}}', 'TEMPLATE_EXPRESSION']]) await assert.rejects(compile(await template(inline('A1', text!))), code(expected!));
  await assert.rejects(compile(await template('<s:c r="A1"><s:f>{{formula}}</s:f><s:v>0</s:v></s:c>')), code('UNSUPPORTED_TEMPLATE'));
});
test('XLSX rejects missing data, complex values and invalid XML characters without evaluating accessors', async () => {
  const input = await template(inline('A1', '{{name}}'));
  await assert.rejects(render(input, {}), code('MISSING_DATA'));
  for (const name of [{}, [], Infinity, 'bad\u0001']) await assert.rejects(render(input, { name }), code('INVALID_DATA'));
  let accessed = false;
  await assert.rejects(render(input, Object.defineProperty({}, 'name', { get() { accessed = true; return 'bad'; } })), code('INVALID_DATA'));
  assert.equal(accessed, false);
});
test('XLSX bounds cell strings and total shared-string output before allocation grows unchecked', async () => {
  const compiled = await compile(await template(inline('A1', '{{name}}')));
  assert.equal((await cells(await compiled.render({ name: 'x'.repeat(32767) })))[0]!.text.length, 32767);
  await assert.rejects(compiled.render({ name: 'x'.repeat(32768) }), code('LIMIT_EXCEEDED'));
  const input = await template('<s:c r="A1" t="s"><s:v>0</s:v></s:c>', Array.from({ length: 300 }, () => '<s:t>{{name}}</s:t>'));
  await assert.rejects(render(input, { name: 'x'.repeat(32767) }), code('LIMIT_EXCEEDED'));
});
test('XLSX rejects corrupt shared indices, image links and unsupported format variants', async () => {
  for (const value of ['-1', '1', 'NaN', '0.5']) await assert.rejects(compile(await template(`<s:c r="A1" t="s"><s:v>${value}</s:v></s:c>`, ['<s:t>ok</s:t>'])), code('INVALID_PACKAGE'));
  await assert.rejects(compile(await template('', undefined, { ...pictureParts(picture(1, '{{image:logo}}')), 'xl/drawings/_rels/main.xml.rels': rels('') })), code('INVALID_PACKAGE'));
  const parts = await readPackage(await template(inline('A1', '{{name}}')));
  parts.set('xl/vbaProject.bin', Buffer.from('macro')); await assert.rejects(compile(await writePackage(parts)), code('UNSUPPORTED_FORMAT')); parts.delete('xl/vbaProject.bin');
  parts.set('xl/workbook.xml', Buffer.from(parts.get('xl/workbook.xml')!.toString().replace(ns.spreadsheet, 'http://purl.oclc.org/ooxml/spreadsheetml/main')));
  await assert.rejects(compile(await writePackage(parts)), code('UNSUPPORTED_FORMAT'));
});
test('XLSX API detects format, returns XLSX metadata and rejects cross-format output', async () => {
  const key = 'independent-xlsx-api-integration-key', app = createApp({ apiKey: key });
  const input = await template(inline('A1', '{{value}}'));
  const request = { method: 'POST' as const, url: '/v1/render', headers: { authorization: `Bearer ${key}` } };
  try {
    for (const output of [undefined, 'xlsx']) {
      const response = await app.inject({ ...request, payload: { template: input.toString('base64'), data: { value: 123 }, output } });
      assert.equal(response.statusCode, 200, response.body); assert.match(response.headers['content-type']!, /spreadsheetml.sheet/);
      assert.match(response.headers['content-disposition']!, /document.xlsx/); assert.equal((await cells(response.rawPayload))[0]!.raw, '123');
    }
    const response = await app.inject({ ...request, payload: { template: input.toString('base64'), data: { value: 123 }, output: 'docx' } });
    assert.equal(response.statusCode, 422); assert.equal(response.json().error, 'UNSUPPORTED_FORMAT');
  } finally { await app.close(); }
});
test('XLSX PDF requests send an XLSX filename to the converter', async () => {
  let upload = '';
  const server = createServer(async (request, response) => { const chunks = []; for await (const chunk of request) chunks.push(chunk); upload = Buffer.concat(chunks).toString('latin1'); response.end('%PDF-1.7\nxlsx'); });
  await new Promise<void>(resolve => server.listen(0, '127.0.0.1', resolve));
  const key = 'independent-xlsx-api-integration-key', app = createApp({ apiKey: key, converterUrl: `http://127.0.0.1:${(server.address() as { port: number }).port}` });
  try {
    const response = await app.inject({ method: 'POST', url: '/v1/render', headers: { authorization: `Bearer ${key}` }, payload: { template: (await template(inline('A1', '{{value}}'))).toString('base64'), data: { value: 7 }, output: 'pdf' } });
    assert.equal(response.statusCode, 200, response.body); assert.match(upload, /filename="document.xlsx"/); assert.equal(response.body, '%PDF-1.7\nxlsx');
  } finally { await app.close(); server.closeAllConnections(); await new Promise<void>(resolve => server.close(() => resolve())); }
});

test('XLSX preserves every part of an unbound workbook byte-for-byte', async () => {
  const input = await template(inline('A1', 'Literal _x0041_ text'), undefined, pictureParts(picture(1, 'Keep')));
  assert.deepEqual(await readPackage(await render(input, {})), await readPackage(input));
});
test('XLSX rejects ambiguous picture bindings and bound phonetic strings', async () => {
  const ambiguous = picture(1, '{{image:red}}').replace('name="Picture 1"', 'name="{{image:blue}}"');
  await assert.rejects(compile(await template('', undefined, pictureParts(ambiguous))), code('UNSUPPORTED_TEMPLATE'));
  const cell = '<s:c r="A1" t="inlineStr"><s:is><s:t>{{name}}</s:t><s:rPh sb="0" eb="1"><s:t>phonetic</s:t></s:rPh></s:is></s:c>';
  await assert.rejects(compile(await template(cell)), code('UNSUPPORTED_TEMPLATE'));
});

test('Carbone XLSX picture bindings accept attachments/data URLs, preserve optional placeholders and reject remote URLs', async () => {
  const input = await template('', undefined, pictureParts(picture(1, '{d.logo}')));
  const image = png([10, 20, 30]);
  for (const [data, options] of [[{}, { attachments: { logo: image } }], [{ logo: `data:image/png;base64,${image.toString('base64')}` }, {}]] as const) {
    const parts = await readPackage(await render(input, data, options));
    assert([...parts.values()].some(value => value.equals(image)));
    assert(!parts.get('xl/drawings/main.xml')!.toString().includes('{d.logo}'));
  }
  const missing = await readPackage(await render(input, {}));
  assert(missing.get('xl/drawings/main.xml')!.toString().includes('Image placeholder'));
  await assert.rejects(render(input, { logo: 'http://127.0.0.1/secret' }), code('UNSUPPORTED_IMAGE'));
  await assert.rejects(compile(await template('', undefined, pictureParts(picture(1, '{d.items[i].logo}')))), code('TEMPLATE_EXPRESSION'));
});
