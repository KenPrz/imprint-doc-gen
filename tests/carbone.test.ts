import test from 'node:test';
import assert from 'node:assert/strict';
import { readFile } from 'node:fs/promises';
import { createHash } from 'node:crypto';
import { render, compile, RenderError } from '../src/index.ts';
import { readPackage, writePackage } from '../src/package.ts';
import { parseXml, descendants, ns } from '../src/xml.ts';
import { withArrayLength } from '../scripts/pms/variants.ts';
import { snapshot } from '../scripts/pms/snapshot.ts';
import { createApp } from '../src/app.ts';

const fixtures = new URL('./fixtures/pms/', import.meta.url);
const manifest = JSON.parse(await readFile(new URL('manifest.json', fixtures), 'utf8')) as { templates: { file: string; sha256: string }[] };
for (const fixture of manifest.templates) test(`Carbone PMS template unchanged: ${fixture.file}`, async () => {
  const bytes = await readFile(new URL(fixture.file, fixtures));
  assert.equal(createHash('sha256').update(bytes).digest('hex'), fixture.sha256);
  const name = fixture.file.replace('-template.docx', '');
  const data = JSON.parse(await readFile(new URL(`${name}.json`, fixtures), 'utf8'));
  const expected = JSON.parse(await readFile(new URL(`${name}.expected.json`, fixtures), 'utf8'));
  const result = await render(bytes, data);
  assert.deepEqual(await snapshot(result), expected);
  for (const [part, content] of await readPackage(result)) if (/^word\/(document|header\d*|footer\d*)\.xml$/.test(part)) assert(!content.toString().includes('{d.'), part);
});
const errorCode = (code: string) => (error: unknown) => error instanceof RenderError && error.code === code;
const p = (value: string) => `<w:p><w:r><w:t>${value}</w:t></w:r></w:p>`;
const row = (value: string) => `<w:tr><w:tc>${p(value)}</w:tc></w:tr>`;
async function template(body: string) {
  const parts = await readPackage(await readFile(new URL('confirmation_letter-template.docx', fixtures)));
  const source = parts.get('word/document.xml')!.toString();
  parts.set('word/document.xml', Buffer.from(source.replace(/<w:body>[\s\S]*?<w:sectPr/, `<w:body>${body}<w:sectPr`)));
  return writePackage(parts);
}
async function texts(bytes: Uint8Array) {
  const main = (await readPackage(bytes)).get('word/document.xml')!.toString();
  return descendants(parseXml(main, 'test'), ns.word, 't').map(n => n.textContent).join('|');
}

test('Carbone joins split tags, escapes values, keeps false/zero and renders multiline values without parsing injected tags', async () => {
  const input = await template('<w:p><w:r><w:rPr><w:b/></w:rPr><w:t>{d.na</w:t></w:r><w:r><w:t>me}</w:t></w:r></w:p>' + p('{d.zero} / {d.flag} / {d.missing} / {d.nil}'));
  const result = await render(input, { name: 'A & B\nNext\t{d.literal}', zero: 0, flag: false, nil: null });
  const xml = (await readPackage(result)).get('word/document.xml')!.toString();
  assert(xml.includes('A &amp; B')); assert(xml.includes('<w:br/>')); assert(xml.includes('<w:tab/>'));
  assert(xml.includes('{d.literal}')); assert(xml.includes('<w:b/>'));
  assert((await texts(result)).includes('0 / false /  / '));
});

test('Carbone row markers remove both exemplars on empty arrays and repeat the full pattern on multiple items', async () => {
  const input = await template('<w:tbl>' + row('{d.rows[i].value} / {d.title}') + row('second line {d.rows[i].value}') + row('{d.rows[i+1].value}') + '</w:tbl>' + p('End'));
  for (const count of [0, 1, 3]) {
    const result = await render(input, { title: 'ROOT', rows: Array.from({ length: count }, (_, n) => ({ value: `item-${n}` })) });
    const xml = parseXml((await readPackage(result)).get('word/document.xml')!.toString(), 'test');
    assert.equal(descendants(xml, ns.word, 'tr').length, count * 2);
    assert.equal(descendants(xml, ns.word, 't').filter(n => n.textContent?.includes('ROOT')).length, count);
    assert(!(await texts(result)).includes('[i'));
  }
});

test('Carbone nested blocks preserve root and parent lookups with different inner lengths', async () => {
  const input = await template(p('Group {d.groups[i].name}') + '<w:tbl>' + row('{d.groups[i].items[i].name}/{d.groups[i].name}/{d.title}') + row('{d.groups[i].items[i+1].name}') + '</w:tbl>' + p('{d.groups[i+1].name}'));
  const result = await render(input, { title: 'ROOT', groups: [{ name: 'one', items: [{ name: 'A' }, { name: 'B' }] }, { name: 'two', items: [] }, { name: 'three', items: [{ name: 'C' }] }] });
  assert.equal(await texts(result), 'Group one|A/one/ROOT|B/one/ROOT|Group two|Group three|C/three/ROOT');
});

test('Carbone supports independent repeat regions over the same array and rejects missing/crossing markers', async () => {
  const table = '<w:tbl>' + row('{d.rows[i].name}') + row('{d.rows[i+1].name}') + '</w:tbl>';
  const result = await render(await template(table + p('separator') + table), { rows: [{ name: 'one' }, { name: 'two' }] });
  assert.equal(await texts(result), 'one|two|separator|one|two');
  await assert.rejects(compile(await template(p('{d.rows[i+1].name}'))), errorCode('TEMPLATE_STRUCTURE'));
  await assert.rejects(compile(await template(p('{d.rows[i].name}'))), errorCode('TEMPLATE_EXPRESSION'));
  await assert.rejects(compile(await template(p('{d.outer[i].name}') + p('{d.other[i].name}') + p('{d.outer[i+1].name}') + p('{d.other[i+1].name}'))), errorCode('TEMPLATE_STRUCTURE'));
});

test('Carbone missing arrays render empty; sparse/invalid arrays and excessive expansion remain bounded', async () => {
  const input = await template('<w:tbl>' + row('{d.rows[i].name}') + row('{d.rows[i+1]}') + '</w:tbl>');
  assert.equal(await texts(await render(input, {})), '');
  await assert.rejects(render(input, { rows: ['bad'] }), errorCode('INVALID_DATA'));
  await assert.rejects(render(input, { rows: new Array(2) }), errorCode('INVALID_DATA'));
  await assert.rejects(render(input, { rows: Array.from({ length: 10001 }, () => ({ name: 'x' })) }), errorCode('LIMIT_EXCEEDED'));
});

for (const tag of ['{d.total:formatN(2)}', '{d.rows[i=0].name}', '{d.rows[0].name}', '{c.name}', '{t(hello)}', '{o.lang=en}', '{d.__proto__.x}', '{d.rows[i+2].name}']) test(`Unsupported Carbone syntax fails explicitly: ${tag}`, async () => {
  await assert.rejects(compile(await template(p(tag))), error => error instanceof RenderError && ['UNSUPPORTED_CARBONE_TEMPLATE', 'TEMPLATE_EXPRESSION'].includes(error.code));
});

test('Carbone tags cannot cross a Word line break or mix native loop syntax', async () => {
  await assert.rejects(compile(await template('<w:p><w:r><w:t>{d.na</w:t><w:br/><w:t>me}</w:t></w:r></w:p>')), errorCode('UNSUPPORTED_TEMPLATE'));
  await assert.rejects(compile(await template(p('{{#each groups}}') + '<w:tbl>' + row('{d.rows[i].name}') + row('{d.rows[i+1].name}') + '</w:tbl>' + p('{{/each}}'))), errorCode('UNSUPPORTED_TEMPLATE'));
});

test('Carbone image alt text accepts attachments and data URLs, preserves an absent optional signature, and forbids remote fetches', async () => {
  const input = await readFile(new URL('registration_form-template.docx', fixtures));
  const data = JSON.parse(await readFile(new URL('registration_form.json', fixtures), 'utf8'));
  const logo = await readFile(new URL('../examples/pms/logo.png', import.meta.url));
  const attached = await readPackage(await render(input, data, { attachments: { guest_signature: logo } }));
  assert([...attached.values()].some(bytes => bytes.equals(logo)));
  const embedded = await readPackage(await render(input, { ...data, guest_signature: `data:image/png;base64,${logo.toString('base64')}` }));
  assert([...embedded.values()].some(bytes => bytes.equals(logo)));
  const absent = (await readPackage(await render(input, data))).get('word/document.xml')!.toString();
  assert(absent.includes('Image placeholder')); assert(!absent.includes('{d.guest_signature}'));
  await assert.rejects(render(input, { ...data, guest_signature: 'http://127.0.0.1/secret' }), errorCode('UNSUPPORTED_IMAGE'));
});

test('Carbone image data URLs remain independent for each repeated row', async () => {
  const base = await readPackage(await readFile(new URL('registration_form-template.docx', fixtures)));
  const xml = base.get('word/document.xml')!.toString();
  const drawing = xml.match(/<w:drawing>[\s\S]*?<\/w:drawing>/)![0].replace('{d.guest_signature}', '{d.rows[i].image}');
  base.set('word/document.xml', Buffer.from(xml.replace(/<w:body>[\s\S]*?<w:sectPr/, `<w:body><w:tbl><w:tr><w:tc>${p('{d.rows[i].name}')}<w:p><w:r>${drawing}</w:r></w:p></w:tc></w:tr>${row('{d.rows[i+1].name}')}</w:tbl><w:sectPr`)));
  const red = await readFile(new URL('../corpus/attachments/red.png', import.meta.url)), blue = await readFile(new URL('../corpus/attachments/blue.png', import.meta.url));
  const result = await readPackage(await render(await writePackage(base), { rows: [{ name: 'red', image: `data:image/png;base64,${red.toString('base64')}` }, { name: 'blue', image: `data:image/png;base64,${blue.toString('base64')}` }] }));
  assert([...result.values()].some(bytes => bytes.equals(red))); assert([...result.values()].some(bytes => bytes.equals(blue)));
  const document = parseXml(result.get('word/document.xml')!.toString(), 'test');
  assert.equal(new Set(descendants(document, ns.drawing, 'blip').map(node => node.getAttributeNS(ns.relationship, 'embed'))).size, 2);
});

test('Carbone-compatible endpoint accepts the PMS request shape and retains authentication/schema limits', async () => {
  const key = 'carbone-compatibility-integration-test-key', app = createApp({ apiKey: key });
  const input = { template: (await readFile(new URL('folio_statement-template.docx', fixtures))).toString('base64'), data: JSON.parse(await readFile(new URL('folio_statement.json', fixtures), 'utf8')), convertTo: 'docx' };
  try {
    const url = '/render/template?download=true', headers = { authorization: `Bearer ${key}` };
    assert.equal((await app.inject({ method: 'POST', url, payload: input })).statusCode, 401);
    const response = await app.inject({ method: 'POST', url, payload: input, headers });
    assert.equal(response.statusCode, 200, response.body);
    assert.deepEqual(await snapshot(response.rawPayload), JSON.parse(await readFile(new URL('folio_statement.expected.json', fixtures), 'utf8')));
    assert.equal((await app.inject({ method: 'POST', url, payload: { ...input, convertTo: 'csv' }, headers })).statusCode, 400);
    assert.equal((await app.inject({ method: 'POST', url: '/render/template?download=false', payload: input, headers })).statusCode, 400);
  } finally { await app.close(); }
});


for (const name of ['folio_statement', 'task_sheet', 'trial_balance']) test(`Carbone independent 0/1/3-item repetition oracle: ${name}`, async () => {
  const input = await readFile(new URL(`${name}-template.docx`, fixtures));
  const data = JSON.parse(await readFile(new URL(`${name}.json`, fixtures), 'utf8'));
  const expected = JSON.parse(await readFile(new URL(`${name}.variants.expected.json`, fixtures), 'utf8'));
  for (const count of [0, 1, 3]) assert.deepEqual(await snapshot(await render(input, withArrayLength(data, count) as Record<string, unknown>)), expected[String(count)], `Array length ${count}`);
});

test('Carbone XLSX invoice content matches independent zero/one/three-item oracle outputs', async () => {
  const { xlsxContent } = await import('../scripts/pms/xlsx-content.ts');
  const template = await readFile(new URL('../examples/pms/invoice.xlsx', import.meta.url));
  const data = JSON.parse(await readFile(new URL('../examples/pms/booking.json', import.meta.url), 'utf8'));
  const expected = JSON.parse(await readFile(new URL('./fixtures/carbone-xlsx-invoice.expected.json', import.meta.url), 'utf8'));
  for (const count of [0, 1, 3]) assert.deepEqual(await xlsxContent(await render(template, withArrayLength(data, count) as Record<string, unknown>)), expected[String(count)]);
});

test('Carbone root picture bindings retain attachment lookup inside a repeated row', async () => {
  const base = await readPackage(await readFile(new URL('registration_form-template.docx', fixtures)));
  const source = base.get('word/document.xml')!.toString();
  const drawing = source.match(/<w:drawing>[\s\S]*?<\/w:drawing>/)![0].replace('{d.guest_signature}', '{d.logo}');
  base.set('word/document.xml', Buffer.from(source.replace(/<w:body>[\s\S]*?<w:sectPr/, `<w:body><w:tbl><w:tr><w:tc>${p('{d.rows[i].name}')}<w:p><w:r>${drawing}</w:r></w:p></w:tc></w:tr>${row('{d.rows[i+1].name}')}</w:tbl><w:sectPr`)));
  const logo = await readFile(new URL('../examples/pms/logo.png', import.meta.url));
  const result = await readPackage(await render(await writePackage(base), { rows: [{ name: 'one' }, { name: 'two' }] }, { attachments: { logo } }));
  assert([...result.values()].some(bytes => bytes.equals(logo)));
  assert(!result.get('word/document.xml')!.toString().includes('Image placeholder'));
});

test('Carbone scalar paths keep root meaning inside a legacy DOCX loop', async () => {
  const input = await template(p('{{#each rows}}') + p('{d.title}: {{name}}') + p('{{/each}}'));
  assert.equal(await texts(await render(input, { title: 'ROOT', rows: [{ title: 'wrong scope', name: 'item' }] })), 'ROOT: item');
});
