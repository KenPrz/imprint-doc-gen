import test from 'node:test';
import assert from 'node:assert/strict';
import { createServer } from 'node:http';
import { RenderError } from '../src/errors.ts';
import { ns } from '../src/xml.ts';
import { writePackage } from '../src/package.ts';
import { createApp } from '../src/app.ts';
import { parseSheetRule, sheetData, templateSheetRule } from '../src/sheets.ts';

const code = (value: string) => (error: unknown) => error instanceof RenderError && error.code === value;
const items = (n: number) => Array.from({ length: n }, (_, i) => ({ name: `Item ${i + 1}` }));

test('sheet rule property parses array, rows and padding', () => {
  assert.deepEqual(parseSheetRule('items:10'), { array: 'items', rowsPerSheet: 10, pad: true });
  assert.deepEqual(parseSheetRule(' po.lines : 5 : nopad '), { array: 'po.lines', rowsPerSheet: 5, pad: false });
  for (const bad of ['items', 'items:0', 'items:501', 'items:x', 'items:3:maybe', '1x:3', 'a..b:3', 'items:3:pad:x', 'constructor:3'])
    assert.throws(() => parseSheetRule(bad), code('INVALID_SHEETS'), bad);
});

test('sheet data splits, pads and scopes each sheet', () => {
  const data = { po: 'M-1', items: items(23) };
  const sheets = sheetData(data, { array: 'items', rowsPerSheet: 10, pad: true }, [{ serial: 'A' }, { serial: 'B' }, { serial: 'C' }], 100);
  assert.equal(sheets.length, 3);
  assert.deepEqual(sheets.map((s) => (s.items as unknown[]).length), [10, 10, 10]);
  assert.deepEqual((sheets[2]!.items as unknown[]).slice(3), [{}, {}, {}, {}, {}, {}, {}]);
  assert.deepEqual((sheets[1]!.items as { name: string }[])[0], { name: 'Item 11' });
  assert.deepEqual(sheets.map((s) => s.sheet), [
    { serial: 'A', number: 1, count: 3, isFirst: '1', isLast: '', notLast: '1' },
    { serial: 'B', number: 2, count: 3, isFirst: '', isLast: '', notLast: '1' },
    { serial: 'C', number: 3, count: 3, isFirst: '', isLast: '1', notLast: '' },
  ]);
  assert.equal(sheets[0]!.po, 'M-1');
  assert.equal(data.items.length, 23, 'input is not mutated');
});

test('sheet data handles nested paths, no padding, empty and missing arrays', () => {
  const nested = sheetData({ po: { lines: items(3), id: 7 } }, { array: 'po.lines', rowsPerSheet: 2, pad: false }, undefined, 100);
  assert.deepEqual(nested.map((s) => ((s.po as { lines: unknown[] }).lines).length), [2, 1]);
  assert.equal((nested[1]!.po as { id: number }).id, 7);
  for (const data of [{ items: [] }, {}]) {
    const [only, ...rest] = sheetData(data, { array: 'items', rowsPerSheet: 4, pad: true }, undefined, 100);
    assert.equal(rest.length, 0);
    assert.deepEqual(only!.items, [{}, {}, {}, {}]);
    assert.deepEqual(only!.sheet, { number: 1, count: 1, isFirst: '1', isLast: '1', notLast: '' });
  }
});

test('sheet data rejects wrong shapes and too many sheets', () => {
  const rule = { array: 'items', rowsPerSheet: 10, pad: true };
  assert.throws(() => sheetData({ items: 'x' }, rule, undefined, 100), code('INVALID_SHEETS'));
  assert.throws(() => sheetData({ items: [1] }, rule, undefined, 100), code('INVALID_SHEETS'));
  assert.throws(() => sheetData({ items: items(11) }, rule, [{}], 100), code('SHEET_DATA_MISMATCH'));
  assert.throws(() => sheetData({ items: items(31) }, rule, undefined, 3), code('LIMIT_EXCEEDED'));
});

const docx = (body: string, custom?: string) =>
  writePackage(
    new Map(
      Object.entries({
        '[Content_Types].xml': `<Types xmlns="${ns.contentTypes}"><Default Extension="xml" ContentType="application/xml"/><Default Extension="rels" ContentType="application/vnd.openxmlformats-package.relationships+xml"/><Override PartName="/word/document.xml" ContentType="application/vnd.openxmlformats-officedocument.wordprocessingml.document.main+xml"/></Types>`,
        '_rels/.rels': `<Relationships xmlns="${ns.packageRelationship}"><Relationship Id="main" Type="${ns.relationship}/officeDocument" Target="word/document.xml"/></Relationships>`,
        'word/document.xml': `<?xml version="1.0"?><w:document xmlns:w="${ns.word}"><w:body>${body}<w:sectPr/></w:body></w:document>`,
        ...(custom === undefined
          ? {}
          : {
              'docProps/custom.xml': `<?xml version="1.0"?><Properties xmlns="http://schemas.openxmlformats.org/officeDocument/2006/custom-properties" xmlns:vt="http://schemas.openxmlformats.org/officeDocument/2006/docPropsVTypes"><property fmtid="{D5CDD505-2E9C-101B-9397-08002B2CF9AE}" pid="2" name="Other"><vt:lpwstr>x</vt:lpwstr></property><property fmtid="{D5CDD505-2E9C-101B-9397-08002B2CF9AE}" pid="3" name="imprint.sheets"><vt:lpwstr>${custom}</vt:lpwstr></property></Properties>`,
            }),
      }).map(([name, value]) => [name, Buffer.from(value)]),
    ),
  );
const p = (text: string) => `<w:p><w:r><w:t>${text}</w:t></w:r></w:p>`;
const row = (text: string) => `<w:tr><w:tc>${p(text)}</w:tc></w:tr>`;
const form = `${p('Sheet {d.sheet.number} of {d.sheet.count} {d.sheet.serial}')}<w:tbl>${row('{d.items[i].name}')}${row('{d.items[i+1].name}')}</w:tbl>${p('{d.sheet.isLast:ifNEM():showBegin}Total {d.total}')}${p('{d.sheet.isLast:showEnd}')}`;

test('template sheet rule reads the DOCX and ODF custom properties', async () => {
  const { readPackage } = await import('../src/package.ts');
  assert.deepEqual(templateSheetRule(await readPackage(await docx(p('x'), 'items:10'))), { array: 'items', rowsPerSheet: 10, pad: true });
  assert.equal(templateSheetRule(await readPackage(await docx(p('x')))), null);
  const odt = new Map([
    ['mimetype', Buffer.from('application/vnd.oasis.opendocument.text')],
    ['meta.xml', Buffer.from('<office:document-meta xmlns:office="urn:oasis:names:tc:opendocument:xmlns:office:1.0" xmlns:meta="urn:oasis:names:tc:opendocument:xmlns:meta:1.0"><office:meta><meta:user-defined meta:name="imprint.sheets">rows:4:nopad</meta:user-defined></office:meta></office:document-meta>')],
  ]);
  assert.deepEqual(templateSheetRule(odt), { array: 'rows', rowsPerSheet: 4, pad: false });
});

async function withConverter(run: (endpoint: string, calls: { merge?: string; files: { name: string; text: string }[] }[]) => Promise<void>) {
  const calls: { merge?: string; files: { name: string; text: string }[] }[] = [];
  const server = createServer(async (request, response) => {
    const body = await new Request('http://x', { method: 'POST', headers: request.headers as HeadersInit, body: request as unknown as BodyInit, duplex: 'half' } as RequestInit).formData();
    const { readPackage } = await import('../src/package.ts');
    const files = await Promise.all(
      (body.getAll('files') as File[]).map(async (file) => ({
        name: file.name,
        text: (await readPackage(new Uint8Array(await file.arrayBuffer()))).get('word/document.xml')!.toString().replace(/<[^>]+>/g, ' ').replace(/\s+/g, ' ').trim(),
      })),
    );
    calls.push({ merge: body.get('merge')?.toString(), files });
    response.end('%PDF-1.7\nmerged');
  });
  await new Promise<void>((resolve) => server.listen(0, '127.0.0.1', resolve));
  try {
    await run(`http://127.0.0.1:${(server.address() as { port: number }).port}`, calls);
  } finally {
    server.closeAllConnections();
    await new Promise<void>((resolve) => server.close(() => resolve()));
  }
}

test('HTTP render paginates a declared template into one merged PDF', async () => {
  const key = 'k'.repeat(32);
  await withConverter(async (converterUrl, calls) => {
    const app = createApp({ apiKey: key, converterUrl });
    const post = (url: string, payload: object) => app.inject({ method: 'POST', url, headers: { authorization: `Bearer ${key}` }, payload });
    try {
      const template = (await docx(form, 'items:2')).toString('base64');
      const ok = await post('/v1/render', { template, output: 'pdf', data: { items: items(3), total: 9 }, sheets: { perSheet: [{ serial: 'S1' }, { serial: 'S2' }] } });
      assert.equal(ok.statusCode, 200, ok.body);
      assert.equal(ok.body, '%PDF-1.7\nmerged');
      assert.equal(calls.length, 1);
      assert.equal(calls[0]!.merge, 'true');
      assert.deepEqual(calls[0]!.files.map((f) => f.name), ['sheet-0001.docx', 'sheet-0002.docx']);
      assert.deepEqual(calls[0]!.files.map((f) => f.text), ['Sheet 1 of 2 S1 Item 1 Item 2', 'Sheet 2 of 2 S2 Item 3 Total 9']);

      const override = await post('/v1/preview', { template, data: { items: items(3) }, sheets: { rowsPerSheet: 3 } });
      assert.equal(override.statusCode, 200, override.body);
      assert.deepEqual(calls[1]!.files.map((f) => f.text), ['Sheet 1 of 1 Item 1 Item 2 Item 3 Total']);

      const plain = (await docx(form)).toString('base64');
      const requested = await post('/v1/render', { template: plain, output: 'pdf', data: { items: items(1) }, sheets: { array: 'items', rowsPerSheet: 1 } });
      assert.equal(requested.statusCode, 200, requested.body);
      assert.equal(calls[2]!.files.length, 1);

      const docxOut = await post('/v1/render', { template, data: { items: items(1) } });
      assert.equal(docxOut.statusCode, 422);
      assert.equal(docxOut.json().error, 'SHEETS_REQUIRE_PDF');
      const incomplete = await post('/v1/render', { template: plain, output: 'pdf', data: {}, sheets: { rowsPerSheet: 2 } });
      assert.equal(incomplete.json().error, 'INVALID_SHEETS');
      const mismatch = await post('/v1/render', { template, output: 'pdf', data: { items: items(3) }, sheets: { perSheet: [{}] } });
      assert.equal(mismatch.json().error, 'SHEET_DATA_MISMATCH');
      const unpaginated = await post('/v1/render', { template: plain, output: 'pdf', data: { items: items(1) } });
      assert.equal(unpaginated.statusCode, 200);
      assert.equal(calls.at(-1)!.merge, undefined);

      const inspected = await post('/v1/inspect', { template });
      assert.deepEqual(inspected.json().sheets, { array: 'items', rowsPerSheet: 2, pad: true });
      assert.equal((await post('/v1/inspect', { template: plain })).json().sheets, null);
    } finally {
      await app.close();
    }
  });
});
