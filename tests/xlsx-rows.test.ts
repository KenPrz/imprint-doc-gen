import test from 'node:test';
import assert from 'node:assert/strict';
import { compile, render, RenderError } from '../src/index.ts';
import { readPackage, writePackage } from '../src/package.ts';
import { ns, escapeXml, parseXml, descendants } from '../src/xml.ts';
import { png } from '../scripts/corpus/common.ts';

const code = (value: string) => (error: unknown) => error instanceof RenderError && error.code === value;
const t = (ref: string, text: string) => `<s:c r="${ref}" s="0" t="inlineStr"><s:is><s:t>${escapeXml(text)}</s:t></s:is></s:c>`;
const n = (ref: string, value: number) => `<s:c r="${ref}"><s:v>${value}</s:v></s:c>`;
const f = (ref: string, formula: string, attributes = '') => `<s:c r="${ref}"><s:f ${attributes}>${escapeXml(formula)}</s:f><s:v>999</s:v></s:c>`;
const row = (number: number, content: string) => `<s:row r="${number}" ht="24" customHeight="1">${content}</s:row>`;
const loop = (number = 2, path = 'items') => row(number, t(`A${number}`, `{{#each ${path}}}{{name}}`) + t(`B${number}`, '{{quantity}}') + t(`C${number}`, '{{price}}') + f(`D${number}`, `B${number}*C${number}`) + t(`E${number}`, '{{/each}}'));
const link = (id: string, type: string, target: string) => `<Relationship Id="${id}" Type="${ns.relationship}/${type}" Target="${target}"/>`;
const rels = (links: string) => `<Relationships xmlns="${ns.packageRelationship}">${links}</Relationships>`;
async function book(sheets: { name?: string; rows: string; suffix?: string }[], extra: Record<string, string | Buffer> = {}, names = '') {
  const parts: Record<string, string | Buffer> = {
    '[Content_Types].xml': `<Types xmlns="${ns.contentTypes}"><Default Extension="xml" ContentType="application/xml"/><Default Extension="rels" ContentType="application/vnd.openxmlformats-package.relationships+xml"/><Default Extension="png" ContentType="image/png"/><Override PartName="/xl/workbook.xml" ContentType="application/vnd.openxmlformats-officedocument.spreadsheetml.sheet.main+xml"/></Types>`,
    '_rels/.rels': rels(link('main', 'officeDocument', 'xl/workbook.xml')),
    'xl/workbook.xml': `<s:workbook xmlns:s="${ns.spreadsheet}" xmlns:r="${ns.relationship}"><s:sheets>${sheets.map((s, i) => `<s:sheet name="${escapeXml(s.name ?? 'Main')}" sheetId="${i + 1}" r:id="s${i}"/>`).join('')}</s:sheets>${names ? `<s:definedNames>${names}</s:definedNames>` : ''}<s:calcPr/></s:workbook>`,
    'xl/_rels/workbook.xml.rels': rels(sheets.map((_, i) => link(`s${i}`, 'worksheet', `worksheets/s${i}.xml`)).join('')),
    'customXml/keep.xml': '<keep>unchanged</keep>', ...extra,
  };
  for (const [i, sheet] of sheets.entries()) parts[`xl/worksheets/s${i}.xml`] = `<s:worksheet xmlns:s="${ns.spreadsheet}" xmlns:r="${ns.relationship}"><s:dimension ref="A1:F10"/><s:sheetData>${sheet.rows}</s:sheetData>${sheet.suffix ?? ''}</s:worksheet>`;
  return writePackage(new Map(Object.entries(parts).map(([path, value]) => [path, Buffer.from(value)])));
}
async function sheet(bytes: Uint8Array, index = 0) {
  const parts = await readPackage(bytes), source = parts.get(`xl/worksheets/s${index}.xml`)!.toString(), doc = parseXml(source, 'sheet');
  const cells = new Map(descendants(doc, ns.spreadsheet, 'c').map(c => [c.getAttribute('r')!, { type: c.getAttribute('t') || 'n', text: descendants(c, ns.spreadsheet, 't').map(t => t.textContent).join(''), value: descendants(c, ns.spreadsheet, 'v')[0]?.textContent, formula: descendants(c, ns.spreadsheet, 'f')[0]?.textContent }]));
  return { parts, source, doc, cells, rows: descendants(doc, ns.spreadsheet, 'row').map(n => Number(n.getAttribute('r'))) };
}
const items = [{ name: 'First', quantity: 2, price: 12.5 }, { name: 'Second', quantity: 3, price: 20 }];

test('XLSX rows repeat typed cells, calculate totals and remove marker-only cells', async () => {
  const input = await book([{ rows: row(1, t('A1', 'Before')) + loop() + row(3, t('A3', 'After')) }]);
  const output = await sheet(await render(input, { items }));
  assert.deepEqual(output.rows, [1, 2, 3, 4]);
  assert.equal(output.cells.get('A2')!.text, 'First'); assert.equal(output.cells.get('A3')!.text, 'Second'); assert.equal(output.cells.get('A4')!.text, 'After');
  assert.equal(output.cells.get('D2')!.value, '25'); assert.equal(output.cells.get('D3')!.value, '60'); assert.equal(output.cells.get('D3')!.formula, 'B3*C3');
  assert(!output.cells.has('E2')); assert(!output.cells.has('E3')); assert(output.source.includes('ht="24" customHeight="1"')); assert(!output.source.includes('{{'));
  assert.deepEqual(output.parts.get('customXml/keep.xml'), (await readPackage(input)).get('customXml/keep.xml'));
});
test('XLSX repeated rows use explicit parent scope and never reparse input tags', async () => {
  const input = await book([{ rows: row(2, t('A2', '{{#each items}}{{../title}} / {{name}}') + t('B2', '{{/each}}')) }]);
  const output = await sheet(await render(input, { title: 'Parent', items: [{ name: '{{literal}}' }] }));
  assert.equal(output.cells.get('A2')!.text, 'Parent / {{literal}}');
  await assert.rejects(render(input, { title: 'Parent', name: 'Root fallback', items: [{}] }), code('MISSING_DATA'));
});
test('XLSX shared rich strings render independently for each repeated item', async () => {
  const input = await book([{ rows: row(2, '<s:c r="A2" t="s"><s:v>0</s:v></s:c>' + t('B2', '{{/each}}')) }], {
    'xl/_rels/workbook.xml.rels': rels(link('s0', 'worksheet', 'worksheets/s0.xml') + link('sst', 'sharedStrings', 'sharedStrings.xml')),
    'xl/sharedStrings.xml': `<q:sst xmlns:q="${ns.spreadsheet}" xmlns:imprintS="urn:collision" count="1" uniqueCount="1"><q:si><q:r><q:rPr><q:b/></q:rPr><q:t>{{#each items}}{{na</q:t></q:r><q:r><q:rPr><q:i/></q:rPr><q:t>me}}</q:t></q:r></q:si></q:sst>`,
  });
  const output = await sheet(await render(input, { items }));
  assert.deepEqual([output.cells.get('A2')!.text, output.cells.get('A3')!.text], ['First', 'Second']);
  assert(output.source.includes('<q:rPr><q:b/></q:rPr>')); assert(output.source.includes('<q:rPr><q:i/></q:rPr>'));
  assert(!output.parts.get('xl/sharedStrings.xml')!.toString().includes('{{')); assert(output.parts.get('xl/sharedStrings.xml')!.toString().includes('count="0"'));
});
test('XLSX empty loops remove rows and shrink ranges without stale formulas', async () => {
  const input = await book([{ rows: row(1, n('B1', 7)) + loop() + row(3, f('A3', 'B2') + f('B3', 'SUM(C1:C2)')), suffix: '<s:autoFilter ref="A1:D2"/><s:conditionalFormatting sqref="B2:B2"><s:cfRule type="expression" priority="1"><s:formula>B2&gt;0</s:formula></s:cfRule></s:conditionalFormatting><s:dataValidations count="1"><s:dataValidation sqref="C2:C2" type="whole"><s:formula1>1</s:formula1></s:dataValidation></s:dataValidations>' }]);
  const output = await sheet(await render(input, { items: [] }));
  assert.deepEqual(output.rows, [1, 2]); assert.equal(output.cells.get('A2')!.formula, '#REF!'); assert.equal(output.cells.get('A2')!.type, 'e'); assert.equal(output.cells.get('A2')!.value, '#REF!');
  assert.equal(output.cells.get('B2')!.formula, 'SUM(C1:C1)'); assert.equal(output.cells.get('B2')!.value, '0');
  assert(output.source.includes('ref="A1:D1"')); assert(!output.source.includes('<s:conditionalFormatting')); assert(output.source.includes('dataValidations count="0"'));
});
test('XLSX missing, null, scalar and sparse loop data fail explicitly', async () => {
  const input = await book([{ rows: loop() }]);
  await assert.rejects(render(input, {}), code('MISSING_DATA'));
  for (const items of [null, 1, {}, [1], [null], new Array(1)]) await assert.rejects(render(input, { items }), code('INVALID_DATA'));
});
test('XLSX multiple row loops accumulate offsets and update external range consumers', async () => {
  const input = await book([{ rows: row(1, t('A1', 'Header')) + loop(2, 'first') + row(3, n('B3', 10)) + loop(4, 'second') + row(5, f('F5', 'SUM(B2:B4)')) }]);
  const output = await sheet(await render(input, { first: items, second: [items[0], items[0], items[0]] }));
  assert.deepEqual(output.rows, [1, 2, 3, 4, 5, 6, 7, 8]); assert.equal(output.cells.get('F8')!.formula, 'SUM(B2:B7)'); assert.equal(output.cells.get('F8')!.value, '21');
});
test('XLSX formula copying distinguishes relative, absolute, mixed and structural shifts', async () => {
  const input = await book([{ rows: row(1, n('B1', 10)) + row(2, t('A2', '{{#each items}}{{name}}') + t('B2', '{{quantity}}') + f('C2', 'B2+$B2+B$1+$B$1') + t('E2', '{{/each}}')) + row(3, n('B3', 50) + f('F3', '$B$3+B3')) }]);
  const output = await sheet(await render(input, { items }));
  assert.equal(output.cells.get('C3')!.formula, 'B3+$B3+B$1+$B$1'); assert.equal(output.cells.get('C3')!.value, '26');
  assert.equal(output.cells.get('F4')!.formula, '$B$4+B4'); assert.equal(output.cells.get('F4')!.value, '100');
});
test('XLSX cross-sheet dependencies and quoted sheet names are rewritten and recalculated', async () => {
  const input = await book([
    { name: "O'Brien!", rows: loop() + row(3, n('B3', 9)) },
    { name: 'Observer', rows: row(1, f('A1', "SUM('O''Brien!'!B2:B2)") + f('B1', "'O''Brien!'!$B$3")) },
  ]);
  const output = await sheet(await render(input, { items }), 1);
  assert.equal(output.cells.get('A1')!.formula, "SUM('O''Brien!'!B2:B3)"); assert.equal(output.cells.get('A1')!.value, '5');
  assert.equal(output.cells.get('B1')!.formula, "'O''Brien!'!$B$4"); assert.equal(output.cells.get('B1')!.value, '9');
});
test('XLSX tables, names, filters, validation and conditional formatting expand together', async () => {
  const input = await book([{ rows: row(1, t('A1', 'Header')) + loop() + row(3, n('D3', 9)), suffix: '<s:autoFilter ref="A1:D2"/><s:conditionalFormatting sqref="B2:B2"><s:cfRule type="expression" priority="1"><s:formula>B2&gt;0</s:formula></s:cfRule></s:conditionalFormatting><s:dataValidations count="1"><s:dataValidation sqref="C2:C2 D3:D3" type="whole"><s:formula1>1</s:formula1></s:dataValidation></s:dataValidations><s:tableParts count="1"><s:tablePart r:id="table"/></s:tableParts>' }], {
    'xl/worksheets/_rels/s0.xml.rels': rels(link('table', 'table', '../tables/t.xml')),
    'xl/tables/t.xml': `<table xmlns="${ns.spreadsheet}" id="1" name="Items" ref="A1:D2"><autoFilter ref="A1:D2"/><tableColumns count="1"><tableColumn id="1" name="Header"/></tableColumns></table>`,
  }, '<s:definedName name="AllItems">Main!$A$2:$D$2</s:definedName><s:definedName name="LocalItems" localSheetId="0">$B$2:$B$2</s:definedName>');
  const output = await sheet(await render(input, { items }));
  assert(output.source.includes('sqref="B2:B3"')); assert(output.source.includes('sqref="C2:C3 D4:D4"')); assert(output.source.includes('ref="A1:D3"'));
  assert(output.parts.get('xl/tables/t.xml')!.toString().includes('ref="A1:D3"'));
  const workbook = output.parts.get('xl/workbook.xml')!.toString(); assert(workbook.includes('Main!$A$2:$D$3')); assert(workbook.includes('$B$2:$B$3'));
});
test('XLSX merges below loops move and contained merges repeat while legacy multi-row markers are rejected', async () => {
  const output = await sheet(await render(await book([{ rows: loop() + row(3, t('A3', 'Merged')), suffix: '<s:mergeCells count="1"><s:mergeCell ref="A3:C3"/></s:mergeCells>' }]), { items }));
  assert(output.source.includes('ref="A4:C4"'));
  const merged = await sheet(await render(await book([{ rows: loop(), suffix: '<s:mergeCells><s:mergeCell ref="A2:C2"/></s:mergeCells>' }]), { items })); assert(merged.source.includes('ref="A3:C3"'));
  await assert.rejects(compile(await book([{ rows: row(2, t('A2', '{{#each items}}')) + row(3, t('A3', '{{/each}}')) }])), code('TEMPLATE_STRUCTURE'));
});
test('XLSX drawing anchors below loops move and pictures crossing a loop are rejected', async () => {
  const drawing = (start: number) => `<xdr:wsDr xmlns:xdr="${ns.spreadsheetDrawing}" xmlns:a="${ns.drawing}" xmlns:r="${ns.relationship}"><xdr:oneCellAnchor><xdr:from><xdr:col>1</xdr:col><xdr:colOff>0</xdr:colOff><xdr:row>${start}</xdr:row><xdr:rowOff>0</xdr:rowOff></xdr:from><xdr:ext cx="914400" cy="609600"/><xdr:pic><xdr:nvPicPr><xdr:cNvPr id="1" name="Logo" descr="{{image:logo}}"/></xdr:nvPicPr><xdr:blipFill><a:blip r:embed="image"/></xdr:blipFill></xdr:pic><xdr:clientData/></xdr:oneCellAnchor></xdr:wsDr>`;
  const input = (start: number) => book([{ rows: loop(), suffix: '<s:drawing r:id="drawing"/>' }], {
    'xl/worksheets/_rels/s0.xml.rels': rels(link('drawing', 'drawing', '../drawings/d.xml')),
    'xl/drawings/d.xml': drawing(start), 'xl/drawings/_rels/d.xml.rels': rels(link('image', 'image', '../media/old.png')), 'xl/media/old.png': png([0, 0, 0]),
  });
  const output = await readPackage(await render(await input(4), { items }, { attachments: { logo: png([255, 0, 0]) } }));
  assert(output.get('xl/drawings/d.xml')!.toString().includes('<xdr:row>5</xdr:row>')); assert(output.get('xl/drawings/d.xml')!.toString().includes('Attachment logo'));
  const repeated = await readPackage(await render(await input(1), { items: items.map(item => ({ ...item, logo: 'logo' })) }, { attachments: { logo: png([255,0,0]) } }));
  assert.equal(descendants(parseXml(repeated.get('xl/drawings/d.xml')!.toString(), 'test'), ns.spreadsheetDrawing, 'pic').length, 2);
});
test('XLSX obsolete calculation chain and its metadata are removed', async () => {
  const input = await book([{ rows: loop() }], {
    'xl/_rels/workbook.xml.rels': rels(link('s0', 'worksheet', 'worksheets/s0.xml') + link('chain', 'calcChain', 'calcChain.xml')),
    'xl/calcChain.xml': `<calcChain xmlns="${ns.spreadsheet}"><c r="D2" i="1"/></calcChain>`,
  });
  const parts = await readPackage(input), types = parts.get('[Content_Types].xml')!.toString().replace('</Types>', '<Override PartName="/xl/calcChain.xml" ContentType="application/vnd.openxmlformats-officedocument.spreadsheetml.calcChain+xml"/></Types>');
  parts.set('[Content_Types].xml', Buffer.from(types));
  const output = await readPackage(await render(await writePackage(parts), { items }));
  assert(!output.has('xl/calcChain.xml')); assert(!output.get('[Content_Types].xml')!.toString().includes('calcChain')); assert(!output.get('xl/_rels/workbook.xml.rels')!.toString().includes('calcChain'));
});
test('XLSX calculator resolves dependencies, aggregates, errors and lazy IF branches', async () => {
  const input = await book([{ rows: row(1, f('A1', 'SUM(D2:D2)') + f('B1', 'AVERAGE(B2:B2)') + f('C1', 'COUNT(B2:B2)') + f('D1', 'MAX(B2:B2)-MIN(B2:B2)') + f('E1', 'IF(A1>0,"yes",1/0)') + f('F1', '1/0')) + loop() }]);
  const output = await sheet(await render(input, { items }));
  assert.deepEqual(['A1', 'B1', 'C1', 'D1', 'E1', 'F1'].map(ref => output.cells.get(ref)!.value), ['85', '2.5', '2', '1', 'yes', '#DIV/0!']);
  assert.equal(output.cells.get('E1')!.type, 'str'); assert.equal(output.cells.get('F1')!.type, 'e');
});
test('XLSX formula string literals stay untouched when row references change', async () => {
  const input = await book([{ rows: loop() + row(3, f('A3', '"B3"&B3') + n('B3', 7)) }]);
  const output = await sheet(await render(input, { items }));
  assert.equal(output.cells.get('A4')!.formula, '"B3"&B4'); assert.equal(output.cells.get('A4')!.value, 'B37');
});
test('XLSX circular dependencies fail rather than use stale formula caches', async () => {
  const input = await book([{ rows: row(1, f('A1', 'B1') + f('B1', 'A1')) + loop() }]);
  await assert.rejects(render(input, { items }), code('FORMULA_CYCLE'));
});
test('XLSX unsupported formulas and special formula storage fail before row rendering', async () => {
  for (const formula of ['NOW()', 'SUM(Table1[Amount])', 'UnknownName', '[other.xlsx]Sheet1!A1', 'SUM(A:A)', '1E309', 'A1:A3', 'SUM(B3:A1)']) await assert.rejects(compile(await book([{ rows: row(1, f('A1', formula)) + loop() }])), code('UNSUPPORTED_FORMULA'));
  for (const storage of ['t="shared" si="0"', 't="array" ref="A1:A3"', 't="dataTable"']) await assert.rejects(compile(await book([{ rows: row(1, f('A1', '1+1', storage)) + loop() }])), code('UNSUPPORTED_FORMULA'));
});
test('XLSX row expansion is bounded and compiled workbooks can be reused concurrently', async () => {
  const compiled = await compile(await book([{ rows: loop() }]));
  await assert.rejects(compiled.render({ items: Array.from({ length: 10001 }, () => items[0]) }), code('LIMIT_EXCEEDED'));
  const results = await Promise.all([[], items, [items[0]]].map(items => compiled.render({ items })));
  assert.deepEqual(await Promise.all(results.map(async value => (await sheet(value)).rows.length)), [0, 2, 1]);
  assert.deepEqual(results[1], await compiled.render({ items }));
});

test('XLSX row bounds and formula operation budgets stop pathological expansion', async () => {
  await assert.rejects(render(await book([{ rows: loop(1048576) }]), { items }), code('LIMIT_EXCEEDED'));
  await assert.rejects(render(await book([{ rows: row(1, f('A1', 'SUM(B1:B1048576)')) + loop() }]), { items: [] }), code('LIMIT_EXCEEDED'));
});
test('XLSX deleting the only row leaves an empty worksheet with valid selection coordinates', async () => {
  const parts = await readPackage(await book([{ rows: loop(1), suffix: '<s:sheetViews><s:sheetView workbookViewId="0"><s:selection activeCell="A1" sqref="A1"/></s:sheetView></s:sheetViews>' }]));
  parts.set('xl/worksheets/s0.xml', Buffer.from(parts.get('xl/worksheets/s0.xml')!.toString().replace('ref="A1:F10"', 'ref="A1:F1"')));
  const output = await sheet(await render(await writePackage(parts), { items: [] }));
  assert.deepEqual(output.rows, []); assert(output.source.includes('dimension ref="A1"')); assert(output.source.includes('activeCell="A1" sqref="A1"'));
});

test('XLSX COUNT ignores direct error arguments as well as referenced errors', async () => {
  const output = await sheet(await render(await book([{ rows: row(1, f('A1', 'COUNT(1,1/0)')) + loop() }]), { items }));
  assert.equal(output.cells.get('A1')!.value, '1');
});

test('XLSX repeated output and deeply chained formulas have explicit budgets', async () => {
  const repeated = row(2, Array.from({ length: 20 }, (_, i) => t(`${String.fromCharCode(65 + i)}2`, (i === 0 ? '{{#each items}}' : '') + '{{text}}')).join('') + t('U2', '{{/each}}'));
  const compiled = await compile(await book([{ rows: repeated }]));
  await assert.rejects(compiled.render({ items: Array.from({ length: 25 }, () => ({ text: 'x'.repeat(20000) })) }), code('LIMIT_EXCEEDED'));
  assert.equal((await sheet(await compiled.render({ items: [{ text: 'small' }] }))).cells.get('A2')!.text, 'small');
  await assert.rejects(compile(await book([{ rows: row(1, f('A1', Array(300).fill('1').join('+'))) + loop() }])), code('LIMIT_EXCEEDED'));
});

test('XLSX calculation rejects ambiguous cross-application coercion rather than caching a guess', async () => {
  for (const formula of ['TRUE&FALSE', '0^0', '0^-1']) await assert.rejects(render(await book([{ rows: row(1, f('A1', formula)) + loop() }]), { items }), code('UNSUPPORTED_FORMULA'));
  const output = await sheet(await render(await book([{ rows: row(1, f('A1', 'IF("TRUE",1,2)')) + loop() }]), { items }));
  assert.equal(output.cells.get('A1')!.value, '1');
});

const standardRow = (number = 2, path = 'items') => row(number, t(`A${number}`, `{d.${path}[i].name} / {d.title}`) + t(`B${number}`, `{d.${path}[i].quantity}`) + t(`C${number}`, `{d.${path}[i].price}`) + f(`D${number}`, `B${number}*C${number}`)) + row(number + 1, t(`A${number + 1}`, `{d.${path}[i+1].name}`));
test('Standard XLSX rows discard the increment exemplar, preserve root scope and calculate typed formulas', async () => {
  const input = await book([{ rows: standardRow() + row(4, t('A4', 'After') + f('D4', 'SUM(D2:D3)')) }]);
  const compiled = await compile(input);
  for (const count of [0, 1, 3]) {
    const data = Array.from({ length: count }, (_, n) => ({ name: `Item ${n}`, quantity: n + 1, price: 10 }));
    const output = await sheet(await compiled.render({ title: 'ROOT', items: data }));
    assert.equal(output.rows.length, count + 1);
    assert.equal(output.cells.get(`A${count + 2}`)!.text, 'After');
    for (let n = 0; n < count; n++) {
      assert.equal(output.cells.get(`A${n + 2}`)!.text, `Item ${n} / ROOT`);
      assert.equal(output.cells.get(`B${n + 2}`)!.type, 'n');
      assert.equal(output.cells.get(`D${n + 2}`)!.value, String((n + 1) * 10));
    }
    if (count) assert.equal(output.cells.get(`D${count + 2}`)!.value, String(count * (count + 1) * 5));
    assert(!output.source.includes('{d.'));
  }
  assert.equal((await sheet(await compiled.render({}))).rows.length, 1);
  await assert.rejects(compiled.render({ items: 4 }), code('INVALID_DATA'));
  await assert.rejects(compiled.render({ items: [1] }), code('INVALID_DATA'));
  await assert.rejects(compiled.render({ items: Array.from({ length: 10001 }, () => ({})) }), code('LIMIT_EXCEEDED'));
});
test('Standard XLSX independent repetitions and cross-sheet formulas follow marker-row removal', async () => {
  const input = await book([{ rows: standardRow(2) + standardRow(6, 'other') + row(8, t('A8', 'End')) }, { name: 'Totals', rows: row(1, f('A1', 'SUM(Main!D2:D3)+SUM(Main!D6:D7)')) }]);
  const output = await render(input, { title: 'R', items, other: [items[0]] });
  const main = await sheet(output), totals = await sheet(output, 1);
  assert.equal(main.cells.get('A7')!.text, 'End');
  assert.equal(totals.cells.get('A1')!.value, '110');
});
test('Standard XLSX shared rich strings have independent item/root scopes across repeated and ordinary cells', async () => {
  const sharedCell = (ref: string, index: number) => `<s:c r="${ref}" t="s"><s:v>${index}</s:v></s:c>`;
  const input = await book([{ rows: row(1, sharedCell('A1', 1)) + row(2, sharedCell('A2', 0) + sharedCell('B2', 1)) + row(3, t('A3', '{d.items[i+1].name}')) + row(4, sharedCell('A4', 1)) }], {
    'xl/_rels/workbook.xml.rels': rels(link('s0', 'worksheet', 'worksheets/s0.xml') + link('sst', 'sharedStrings', 'sharedStrings.xml')),
    'xl/sharedStrings.xml': `<q:sst xmlns:q="${ns.spreadsheet}" count="4" uniqueCount="2"><q:si><q:r><q:rPr><q:b/></q:rPr><q:t>{d.items[i].na</q:t></q:r><q:r><q:t>me}</q:t></q:r></q:si><q:si><q:t>{d.title}</q:t></q:si></q:sst>`,
  });
  const output = await sheet(await render(input, { title: 'ROOT', items }));
  assert.deepEqual(['A1', 'B2', 'B3', 'A4'].map(ref => output.cells.get(ref)!.text), Array(4).fill('ROOT'));
  assert.equal(output.cells.get('A2')!.text, 'First'); assert.equal(output.cells.get('A3')!.text, 'Second');
  assert(output.source.includes('<q:b/>')); assert(!output.parts.get('xl/sharedStrings.xml')!.toString().includes('{d.'));
});
test('Standard XLSX rejects unsupported nesting, unmatched or distant markers and mixed repeat syntax', async () => {
  for (const rows of [row(2, t('A2', '{d.items[i].name}')), row(2, t('A2', '{d.items[i+1].name}')), row(2, t('A2', '{d.items[i].name}')) + row(4, t('A4', '{d.items[i+1].name}')), row(2, t('A2', '{d.items[i].children[i].name}')) + row(3, t('A3', '{d.items[i].children[i+1].name}')), standardRow() + loop(5)]) {
    await assert.rejects(compile(await book([{ rows }])), error => error instanceof RenderError && ['UNSUPPORTED_EXPRESSION', 'TEMPLATE_STRUCTURE', 'UNSUPPORTED_TEMPLATE'].includes(error.code));
  }
});
test('Standard XLSX missing values become blank while false, zero, multiline text and literal tags survive', async () => {
  const input = await book([{ rows: row(1, t('A1', '{d.zero}') + t('B1', '{d.flag}') + t('C1', '{d.missing}') + t('D1', '{d.text}')) }]);
  const output = await sheet(await render(input, { zero: 0, flag: false, text: 'A & B\n\t{d.literal}' }));
  assert.equal(output.cells.get('A1')!.value, '0'); assert.equal(output.cells.get('B1')!.type, 'b'); assert.equal(output.cells.get('B1')!.value, '0');
  assert.equal(output.cells.get('C1')!.value, undefined); assert.equal(output.cells.get('D1')!.text, 'A & B\n\t{d.literal}');
});
test('Standard XLSX formulas and numeric payloads cannot silently contain unresolved tags', async () => {
  for (const cell of [f('A1', '{d.formula}'), '<s:c r="A1"><s:v>{d.value}</s:v></s:c>']) await assert.rejects(compile(await book([{ rows: row(1, cell) }])), code('UNSUPPORTED_TEMPLATE'));
});

test('Standard scalar paths keep root meaning inside a legacy XLSX loop', async () => {
  const input = await book([{ rows: row(2, t('A2', '{{#each items}}{d.title}: {{name}}') + t('B2', '{d.total}') + t('C2', '{{/each}}')) }]);
  const output = await sheet(await render(input, { title: 'ROOT', total: 7, items: [{ title: 'wrong scope', name: 'item', total: 99 }] }));
  assert.equal(output.cells.get('A2')!.text, 'ROOT: item'); assert.equal(output.cells.get('B2')!.value, '7');
});

test('Standard multi-row blocks interleave complete items and clone contained merges', async () => {
  const rows = row(1,t('A1','Header')) + row(2,t('A2','{d.items[i].name}') + n('B2',2)) + row(3,t('A3','Detail {d.items[i].name}') + f('B3','SUM(B2:B2)')) + row(4,t('A4','{d.items[i+1].name}')) + row(5,f('B5','SUM(B2:B4)'));
  const template = await book([{ rows, suffix: '<s:mergeCells count="1"><s:mergeCell ref="C2:D3"/></s:mergeCells>' }]);
  const output = await sheet(await render(template, { items }));
  assert.deepEqual(output.rows,[1,2,3,4,5,6]);
  assert.deepEqual(['A2','A3','A4','A5'].map(r=>output.cells.get(r)?.text),['First','Detail First','Second','Detail Second']);
  assert.equal(output.cells.get('B5')?.formula,'SUM(B4:B4)'); assert.equal(output.cells.get('B6')?.value,'8');
  assert(output.source.includes('ref="C4:D5"'));
  const empty = await sheet(await render(template,{items:[]})); assert.deepEqual(empty.rows,[1,2]); assert(!empty.source.includes('mergeCells'));
});
test('Additional formula functions calculate rounding, text, conditions and lazy errors', async () => {
  const formulas = ['ROUND(1.005,2)','ROUND(-1.5,0)','ROUNDUP(-1.21,1)','ROUNDDOWN(1.29,1)','ABS(-2)','LEN("text")','IFERROR(1/0,42)','AND(TRUE,1)','OR(FALSE,0)','NOT(0)','CONCAT("A",1,TRUE)'];
  const output = await sheet(await render(await book([{ rows: row(1,formulas.map((value,i)=>f(`${String.fromCharCode(65+i)}1`,value)).join('')) + loop() }]),{items:[]}));
  assert.deepEqual(formulas.map((_,i)=>output.cells.get(`${String.fromCharCode(65+i)}1`)?.value),['1.01','-2','-1.3','1.2','2','4','42','1','0','1','A1TRUE']);
});
