import test from 'node:test';
import assert from 'node:assert/strict';
import { buildDocx } from '../docx.ts';
import { buildXlsx } from '../xlsx.ts';
import { inputData, recipe } from '../recipes.ts';
import { compare, validatePackage, assertGolden } from '../inspect.ts';
import { pack, unpack, perturb, png, type Parts, type Format } from '../common.ts';

const preserved = (format: Format) => [format === 'docx' ? 'word/styles.xml' : 'xl/styles.xml', 'customXml/fixture.xml'];
function mutate(parts: Parts, name: string, from: string, to: string) {
  const out = new Map(parts), source = out.get(name)!.toString('utf8');
  assert(source.includes(from), `Mutation must apply: ${from}`);
  out.set(name, Buffer.from(source.replace(from, to))); return out;
}
test('100 recipes per format have unique IDs and cumulative severity', () => {
  for (const format of ['docx', 'xlsx'] as Format[]) {
    const recipes = Array.from({ length: 100 }, (_, i) => recipe(format, i + 1));
    assert.equal(new Set(recipes.map(r => r.id)).size, 100);
    for (let i = 1; i < recipes.length; i++) {
      const prev = recipes[i - 1]!.dimensions, next = recipes[i]!.dimensions;
      assert.equal(next.reduce((a, b) => a + b, 0), i + 1);
      next.forEach((v, j) => assert(v >= prev[j]!));
    }
  }
});
for (const format of ['docx', 'xlsx'] as Format[]) {
  const builder = format === 'docx' ? buildDocx : buildXlsx;
  test(`${format}: every curriculum boundary has valid expected semantics`, () => {
    for (const level of [1, 10, 11, 20, 21, 30, 31, 40, 41, 50, 51, 60, 61, 70, 71, 80, 81, 90, 91, 100]) {
      const data = inputData(level), parts = builder(level, data, true);
      validatePackage(parts); assertGolden(parts, format, level, data);
    }
  });
  test(`${format}: unchanged template fails`, () => {
    const data = inputData(100);
    assert.throws(() => compare(builder(100, data, false), builder(100, data, true), format, preserved(format)));
  });
  test(`${format}: ZIP entry ordering and namespace aliases do not affect semantics`, async () => {
    const data = inputData(80), expected = builder(80, data, true);
    const mutated = perturb(expected, 9);
    compare(await unpack(await pack(mutated, 7)), expected, format, []);
  });
  test(`${format}: deterministic ZIP output`, async () => {
    const parts = builder(25, inputData(25), true);
    assert.deepEqual(await pack(parts), await pack(parts));
  });
  test(`${format}: replacing an image with the wrong image is detected`, () => {
    const level = 30, expected = builder(level, inputData(level), true), actual = new Map(expected);
    actual.set(`${format === 'docx' ? 'word' : 'xl'}/media/red.png`, png([0, 255, 0]));
    assert.throws(() => compare(actual, expected, format, preserved(format)));
  });
  test(`${format}: missing image relationships are detected`, () => {
    const expected = builder(30, inputData(30), true), actual = new Map(expected);
    actual.delete(`${format === 'docx' ? 'word' : 'xl'}/media/red.png`);
    assert.throws(() => validatePackage(actual), /Broken relationship/);
  });
  test(`${format}: unrelated-part loss is detected`, () => {
    const expected = builder(1, inputData(1), true), actual = new Map(expected); actual.delete('customXml/fixture.xml');
    assert.throws(() => compare(actual, expected, format, preserved(format)), /Preservation failure/);
  });
}
test('DOCX: incorrect text and removed bold formatting are detected', () => {
  const expected = buildDocx(40, inputData(40), true);
  assert.throws(() => compare(mutate(expected, 'word/document.xml', 'Zoë', 'Wrong'), expected, 'docx', []));
  assert.throws(() => compare(mutate(expected, 'word/document.xml', '<w:b/>', ''), expected, 'docx', []));
});
test('XLSX: stale formula, incorrect numeric type and missing rich formatting are detected', () => {
  const expected = buildXlsx(60, inputData(60), true);
  assert.throws(() => compare(mutate(expected, 'xl/worksheets/sheet1.xml', 'B41*C41', 'B40*C40'), expected, 'xlsx', []));
  assert.throws(() => compare(mutate(expected, 'xl/worksheets/sheet1.xml', 'r="B2" s="0"', 'r="B2" s="0" t="b"'), expected, 'xlsx', []));
  assert.throws(() => compare(mutate(expected, 'xl/sharedStrings.xml', '<b/>', ''), expected, 'xlsx', []));
});
test('XLSX: stale dependent table range is detected', () => {
  const expected = buildXlsx(80, inputData(80), true);
  assert.throws(() => compare(mutate(expected, 'xl/tables/table1.xml', 'A39:D51', 'A39:D40'), expected, 'xlsx', []));
});
test('Malformed XML cannot pass validation', () => {
  const expected = buildDocx(1, inputData(1), true);
  assert.throws(() => validatePackage(mutate(expected, 'word/document.xml', '</w:body>', '</w:wrong>')));
});
test('stored ZIP entries include sizes in local headers for Office interoperability', async () => {
  const bytes = await pack(new Map([['sample.xml', Buffer.from('<sample/>')]]), 1);
  assert.equal(bytes.readUInt16LE(8), 0); // Stored compression method.
  assert.equal(bytes.readUInt16LE(6) & 8, 0); // No streaming data descriptor.
  assert.equal(bytes.readUInt32LE(22), Buffer.byteLength('<sample/>'));
});

test('DOCX: duplicate drawing identities are detected independently of picture bytes', () => {
  const expected = buildDocx(60, inputData(60), true);
  assert.throws(() => validatePackage(mutate(expected, 'word/document.xml', '<wp:docPr id="2"', '<wp:docPr id="1"')), /Duplicate drawing ID/);
});
