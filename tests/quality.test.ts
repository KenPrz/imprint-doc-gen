import test from 'node:test';
import assert from 'node:assert/strict';
import { readFile } from 'node:fs/promises';
import { encodeTextTemplate, render } from '../src/index.ts';
import { compareDocuments, snapshotDocument, sha256 } from '../src/document-diff.ts';
import { runRegressionSuite } from '../src/regression.ts';
import { createReplayBundle, replayBundle } from '../src/replay.ts';
import { readPackage, writePackage } from '../src/package.ts';
const snapshot = (text: string) => snapshotDocument(Buffer.from(text), 'txt');
const canonical = (value: unknown): unknown =>
  Array.isArray(value)
    ? value.map(canonical)
    : value && typeof value === 'object'
      ? Object.fromEntries(
          Object.entries(value)
            .sort(([a], [b]) => (a < b ? -1 : a > b ? 1 : 0))
            .map(([k, v]) => [k, canonical(v)]),
        )
      : value;
test('semantic differences preserve Unicode, whitespace, additions and removals; masks fail closed', async () => {
  const a = await snapshot('Zoë\n10\nold'),
    b = await snapshot('Zoë\n20');
  assert.deepEqual(compareDocuments(a, b).changes, [
    { location: 'text/line/2', kind: 'text', before: '10', after: '20' },
    { location: 'text/line/3', kind: 'text', before: 'old' },
  ]);
  assert.equal(compareDocuments(a, a).equal, true);
  assert.equal(compareDocuments(await snapshot('a '), await snapshot('a')).equal, false);
  assert.equal(compareDocuments(await snapshot('a\r\nb'), await snapshot('a\nb')).equal, true);
  assert.equal(compareDocuments(await snapshot('a\r\nb'), await snapshot('a\nb')).binaryEqual, false);
  assert.equal(compareDocuments(a, b, ['text/line/2']).changes.length, 1);
  assert.throws(() => compareDocuments(a, b, ['text/line/3']), /both/);
  assert.throws(() => compareDocuments(a, b, ['text\/line\/.*']), /both/);
  assert.throws(() => compareDocuments(a, b, ['text/line/2', 'text/line/2']), /Invalid/);
  assert.throws(() => compareDocuments(a, { ...b, format: 'csv' }), /different formats/);
  await assert.rejects(snapshotDocument(Buffer.from([255]), 'txt'));
});
test('CSV compares logical rows with embedded newlines and escaped quotes', async () => {
  const a = await snapshotDocument(Buffer.from('name,note\r\n"Zoë","line1\nline2 ""ok"""\r\n'), 'csv');
  assert.equal(a.units.length, 2);
  assert.deepEqual(JSON.parse(a.units[1]!.value), ['Zoë', 'line1\nline2 "ok"']);
  const b = await snapshotDocument(Buffer.from('name,note\nZoë,"line1\nline2 ""ok"""'), 'csv');
  assert.equal(compareDocuments(a, b).equal, true);
  for (const invalid of ['"unclosed', 'a"b', '"a"b', '"a""'])
    await assert.rejects(snapshotDocument(Buffer.from(invalid), 'csv'));
  assert.equal((await snapshotDocument(Buffer.from('a,'), 'csv')).units[0]!.value, '["a",""]');
});
test('Office snapshots detect changed values and preserve binary-only differences', async () => {
  const template = await readFile('examples/invoice/invoice.docx');
  const parts = await readPackage(template);
  const a = await snapshotDocument(template, 'docx');
  assert.ok(a.units.length > 0);
  parts.set(
    'word/document.xml',
    Buffer.from(
      parts
        .get('word/document.xml')!
        .toString()
        .replace('</w:body>', '<w:p><w:r><w:t>Added total: 10</w:t></w:r></w:p></w:body>'),
    ),
  );
  const b = await snapshotDocument(await writePackage(parts), 'docx');
  assert.ok(compareDocuments(a, b).changes.some((c) => c.after === 'Added total: 10'));
  parts.set('word/document.xml', (await readPackage(template)).get('word/document.xml')!);
  parts.set('custom-test.txt', Buffer.from('asset change'));
  const c = await snapshotDocument(await writePackage(parts), 'docx');
  assert.equal(compareDocuments(a, c).equal, true);
  assert.equal(compareDocuments(a, c).binaryEqual, false);
});
test('regression suite executes rendering, defaults to exact acceptance and continues after failures', async () => {
  const template = encodeTextTemplate('txt', 'Header\nHello {d.name}');
  const suite = await runRegressionSuite(template, [
    { name: 'approved', data: { name: 'World' }, approved: Buffer.from('Header\nHello World') },
    { name: 'changed', data: { name: 'New' }, approved: Buffer.from('Header\nHello World') },
    { name: 'malformed baseline', data: {}, approved: Buffer.from([255]) },
  ]);
  assert.equal(suite.passed, false);
  assert.deepEqual(
    suite.results.map((r) => r.passed),
    [true, false, false],
  );
  assert.ok(suite.results[2]!.error);
  const masked = {
    name: 'timestamp',
    data: { name: 'New' },
    approved: Buffer.from('Header\nHello World'),
    masks: ['text/line/2'],
  };
  await assert.rejects(runRegressionSuite(template, [masked]), /explicit semantic/);
  assert.equal((await runRegressionSuite(template, [masked], 'semantic')).passed, true);
  await assert.rejects(runRegressionSuite(template, []), /1–1000/);
  await assert.rejects(runRegressionSuite(template, [masked, masked], 'semantic'), /unique/);
});
test('replay round-trips JSON-serialized bundle, detects corruption, engine and output mismatch', async () => {
  const template = encodeTextTemplate('txt', '{d.person.name}: {d.amount}');
  const bundle = await createReplayBundle(template, { person: { name: 'Zoë' }, amount: 12 });
  assert.equal((await replayBundle(JSON.parse(JSON.stringify(bundle)))).toString(), 'Zoë: 12');
  const changed = structuredClone(bundle);
  changed.payload.data.amount = 13;
  await assert.rejects(replayBundle(changed), /integrity/);
  changed.sha256 = sha256(JSON.stringify(canonical(changed.payload)));
  await assert.rejects(replayBundle(changed), /output differs/);
  const mismatch = structuredClone(bundle);
  mismatch.payload.engine = 'different';
  mismatch.sha256 = sha256(JSON.stringify(canonical(mismatch.payload)));
  await assert.rejects(replayBundle(mismatch), /engine mismatch/);
  await assert.rejects(replayBundle({ ...bundle, extra: true }), /envelope/);
  await assert.rejects(replayBundle(null), /envelope/);
});
test('replay rejects lossy/non-JSON inputs and preserves attachments', async () => {
  const template = encodeTextTemplate('txt', '{d.value}');
  for (const value of [undefined, NaN, Infinity, new Date(), new Uint8Array([1]), 1n, () => 1, new Array(2)])
    await assert.rejects(createReplayBundle(template, { value }));
  const cycle: Record<string, unknown> = {};
  cycle.self = cycle;
  await assert.rejects(createReplayBundle(template, cycle), /circular/);
  const data = Object.defineProperty({}, 'value', {
    get() {
      throw new Error('getter ran');
    },
    enumerable: true,
  });
  await assert.rejects(createReplayBundle(template, data), /getters/);
  const bundle = await createReplayBundle(
    template,
    { value: 'ok' },
    { attachments: { 'file.txt': Buffer.from('asset') }, locale: 'en-US', timeZone: 'UTC' },
  );
  assert.equal((await replayBundle(bundle)).toString(), 'ok');
  assert.equal((bundle.payload.options.attachments as Record<string, string>)['file.txt'], 'YXNzZXQ=');
  assert.equal((await render(template, { value: 'ok' })).toString(), 'ok');
});
test('independent fixtures for PPTX/ODT/ODS/ODP expose known text and value changes', async () => {
  const { fixture } = await import('../scripts/formats/fixtures.ts');
  for (const format of ['pptx', 'odt', 'ods', 'odp'] as const) {
    const bytes = await fixture(format, 1, true),
      before = await snapshotDocument(bytes, format);
    assert.ok(
      before.units.some((u) => u.value.includes('Example document 1')),
      format,
    );
    const parts = await readPackage(bytes);
    for (const [path, content] of parts)
      if (path.endsWith('.xml'))
        parts.set(path, Buffer.from(content.toString().replaceAll('Example document 1', 'Changed title 42')));
    const after = await snapshotDocument(await writePackage(parts), format);
    const diff = compareDocuments(before, after);
    assert.equal(diff.equal, false, format);
    assert.ok(
      diff.changes.some(
        (c) => c.before?.includes('Example document 1') && c.after?.includes('Changed title 42'),
      ),
      format,
    );
    assert.equal(diff.totalChanges, diff.changes.length);
    assert.equal(diff.truncated, false);
  }
});
test('XLSX compares cell coordinates, explicit value types, cached values and formulas', async () => {
  const parts = await readPackage(await readFile('examples/invoice/invoice.xlsx'));
  const path = [...parts.keys()].find((p) => /^xl\/worksheets\/[^/]+\.xml$/.test(p))!;
  parts.set(
    path,
    Buffer.from(
      '<worksheet xmlns="http://schemas.openxmlformats.org/spreadsheetml/2006/main"><sheetData><row r="1"><c r="A1"><v>12</v></c><c r="B1"><f>A1*2</f><v>24</v></c><c r="C1" t="inlineStr"><is><t>Zoë</t></is></c></row></sheetData></worksheet>',
    ),
  );
  const before = await snapshotDocument(await writePackage(parts), 'xlsx');
  assert.deepEqual(JSON.parse(before.units.find((u) => u.location === `${path}/cell/B1`)!.value), {
    type: 'n',
    value: '24',
    formula: 'A1*2',
    formulaAttributes: [],
  });
  assert.equal(JSON.parse(before.units.find((u) => u.location === `${path}/cell/C1`)!.value).value, 'Zoë');
  parts.set(
    path,
    Buffer.from(parts.get(path)!.toString().replace('A1*2', 'A1*3').replace('<v>24</v>', '<v>36</v>')),
  );
  assert.equal(
    compareDocuments(before, await snapshotDocument(await writePackage(parts), 'xlsx')).changes.length,
    1,
  );
  await assert.rejects(snapshotDocument(await writePackage(parts), 'docx'), /format does not match/);
  parts.set(path, Buffer.from(parts.get(path)!.toString().replace('t="inlineStr"', 't="s"')));
  await assert.rejects(snapshotDocument(await writePackage(parts), 'xlsx'), /shared string/);
});
test('malformed snapshots, masks covering everything, replay options and complexity fail closed', async () => {
  const s = await snapshot('one');
  assert.throws(() => compareDocuments(s, s, ['text/line/1']), /every semantic unit/);
  assert.throws(
    () => compareDocuments({ ...s, units: [s.units[0]!, s.units[0]!] }, s),
    /Invalid document snapshot/,
  );
  const template = encodeTextTemplate('txt', 'ok');
  for (const options of [
    { restartNumbering: 'true' },
    { locale: 'no_such_locale' },
    { timeZone: 'not a zone' },
    { images: [] },
    { images: { logo: { rotate: 361 } } },
    { images: { logo: { width: -1 } } },
    { charts: { '../../evil': { categories: [], series: [] } } },
    { unrecognized: true },
  ])
    await assert.rejects(createReplayBundle(template, {}, options as never));
  let deep: Record<string, unknown> = {};
  for (let i = 0; i < 70; i++) deep = { child: deep };
  await assert.rejects(createReplayBundle(template, deep), /complexity/);
  await assert.rejects(createReplayBundle(Buffer.alloc(8 * 1024 * 1024 + 1), {}), /template/);
  const bundle = await createReplayBundle(template, {});
  bundle.payload.template = '../';
  bundle.sha256 = sha256(JSON.stringify(canonical(bundle.payload)));
  await assert.rejects(replayBundle(bundle), /base64/);
});
test('PNG visual comparison checks RGB and alpha, thresholds, bounded masks and dimensions', async () => {
  const sharp = (await import('sharp')).default;
  const { comparePageImages } = await import('../src/visual-diff.ts');
  const raw = Buffer.alloc(3 * 2 * 4, 255);
  const png = (bytes: Buffer, width = 3, height = 2) =>
    sharp(bytes, { raw: { width, height, channels: 4 } })
      .png()
      .toBuffer();
  const a = await png(raw),
    altered = Buffer.from(raw);
  altered[0] = 240;
  altered[7] = 200;
  const b = await png(altered);
  assert.equal((await comparePageImages(a, a)).passed, true);
  const diff = await comparePageImages(a, b);
  assert.equal(diff.changedPixels, 2);
  assert.equal(diff.changedRatio, 2 / 6);
  assert.equal(diff.passed, false);
  assert.equal((await comparePageImages(a, b, { channelTolerance: 15 })).changedPixels, 1);
  assert.equal((await comparePageImages(a, b, { maxChangedRatio: 1 / 3 })).passed, true);
  const masked = await comparePageImages(a, b, { masks: [{ left: 0, top: 0, width: 2, height: 1 }] });
  assert.equal(masked.maskedPixels, 2);
  assert.equal(masked.comparedPixels, 4);
  assert.equal(masked.passed, true);
  for (const options of [
    { maxChangedRatio: NaN },
    { channelTolerance: 256 },
    { maxChangedRatio: 1 },
    { masks: [{ left: 0, top: 0, width: 3, height: 2 }] },
    { masks: [{ left: 2, top: 0, width: 2, height: 1 }] },
    { masks: [{ left: 0, top: 0, width: 0, height: 1 }] },
  ])
    await assert.rejects(comparePageImages(a, b, options));
  await assert.rejects(comparePageImages(a, await png(Buffer.alloc(4), 1, 1)), /dimensions/);
  await assert.rejects(comparePageImages(a, Buffer.from('not a png')));
  const jpeg = await sharp(a).jpeg().toBuffer();
  await assert.rejects(comparePageImages(a, jpeg), /PNG/);
});
test('replay rejects array/options accessors without invoking them and noncanonical base64', async () => {
  const template = encodeTextTemplate('txt', 'ok');
  const getter = () => {
    throw new Error('accessor invoked');
  };
  const array = Object.defineProperty([1], '0', { get: getter });
  await assert.rejects(createReplayBundle(template, { array }), /does not support getters/);
  const options = Object.defineProperty({}, 'locale', { get: getter, enumerable: true });
  await assert.rejects(createReplayBundle(template, {}, options), /does not support getters/);
  const attachments = Object.defineProperty({}, 'image.png', { get: getter, enumerable: true });
  await assert.rejects(createReplayBundle(template, {}, { attachments }), /does not support getters/);
  await assert.rejects(createReplayBundle(template, { [Symbol('key')]: 'value' }), /symbol keys/);
  const bundle = await createReplayBundle(template, {}, { attachments: { file: Buffer.from('x') } });
  (bundle.payload.options.attachments as Record<string, string>).file = 'eB==';
  bundle.sha256 = sha256(JSON.stringify(canonical(bundle.payload)));
  await assert.rejects(replayBundle(bundle), /base64/);
});
