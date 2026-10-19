import test from 'node:test';
import assert from 'node:assert/strict';
import { readFile } from 'node:fs/promises';
import sharp from 'sharp';
import { readPackage, writePackage } from '../src/package.ts';
import { snapshotDocument, compareDocuments } from '../src/document-diff.ts';
import { comparePageImages } from '../src/visual-diff.ts';
import { createReplayBundle, replayBundle } from '../src/replay.ts';
import { runRegressionSuite } from '../src/regression.ts';
import { encodeTextTemplate } from '../src/text-template.ts';
test('semantic Word comparison distinguishes line breaks and tabs from adjacent text', async () => {
  const parts = await readPackage(await readFile('examples/invoice/invoice.docx'));
  const xml = (inner: string) =>
    Buffer.from(
      `<w:document xmlns:w="http://schemas.openxmlformats.org/wordprocessingml/2006/main"><w:body><w:p><w:r>${inner}</w:r></w:p></w:body></w:document>`,
    );
  parts.set('word/document.xml', xml('<w:t>A</w:t><w:t>B</w:t>'));
  const before = await snapshotDocument(await writePackage(parts), 'docx');
  for (const control of ['<w:br/>', '<w:tab/>', '<w:cr/>']) {
    parts.set('word/document.xml', xml(`<w:t>A</w:t>${control}<w:t>B</w:t>`));
    const after = await snapshotDocument(await writePackage(parts), 'docx');
    assert.equal(compareDocuments(before, after).equal, false, control);
  }
});
test('semantic spreadsheet comparison retains formula extent and shared identity', async () => {
  const parts = await readPackage(await readFile('examples/invoice/invoice.xlsx'));
  const path = [...parts.keys()].find((p) => /^xl\/worksheets\/[^/]+\.xml$/.test(p))!;
  const xml = (attrs: string) =>
    Buffer.from(
      `<worksheet xmlns="http://schemas.openxmlformats.org/spreadsheetml/2006/main"><sheetData><row r="1"><c r="A1"><f ${attrs}>B1*2</f><v>4</v></c></row></sheetData></worksheet>`,
    );
  parts.set(path, xml('t="shared" si="0" ref="A1:A2"'));
  const before = await snapshotDocument(await writePackage(parts), 'xlsx');
  for (const attrs of [
    't="shared" si="1" ref="A1:A2"',
    't="shared" si="0" ref="A1:A3"',
    't="array" ref="A1:A2"',
  ]) {
    parts.set(path, xml(attrs));
    assert.equal(
      compareDocuments(before, await snapshotDocument(await writePackage(parts), 'xlsx')).equal,
      false,
      attrs,
    );
  }
});
test('visual comparisons independently count changed pixels at exact threshold boundaries', async () => {
  const png = (values: Buffer) =>
    sharp(values, { raw: { width: 2, height: 2, channels: 4 } })
      .png()
      .toBuffer();
  const data = Buffer.alloc(16, 255),
    a = await png(data);
  for (let changed = 1; changed <= 4; changed++) {
    const altered = Buffer.from(data);
    for (let p = 0; p < changed; p++) altered[p * 4] = 253;
    const b = await png(altered),
      diff = await comparePageImages(a, b, { channelTolerance: 1 });
    assert.equal(diff.changedPixels, changed);
    assert.equal(diff.changedRatio, changed / 4);
    assert(!diff.passed);
    assert.equal((await comparePageImages(a, b, { channelTolerance: 2 })).changedPixels, 0);
    if (changed < 4) assert((await comparePageImages(a, b, { maxChangedRatio: changed / 4 })).passed);
  }
  await assert.rejects(comparePageImages(a, a, { channelTolerance: 255 }));
  await assert.rejects(
    comparePageImages(a, a, {
      masks: [
        { left: 0, top: 0, width: 1, height: 2 },
        { left: 1, top: 0, width: 1, height: 2 },
      ],
    }),
  );
});
test('replay cannot silently discard hidden data or hidden render options', async () => {
  const template = encodeTextTemplate('txt', '{d.name}');
  const data = Object.defineProperty({}, 'name', { value: 'hidden', enumerable: false });
  await assert.rejects(createReplayBundle(template, data));
  const options = Object.defineProperty({}, 'locale', { value: 'de-DE', enumerable: false });
  await assert.rejects(createReplayBundle(template, { name: 'x' }, options));
});
test('replay snapshots caller data and regression defaults reject differing bytes', async () => {
  const template = encodeTextTemplate('txt', '{d.name}'),
    data = { name: 'initial' };
  const pending = createReplayBundle(template, data);
  data.name = 'mutated';
  const bundle = await pending;
  assert.equal((await replayBundle(bundle)).toString(), 'initial');
  const suite = await runRegressionSuite(encodeTextTemplate('txt', 'A\nB'), [
    { name: 'line-ending', data: {}, approved: Buffer.from('A\r\nB') },
  ]);
  assert(!suite.passed);
  assert(suite.results[0]!.diff!.equal);
  assert(!suite.results[0]!.diff!.binaryEqual);
});
