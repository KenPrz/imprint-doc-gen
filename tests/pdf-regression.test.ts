import test from 'node:test';
import assert from 'node:assert/strict';
import { access, mkdtemp, writeFile, rm } from 'node:fs/promises';
import { join } from 'node:path';
import { tmpdir } from 'node:os';
import { PDFDocument, StandardFonts, rgb } from 'pdf-lib';
import { comparePdfDocuments } from '../src/pdf-regression.ts';
const pdftoppmPath = process.env.IMPRINT_TEST_PDFTOPPM ?? '/usr/bin/pdftoppm';
// CI installs Poppler; this suite deliberately fails when its independent renderer is missing.
async function pdf(text = 'Approved invoice', pages = 1, width = 240, x = 20) {
  const doc = await PDFDocument.create(),
    font = await doc.embedFont(StandardFonts.Helvetica);
  for (let i = 0; i < pages; i++) {
    const page = doc.addPage([width, 120]);
    page.drawText(text, { x, y: 75, size: 12, font });
    page.drawRectangle({ x: 20, y: 20, width: 40, height: 20, color: rgb(0, 0, 1) });
  }
  return doc.save();
}
test('real Poppler PDF regression compares independently generated PDFs and records tool version', async () => {
  await access(pdftoppmPath);
  const approved = await pdf();
  const same = await comparePdfDocuments(approved, await pdf(), { pdftoppmPath, dpi: 72 });
  assert.equal(same.passed, true);
  assert.equal(same.beforePages, 1);
  assert.match(same.renderer.version, /^pdftoppm version/);
  assert.equal(same.pages[0]!.diff!.comparedPixels, 240 * 120);
  const changed = await comparePdfDocuments(approved, await pdf('Different amount'), {
    pdftoppmPath,
    dpi: 72,
  });
  assert.equal(changed.passed, false);
  assert.ok(changed.pages[0]!.diff!.changedPixels > 0);
  const moved = await comparePdfDocuments(approved, await pdf('Approved invoice', 1, 240, 30), {
    pdftoppmPath,
    dpi: 72,
  });
  assert.equal(moved.passed, false);
  const masked = await comparePdfDocuments(approved, await pdf('Different amount'), {
    pdftoppmPath,
    dpi: 72,
    pageMasks: [{ page: 1, masks: [{ left: 10, top: 25, width: 220, height: 30 }] }],
  });
  assert.equal(masked.passed, true);
});
test('PDF page counts, dimensions, budgets and masks fail closed', async () => {
  const approved = await pdf();
  const count = await comparePdfDocuments(approved, await pdf('Approved invoice', 2), { pdftoppmPath });
  assert.equal(count.passed, false);
  assert.equal(count.reason, 'Page counts differ');
  const size = await comparePdfDocuments(approved, await pdf('Approved invoice', 1, 250), { pdftoppmPath });
  assert.equal(size.passed, false);
  assert.equal(size.pages[0]!.reason, 'Page dimensions differ');
  for (const options of [
    { maxPages: 0 },
    { maxTotalPixels: 1 },
    { dpi: 201 },
    { channelTolerance: 255 },
    { maxChangedRatio: 1 },
    { timeoutMs: NaN },
    { pageMasks: [{ page: 2, masks: [] }] },
    { pdftoppmPath: 'pdftoppm' },
  ])
    await assert.rejects(comparePdfDocuments(approved, approved, { pdftoppmPath, ...options }));
  await assert.rejects(
    comparePdfDocuments(approved, await pdf('text', 2), { pdftoppmPath, maxPages: 1 }),
    /page budget/,
  );
  await assert.rejects(
    comparePdfDocuments(approved, approved, {
      pdftoppmPath,
      dpi: 72,
      pageMasks: [{ page: 1, masks: [{ left: 0, top: 0, width: 240, height: 120 }] }],
    }),
    /entire page/,
  );
  await assert.rejects(
    comparePdfDocuments(Buffer.from('not a PDF'), approved, { pdftoppmPath }),
    /PDF inputs/,
  );
  await assert.rejects(comparePdfDocuments(Buffer.from('%PDF-corrupt'), approved, { pdftoppmPath }));
});
test('PDF rasterizer errors, impersonation, deadlines and mutation are contained', async () => {
  const dir = await mkdtemp(join(tmpdir(), 'imprint-pdf-test-'));
  try {
    const path = join(dir, 'tool');
    const approved = await pdf();
    await writeFile(path, '#!/bin/sh\nexec sleep 30\n', { mode: 0o700 });
    await assert.rejects(
      comparePdfDocuments(approved, approved, { pdftoppmPath: path, timeoutMs: 100 }),
      /deadline/,
    );
    await writeFile(path, '#!/bin/sh\nprintf "other program\\n"\n', { mode: 0o700 });
    await assert.rejects(comparePdfDocuments(approved, approved, { pdftoppmPath: path }), /identify/);
    await assert.rejects(
      comparePdfDocuments(approved, approved, { pdftoppmPath: join(dir, 'missing') }),
      /Cannot execute/,
    );
    const candidate = Buffer.from(approved),
      options = { pdftoppmPath, dpi: 72 };
    const pending = comparePdfDocuments(approved, candidate, options);
    candidate.fill(0);
    options.dpi = 0;
    assert.equal((await pending).passed, true);
  } finally {
    await rm(dir, { recursive: true, force: true });
  }
});
test('PDF visual comparison respects crop boxes and rotations and bounds unusual page geometry', async () => {
  const { degrees, PDFName, PDFNumber } = await import('pdf-lib');
  const approved = await pdf();
  const rotated = await PDFDocument.load(approved);
  rotated.getPage(0).setRotation(degrees(90));
  const rotatedBytes = await rotated.save();
  const same = await comparePdfDocuments(rotatedBytes, rotatedBytes, { pdftoppmPath, dpi: 72 });
  assert.equal(same.passed, true);
  assert.equal(same.pages[0]!.diff!.width, 120);
  assert.equal(same.pages[0]!.diff!.height, 240);
  assert.equal((await comparePdfDocuments(approved, rotatedBytes, { pdftoppmPath, dpi: 72 })).passed, false);
  const cropped = await PDFDocument.load(approved);
  cropped.getPage(0).setCropBox(10, 10, 200, 100);
  const croppedBytes = await cropped.save();
  const cropSame = await comparePdfDocuments(croppedBytes, croppedBytes, { pdftoppmPath, dpi: 72 });
  assert.equal(cropSame.passed, true);
  assert.equal(cropSame.pages[0]!.diff!.width, 200);
  assert.equal(cropSame.pages[0]!.diff!.height, 100);
  assert.equal((await comparePdfDocuments(approved, croppedBytes, { pdftoppmPath, dpi: 72 })).passed, false);
  cropped.getPage(0).node.set(PDFName.of('UserUnit'), PDFNumber.of(10));
  await assert.rejects(comparePdfDocuments(approved, await cropped.save(), { pdftoppmPath }), /UserUnit/);
  await assert.rejects(
    comparePdfDocuments(approved, await pdf('giant', 1, 5000), { pdftoppmPath, dpi: 72 }),
    /pixel budget/,
  );
});
test('PDF regression detects embedded image changes and changes on later pages', async () => {
  const sharp = (await import('sharp')).default;
  const imagePdf = async (color: string) => {
    const doc = await PDFDocument.create();
    const image = await doc.embedPng(
      await sharp({ create: { width: 8, height: 8, channels: 3, background: color } })
        .png()
        .toBuffer(),
    );
    doc.addPage([120, 120]);
    doc.addPage([120, 120]).drawImage(image, { x: 10, y: 10, width: 80, height: 80 });
    return doc.save();
  };
  const before = await imagePdf('#ff0000'),
    after = await imagePdf('#0000ff');
  const result = await comparePdfDocuments(before, after, { pdftoppmPath, dpi: 72 });
  assert.equal(result.passed, false);
  assert.equal(result.pages.length, 2);
  assert.equal(result.pages[0]!.passed, true);
  assert.equal(result.pages[1]!.passed, false);
  assert.ok(result.pages[1]!.diff!.changedPixels >= 6400);
});
test('fractional point sizes retain exact bounded pixel dimensions across rotated rasterization', async () => {
  const { degrees } = await import('pdf-lib');
  for (const angle of [0, 90, 180, 270]) {
    const doc = await PDFDocument.create(),
      page = doc.addPage([39.09133333333333, 38.008]);
    page.setRotation(degrees(angle));
    const bytes = await doc.save();
    const result = await comparePdfDocuments(bytes, bytes, { pdftoppmPath, dpi: 96 });
    assert.equal(result.passed, true);
    assert.equal(result.pages[0]!.diff!.width, angle % 180 ? 51 : 53);
    assert.equal(result.pages[0]!.diff!.height, angle % 180 ? 53 : 51);
  }
});
