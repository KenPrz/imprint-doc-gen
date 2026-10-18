import test from 'node:test';
import assert from 'node:assert/strict';
import { mkdtemp, writeFile, rm } from 'node:fs/promises';
import { tmpdir } from 'node:os';
import { join } from 'node:path';
import { PDFDocument } from 'pdf-lib';
import sharp from 'sharp';
import { parseVeraPdfReport, veraPdfValidator } from '../src/verapdf.ts';
import { validateDelivery } from '../src/delivery.ts';
import { comparePdfDocuments } from '../src/pdf-regression.ts';
// Shape independently observed from pinned veraPDF CLI on the CC BY positive fixture.
const report = `<report><buildInformation><releaseDetails id="core" version="1.31.50"/><releaseDetails id="apps" version="1.31.202"/></buildInformation><jobs><job><item size="10231"><name>/fixtures/pdfa-1b-pass.pdf</name></item><validationReport jobEndStatus="normal" profileName="PDF/A-1b validation profile" isCompliant="true"><details passedRules="129" failedRules="0" passedChecks="553" failedChecks="0"/></validationReport></job></jobs><batchSummary totalJobs="1" failedToParse="0" encrypted="0" outOfMemory="0" veraExceptions="0"><validationReports compliant="1" nonCompliant="0" failedJobs="0">1</validationReports></batchSummary></report>`;
function refuses(xml: string) {
  try {
    return !parseVeraPdfReport(xml, '1b', '1.31.50').passed;
  } catch {
    return true;
  }
}
test('veraPDF positive report matches independently observed CLI report shape', () => {
  assert.equal(parseVeraPdfReport(report, '1b', '1.31.50').passed, true);
});
for (const [name, xml] of [
  [
    'extra job contradicts summary',
    report.replace('</jobs>', '<job><item><name>other.pdf</name></item></job></jobs>'),
  ],
  ['report outside jobs', report.replace('<jobs><job>', '').replace('</job></jobs>', '')],
  ['aborted validation job', report.replace('jobEndStatus="normal"', 'jobEndStatus="exception"')],
  ['failed rules contradict compliance', report.replace('failedRules="0"', 'failedRules="1"')],
  ['failed checks contradict compliance', report.replace('failedChecks="0"', 'failedChecks="1"')],
  ['missing validation details', report.replace(/<details[^>]*\/>/, '')],
  [
    'conflicting core versions',
    report.replace('</buildInformation>', '<releaseDetails id="core" version="9.99.99"/></buildInformation>'),
  ],
  ['foreign namespace report', report.replace('<report>', '<report xmlns="urn:untrusted-schema">')],
  [
    'wrong count content',
    report.replace('failedJobs="0">1</validationReports>', 'failedJobs="0">2</validationReports>'),
  ],
] as const)
  test(`veraPDF refuses malformed or contradictory evidence: ${name}`, () => {
    assert.equal(refuses(xml), true);
  });

test('veraPDF executable result is bound to the actual submitted artifact path', async () => {
  const dir = await mkdtemp(join(tmpdir(), 'imprint-verapdf-adversarial-'));
  try {
    const executable = join(dir, 'fake-validator');
    await writeFile(executable, `#!/bin/sh\ncat <<'VERAPDF_REPORT'\n${report}\nVERAPDF_REPORT\n`, {
      mode: 0o700,
    });
    const result = await validateDelivery(Buffer.from('%PDF-not-the-approved-fixture'), [
      veraPdfValidator({ executable, version: '1.31.50', profile: '1b' }),
    ]);
    assert.equal(result.evidence.status, 'quarantined');
    assert.equal(result.artifact, undefined);
  } finally {
    await rm(dir, { recursive: true, force: true });
  }
});

test('PDF regression refuses raster dimensions that do not match submitted page geometry', async () => {
  const dir = await mkdtemp(join(tmpdir(), 'imprint-pdf-dimensions-'));
  try {
    const document = await PDFDocument.create();
    document.addPage([240, 120]);
    const bytes = await document.save();
    const png = await sharp({ create: { width: 1, height: 1, channels: 3, background: 'white' } })
      .png()
      .toBuffer();
    const executable = join(dir, 'fake-rasterizer');
    await writeFile(
      executable,
      `#!${process.execPath}\nif(process.argv.includes('-v'))process.stdout.write('pdftoppm version fixture\\n');else process.stdout.write(Buffer.from('${png.toString('base64')}','base64'));\n`,
      { mode: 0o700 },
    );
    let passed = false;
    try {
      passed = (await comparePdfDocuments(bytes, bytes, { pdftoppmPath: executable, dpi: 72 })).passed;
    } catch {
      /* Rejection is fail-closed. */
    }
    assert.equal(passed, false);
  } finally {
    await rm(dir, { recursive: true, force: true });
  }
});
