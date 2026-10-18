import { test } from 'node:test';
import assert from 'node:assert/strict';
import { parseVeraPdfReport, veraPdfValidator } from '../src/verapdf.ts';
const report = `<report><buildInformation><releaseDetails id="core" version="1.30.2"/></buildInformation><jobs><job><item size="100"><name>/test.pdf</name></item><validationReport jobEndStatus="normal" profileName="PDF/A-1b validation profile" isCompliant="true"><details passedRules="100" failedRules="0" passedChecks="500" failedChecks="0"/></validationReport></job></jobs><batchSummary totalJobs="1" failedToParse="0" encrypted="0" outOfMemory="0" veraExceptions="0"><validationReports compliant="1" nonCompliant="0" failedJobs="0">1</validationReports></batchSummary></report>`;
test('veraPDF evidence requires exact profile, engine, one job and zero failures', () => {
  assert.equal(parseVeraPdfReport(report, '1b', '1.30.2').passed, true);
  for (const [from, to] of [
    ['isCompliant="true"', 'isCompliant="false"'],
    ['failedToParse="0"', 'failedToParse="1"'],
    ['compliant="1"', 'compliant="0"'],
    ['totalJobs="1"', 'totalJobs="2"'],
    ['failedJobs="0"', 'failedJobs="1"'],
  ])
    assert.equal(parseVeraPdfReport(report.replace(from, to), '1b', '1.30.2').passed, false);
  assert.throws(() => parseVeraPdfReport(report, 'ua1', '1.30.2'));
  assert.throws(() => parseVeraPdfReport(report, '1b', '1.30.1'));
  for (const xml of [
    '',
    report.replace('</jobs>', '<validationReport/></jobs>'),
    report.replace('<report>', '<!DOCTYPE report><report>'),
    report.replace('version="1.30.2"', 'version="fake"'),
  ])
    assert.throws(() => parseVeraPdfReport(xml, '1b', '1.30.2'));
});
test('veraPDF does not accept arbitrary executable/profile configuration', () => {
  assert.throws(() => veraPdfValidator({ executable: 'verapdf', profile: '1b', version: '1.30.2' }));
  assert.throws(() =>
    veraPdfValidator({ executable: '/usr/bin/verapdf', profile: 'auto' as any, version: '1.30.2' }),
  );
  assert.throws(() =>
    veraPdfValidator({ executable: '/usr/bin/verapdf', profile: '1b', version: '1.30.2', timeoutMs: 0 }),
  );
});
