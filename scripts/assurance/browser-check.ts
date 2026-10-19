import assert from 'node:assert/strict';
import { mkdir, readFile, writeFile } from 'node:fs/promises';
import { resolve } from 'node:path';
import { pathToFileURL } from 'node:url';
import { chromium } from 'playwright';
import { createReportSnapshot, reportSnapshotHtml, renderReportSnapshot } from '../../src/report-view.ts';
import { encodeTextTemplate } from '../../src/text-template.ts';
const output = resolve('artifacts/assurance');
await mkdir(output, { recursive: true });
const malicious = '</script><img src=https://example.invalid/x onerror=alert(1)>';
const definition = {
  title: 'Sales & customer detail',
  columns: [
    { key: 'name', label: 'Customer', type: 'string' as const },
    { key: 'amount', label: 'Amount', type: 'number' as const, sum: true },
  ],
  rows: [
    { name: 'Alice', amount: 12.5 },
    { name: 'Bob, Inc.', amount: 25 },
    { name: malicious, amount: 0 },
    { name: '=HYPERLINK("https://evil.invalid")', amount: -2 },
  ],
};
await writeFile(resolve(output, 'report.html'), reportSnapshotHtml(createReportSnapshot(definition)));
const browser = await chromium.launch({ headless: true });
try {
  const page = await browser.newPage({ viewport: { width: 1440, height: 900 }, acceptDownloads: true });
  const errors: string[] = [],
    requests: string[] = [];
  page.on('pageerror', (e) => errors.push(e.message));
  page.on('request', (r) => {
    if (r.url().startsWith('http')) requests.push(r.url());
  });
  page.on('dialog', async (d) => {
    errors.push('Unexpected dialog: ' + d.message());
    await d.dismiss();
  });
  await page.goto(pathToFileURL(resolve(output, 'report.html')).href);
  assert.equal(await page.locator('tbody tr').count(), 4);
  await page.getByRole('button', { name: 'Details for row 3', exact: true }).click();
  assert.ok((await page.locator('#detail').textContent())!.includes(malicious));
  await page.screenshot({ path: resolve(output, 'report-desktop.png'), fullPage: true });
  await page.getByRole('searchbox').fill('Bob');
  assert.equal(await page.locator('tbody tr').count(), 1);
  assert.equal(await page.locator('#summary').textContent(), '1 of 4 rows');
  assert.equal(await page.locator('#totals').textContent(), 'Amount: 25');
  const csvDownload = page.waitForEvent('download');
  await page.getByRole('button', { name: 'Download selected CSV', exact: true }).click();
  const csv = await csvDownload;
  await csv.saveAs(resolve(output, 'selected.csv'));
  assert.equal(
    await readFile(resolve(output, 'selected.csv'), 'utf8'),
    '"Customer","Amount"\r\n"Bob, Inc.","25"',
  );
  const jsonDownload = page.waitForEvent('download');
  await page.getByRole('button', { name: 'Download selected data', exact: true }).click();
  await (await jsonDownload).saveAs(resolve(output, 'selection.json'));
  const selected = JSON.parse(await readFile(resolve(output, 'selection.json'), 'utf8'));
  const snap = createReportSnapshot(selected);
  assert.equal(snap.rows.length, 1);
  assert.equal(snap.totals.amount, 25);
  const rendered = await renderReportSnapshot(
    encodeTextTemplate('txt', '{d.report.rows[i].name}: {d.report.rows[i].amount}\n{d.report.rows[i+1]}'),
    snap,
  );
  assert.equal(rendered.toString(), 'Bob, Inc.: 25');
  await page.getByRole('searchbox').fill('no matches');
  assert.equal(await page.locator('tbody tr').count(), 0);
  await page.getByRole('searchbox').fill('');
  await page.setViewportSize({ width: 390, height: 844 });
  await page.screenshot({ path: resolve(output, 'report-mobile.png'), fullPage: true });
  assert.ok(await page.evaluate(() => document.documentElement.scrollWidth <= window.innerWidth));
  const allCsv = page.waitForEvent('download');
  await page.getByRole('button', { name: 'Download selected CSV', exact: true }).click();
  await (await allCsv).saveAs(resolve(output, 'all.csv'));
  assert.match(await readFile(resolve(output, 'all.csv'), 'utf8'), /'=HYPERLINK/);
  await page.pdf({ path: resolve(output, 'report-browser.pdf'), printBackground: true });
  assert.deepEqual(errors, []);
  assert.deepEqual(requests, []);
  await writeFile(
    resolve(output, 'browser-validation.json'),
    JSON.stringify(
      {
        passed: true,
        engine: browser.version(),
        checks: [
          'desktop',
          'mobile',
          'escaping',
          'CSP',
          'no-external-requests',
          'filter',
          'details',
          'CSV-selection',
          'CSV-injection',
          'JSON-selection',
          'SDK-same-selection',
          'empty-selection',
          'print-PDF',
        ],
      },
      null,
      2,
    ),
  );
  console.log(
    'Browser checks passed: filter, drill-down, selected exports, hostile content, desktop/mobile and PDF print',
  );
} finally {
  await browser.close();
}
