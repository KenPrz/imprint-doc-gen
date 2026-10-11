import test from 'node:test';
import assert from 'node:assert/strict';
import { PDFDocument, PDFName } from 'pdf-lib';
import sharp from 'sharp';
import { processPdf, splitPdf } from '../src/pdf.ts';
import { barcode } from '../src/barcode.ts';
import { createApp } from '../src/app.ts';
const source = async () => {
  const d = await PDFDocument.create();
  d.addPage([200, 300]).drawText('First');
  d.addPage([400, 500]).drawText('Second');
  return d.save();
};
test('PDF merges selected pages in order, numbers, watermarks, and embeds attachments', async () => {
  const bytes = await source(),
    result = await processPdf({
      inputs: [
        { bytes, pages: [2] },
        { bytes, pages: [1] },
      ],
      watermark: { text: 'DRAFT' },
      pageNumbers: true,
      attachments: [{ name: 'details.txt', bytes: Buffer.from('Example') }],
    }),
    pdf = await PDFDocument.load(result);
  assert.equal(pdf.getPageCount(), 2);
  assert.equal(pdf.getPage(0).getWidth(), 400);
  assert.equal(pdf.getPage(1).getWidth(), 200);
  assert(pdf.catalog.has(PDFName.of('Names')));
  const split = await splitPdf(bytes, [[2], [1]]);
  assert.equal((await PDFDocument.load(split[0]!)).getPage(0).getWidth(), 400);
  await assert.rejects(processPdf({ inputs: [{ bytes, pages: [0] }] }), /one-based/);
  await assert.rejects(
    processPdf({ inputs: [{ bytes }], attachments: [{ name: '../bad', bytes }] }),
    /filenames/,
  );
  const form = await PDFDocument.create();
  form.addPage();
  form.getForm().createTextField('name');
  await assert.rejects(processPdf({ inputs: [{ bytes: await form.save() }] }), /forms/);
});
test('barcode outputs bounded PNG attachments for four symbologies', async () => {
  for (const type of ['qrcode', 'code128', 'ean13', 'datamatrix'] as const) {
    const png = await barcode({ type, text: type === 'ean13' ? '5901234123457' : 'Example 123' }),
      info = await sharp(png).metadata();
    assert.equal(info.format, 'png');
    assert(info.width! > 20);
  }
  await assert.rejects(barcode({ type: 'qrcode', text: 'x'.repeat(513) }));
  await assert.rejects(barcode({ type: 'ean13', text: 'bad' }));
});
test('PDF and barcode routes use authenticated isolated workers', async () => {
  const apiKey = 'a'.repeat(32),
    app = createApp({ apiKey }),
    headers = { authorization: `Bearer ${apiKey}` };
  try {
    assert.equal(
      (await app.inject({ method: 'POST', url: '/v1/barcode', payload: { type: 'qrcode', text: 'test' } }))
        .statusCode,
      401,
    );
    const response = await app.inject({
      method: 'POST',
      url: '/v1/pdf',
      headers,
      payload: {
        inputs: [{ bytes: Buffer.from(await source()).toString('base64'), pages: [2] }],
        pageNumbers: true,
      },
    });
    assert.equal(response.statusCode, 200, response.body);
    assert.equal((await PDFDocument.load(response.rawPayload)).getPageCount(), 1);
    const split = await app.inject({
      method: 'POST',
      url: '/v1/pdf/split',
      headers,
      payload: { bytes: Buffer.from(await source()).toString('base64'), ranges: [[2], [1]] },
    });
    assert.equal(split.statusCode, 200);
    assert.equal(split.json().documents.length, 2);
    const png = await app.inject({
      method: 'POST',
      url: '/v1/barcode',
      headers,
      payload: { type: 'qrcode', text: 'test' },
    });
    assert.equal(png.statusCode, 200);
    assert.equal((await sharp(png.rawPayload).metadata()).format, 'png');
  } finally {
    await app.close();
  }
});
