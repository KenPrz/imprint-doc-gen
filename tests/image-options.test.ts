import test from 'node:test';
import assert from 'node:assert/strict';
import sharp from 'sharp';
import { readFile } from 'node:fs/promises';
import { normalizeImage } from '../src/render-package.ts';
import { render } from '../src/index.ts';
import { readPackage } from '../src/package.ts';

test('Image fit preserves aspect ratio with transparent contain and filled cover', async () => {
  const input = await sharp({ create: { width: 80, height: 40, channels: 3, background: '#ff0000' } })
    .png()
    .toBuffer();
  for (const fit of ['contain', 'cover'] as const) {
    const { bytes } = await normalizeImage(input, { fit }, { width: 40, height: 40 });
    const { data, info } = await sharp(bytes).ensureAlpha().raw().toBuffer({ resolveWithObject: true });
    assert.equal(info.width, 40);
    assert.equal(info.height, 40);
    assert.equal(data[3], fit === 'contain' ? 0 : 255);
    assert.equal(data[(20 * 40 + 20) * 4], 255);
  }
});
test('Crop, rotation and EXIF orientation produce bounded normalized PNGs', async () => {
  const input = await sharp({ create: { width: 80, height: 40, channels: 3, background: '#224466' } })
    .jpeg()
    .withMetadata({ orientation: 6 })
    .toBuffer();
  const oriented = await normalizeImage(input),
    meta = await sharp(oriented.bytes).metadata();
  assert.equal(meta.width, 40);
  assert.equal(meta.height, 80);
  const cropped = await normalizeImage(input, {
    crop: { left: 0, top: 0, width: 20, height: 30 },
    rotate: 90,
  });
  assert.equal((await sharp(cropped.bytes).metadata()).width, 30);
  await assert.rejects(normalizeImage(input, { fit: 'contain', width: 5000, height: 10 }), /4096/);
  await assert.rejects(normalizeImage(input, { rotate: Infinity }), /Rotation/);
});
test('SDK image options use the existing picture extent as the fitting canvas', async () => {
  const input = await readFile('examples/invoice/invoice.docx');
  const logo = await sharp({ create: { width: 40, height: 80, channels: 3, background: '#ff0000' } })
    .png()
    .toBuffer();
  const data = JSON.parse(await readFile('examples/invoice/data.json', 'utf8'));
  const output = await readPackage(
    await render(input, data, { attachments: { logo }, images: { logo: { fit: 'contain' } } }),
  );
  const image = [...output].find(([name]) => name.includes('/media/imprint-'))![1];
  const meta = await sharp(image).metadata();
  assert.equal(meta.width, 192);
  assert.equal(meta.height, 64);
});

test('Spreadsheet pictures repeat with item-specific media and unique identities', async () => {
  const { writePackage } = await import('../src/package.ts');
  const { parseXml, descendants, ns } = await import('../src/xml.ts');
  const parts = await readPackage(await readFile('examples/invoice/invoice.xlsx'));
  for (const [part, bytes] of parts)
    if (/xl\/drawings\/[^/]+\.xml$/.test(part))
      parts.set(
        part,
        Buffer.from(
          bytes
            .toString()
            .replaceAll('{d.logo}', '{d.items[i].image}')
            .replaceAll('<xdr:row>0</xdr:row>', '<xdr:row>10</xdr:row>'),
        ),
      );
  const input = await writePackage(parts),
    red = await sharp({ create: { width: 8, height: 8, channels: 3, background: '#ff0000' } })
      .png()
      .toBuffer(),
    blue = await sharp({ create: { width: 8, height: 8, channels: 3, background: '#0000ff' } })
      .png()
      .toBuffer();
  const result = await readPackage(
    await render(
      input,
      {
        items: [
          { description: 'one', quantity: 1, price: 1, image: 'red' },
          { description: 'two', quantity: 1, price: 1, image: 'blue' },
        ],
      },
      { attachments: { red, blue } },
    ),
  );
  const drawing = parseXml(
    [...result].find(([p]) => /xl\/drawings\/[^/]+\.xml$/.test(p))![1].toString(),
    'drawing',
  );
  assert.equal(descendants(drawing, ns.spreadsheetDrawing, 'pic').length, 2);
  assert.equal(
    new Set(descendants(drawing, ns.spreadsheetDrawing, 'cNvPr').map((n) => n.getAttribute('id'))).size,
    2,
  );
  assert.deepEqual(
    descendants(drawing, ns.spreadsheetDrawing, 'row').map((n) => n.textContent),
    ['10', '11'],
  );
  assert([...result.values()].some((b) => b.equals(red)));
  assert([...result.values()].some((b) => b.equals(blue)));
  const empty = await readPackage(await render(input, { items: [] }));
  assert(
    ![...empty]
      .find(([p]) => /xl\/drawings\/[^/]+\.xml$/.test(p))![1]
      .toString()
      .includes('<xdr:pic>'),
  );
});
