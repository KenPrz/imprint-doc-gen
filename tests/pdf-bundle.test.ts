import test from 'node:test';
import assert from 'node:assert/strict';
import { execFileSync } from 'node:child_process';
import { mkdtemp, readFile, rm, writeFile } from 'node:fs/promises';
import { tmpdir } from 'node:os';
import { join } from 'node:path';
import { crc32 } from 'node:zlib';
import {
  PDFDocument,
  PDFHexString,
  PDFName,
  PDFString,
  StandardFonts,
  type PDFPage,
} from 'pdf-lib';
import sharp from 'sharp';
import { prepareBundleInput } from '../src/pdf-bundle.ts';
import { RenderError } from '../src/errors.ts';
import { createApp } from '../src/app.ts';

// ponytail: pdfcpu-backed tests skip locally when the binary is absent; CI installs it and never skips.
const hasPdfcpu = (() => {
  try {
    execFileSync('pdfcpu', ['version'], { stdio: 'ignore' });
    return true;
  } catch {
    return false;
  }
})();
const needsPdfcpu = { skip: !hasPdfcpu && !process.env.CI ? 'pdfcpu not installed' : false };

const code = (expected: string) => (error: unknown) =>
  error instanceof RenderError && error.code === expected;

const plainPdf = async (text = 'Rendered document') => {
  const d = await PDFDocument.create();
  d.addPage([300, 400]).drawText(text, { x: 20, y: 200 });
  return d.save();
};
const filledForm = async () => {
  const d = await PDFDocument.create();
  const page = d.addPage([400, 400]);
  const field = d.getForm().createTextField('name');
  field.setText('Ada Lovelace');
  field.addToPage(page, { x: 50, y: 200, width: 200, height: 30 });
  return d.save();
};
const signedPdf = async (withAppearance = true, rect = [50, 50, 250, 90]) => {
  const d = await PDFDocument.create();
  const page = d.addPage([400, 400]);
  const font = await d.embedFont(StandardFonts.Helvetica);
  const sig = d.context.register(
    d.context.obj({
      Type: 'Sig',
      Filter: 'Adobe.PPKLite',
      SubFilter: 'adbe.pkcs7.detached',
      Contents: PDFHexString.of('00'.repeat(64)),
      ByteRange: [0, 0, 0, 0],
    }),
  );
  const widget = d.context.obj({
    Type: 'Annot',
    Subtype: 'Widget',
    FT: 'Sig',
    T: PDFString.of('Signature1'),
    V: sig,
    Rect: rect,
    F: 4,
    P: page.ref,
  });
  if (withAppearance)
    widget.set(
      PDFName.of('AP'),
      d.context.obj({
        N: d.context.register(
          d.context.stream('BT /F1 12 Tf 10 10 Td (Signed by Example) Tj ET', {
            Type: 'XObject',
            Subtype: 'Form',
            BBox: [0, 0, 200, 40],
            Resources: { Font: { F1: font.ref } },
          }),
        ),
      }),
    );
  const ref = d.context.register(widget);
  page.node.set(PDFName.of('Annots'), d.context.obj([ref]));
  d.catalog.set(PDFName.of('AcroForm'), d.context.obj({ Fields: [ref], SigFlags: 3 }));
  d.catalog.set(PDFName.of('Perms'), d.context.obj({ DocMDP: sig }));
  return d.save();
};
const linkWithJavaScript = (d: PDFDocument, page: PDFPage) =>
  page.node.set(
    PDFName.of('Annots'),
    d.context.obj([
      d.context.register(
        d.context.obj({
          Type: 'Annot',
          Subtype: 'Link',
          Rect: [0, 0, 100, 100],
          A: { S: 'JavaScript', JS: PDFString.of('app.alert(1)') },
        }),
      ),
    ]),
  );
const activePdf = async () => {
  const d = await PDFDocument.create();
  const page = d.addPage([300, 300]);
  page.drawText('Active', { x: 20, y: 150 });
  linkWithJavaScript(d, page);
  page.node.set(PDFName.of('AA'), d.context.obj({ O: { S: 'JavaScript', JS: PDFString.of('1') } }));
  d.catalog.set(
    PDFName.of('OpenAction'),
    d.context.register(d.context.obj({ S: 'Launch', F: PDFString.of('calc.exe') })),
  );
  d.catalog.set(
    PDFName.of('Names'),
    d.context.obj({ JavaScript: { Names: [PDFString.of('x'), { S: 'JavaScript', JS: PDFString.of('2') }] } }),
  );
  return d.save();
};
const text = async (pdf: Uint8Array) => {
  const dir = await mkdtemp(join(tmpdir(), 'bundle-test-'));
  try {
    await writeFile(join(dir, 'a.pdf'), pdf);
    return execFileSync('pdftotext', [join(dir, 'a.pdf'), '-']).toString();
  } finally {
    await rm(dir, { recursive: true, force: true });
  }
};
const deepString = (doc: PDFDocument) =>
  doc.context
    .enumerateIndirectObjects()
    .map(([, o]) => o.toString())
    .join('\n');

// A 4:1 tall photo stored sideways (1600×400) with EXIF orientation 6: displayed upright, red on top.
const sidewaysPhoto = () =>
  sharp({ create: { width: 1600, height: 400, channels: 3, background: '#0000ff' } })
    .composite([
      {
        input: { create: { width: 800, height: 400, channels: 3, background: '#ff0000' } },
        left: 0,
        top: 0,
      },
    ])
    .jpeg()
    .withMetadata({ orientation: 6 })
    .toBuffer();
// A valid 1×1 PNG whose header claims 30000×30000 pixels.
const pngBomb = async () => {
  const png = Buffer.from(await sharp({ create: { width: 1, height: 1, channels: 3, background: '#000' } }).png().toBuffer());
  png.writeUInt32BE(30000, 16);
  png.writeUInt32BE(30000, 20);
  png.writeUInt32BE(crc32(png.subarray(12, 29)) >>> 0, 29);
  return png;
};

test('bundle inputs: no flags rejects what /v1/pdf rejects', async () => {
  await assert.rejects(prepareBundleInput(await filledForm()), /forms/);
  await assert.rejects(prepareBundleInput(await signedPdf()), /forms/);
  await assert.rejects(prepareBundleInput(await activePdf()), /Active/);
  await assert.rejects(prepareBundleInput(Buffer.from('not a pdf or image')), code('UNSUPPORTED_INPUT'));
  const bmp = await sharp({ create: { width: 2, height: 2, channels: 3, background: '#000' } }).raw().toBuffer();
  await assert.rejects(prepareBundleInput(Buffer.concat([Buffer.from('BM'), bmp])), code('UNSUPPORTED_INPUT'));
});

test('bundle inputs: the default scan also rejects JavaScript nested directly in an annotation', async () => {
  const d = await PDFDocument.create();
  linkWithJavaScript(d, d.addPage());
  await assert.rejects(prepareBundleInput(await d.save()), /Active/);
});

test('bundle inputs: a plain PDF passes through with its page count', async () => {
  const result = await prepareBundleInput(await plainPdf());
  assert.equal(result.pages, 1);
  assert.deepEqual(result.stripped, []);
  assert.match(await text(result.bytes), /Rendered document/);
});

test('bundle inputs: images become one page each, auto-oriented and inside the margins', async () => {
  const photo = await prepareBundleInput(await sidewaysPhoto());
  const page = (await PDFDocument.load(photo.bytes)).getPage(0);
  assert.deepEqual([Math.round(page.getWidth()), Math.round(page.getHeight())], [595, 842]);
  const dir = await mkdtemp(join(tmpdir(), 'bundle-test-'));
  try {
    await writeFile(join(dir, 'p.pdf'), photo.bytes);
    execFileSync('pdftoppm', ['-r', '20', '-png', '-singlefile', join(dir, 'p.pdf'), join(dir, 'p')]);
    const { data, info } = await sharp(join(dir, 'p.png')).raw().toBuffer({ resolveWithObject: true });
    const at = (x: number, y: number) => [...data.subarray((y * info.width + x) * info.channels, (y * info.width + x) * info.channels + 3)];
    let top = info.height, bottom = 0, left = info.width, right = 0;
    for (let y = 0; y < info.height; y++)
      for (let x = 0; x < info.width; x++)
        if (at(x, y).some((c) => c < 200)) {
          top = Math.min(top, y); bottom = Math.max(bottom, y); left = Math.min(left, x); right = Math.max(right, x);
        }
    assert(top > 0 && left > 0 && bottom < info.height - 1 && right < info.width - 1, 'image stays inside the page margins');
    assert(bottom - top > 3 * (right - left), 'tall photo is upright and fits by height');
    const cx = Math.round((left + right) / 2);
    assert(at(cx, top + 3)[0]! > 200 && at(cx, top + 3)[2]! < 80, 'red half is on top');
    assert(at(cx, bottom - 3)[2]! > 200 && at(cx, bottom - 3)[0]! < 80, 'blue half is at the bottom');
  } finally {
    await rm(dir, { recursive: true, force: true });
  }
  const webp = await sharp({ create: { width: 500, height: 300, channels: 3, background: '#00ff00' } }).webp().toBuffer();
  const gif = await sharp({ create: { width: 50, height: 50, channels: 4, background: '#00ff0080' } }).gif().toBuffer();
  const png = await sharp({ create: { width: 50, height: 50, channels: 4, background: '#00ff0080' } }).png().toBuffer();
  for (const image of [webp, gif, png]) assert.equal((await prepareBundleInput(image)).pages, 1);
  const letter = (await PDFDocument.load((await prepareBundleInput(webp, { pageSize: 'Letter' })).bytes)).getPage(0);
  assert.deepEqual([letter.getWidth(), letter.getHeight()], [612, 792]);
});

test('bundle inputs: decompression bombs are bounded and malformed EXIF is handled', async () => {
  await assert.rejects(prepareBundleInput(await pngBomb()), code('LIMIT_EXCEEDED'));
  const jpeg = Buffer.from(await sharp({ create: { width: 40, height: 40, channels: 3, background: '#888' } }).jpeg().toBuffer());
  const garbage = Buffer.concat([Buffer.from('Exif\0\0MM\0*'), Buffer.alloc(40, 0xff)]);
  const app1 = Buffer.concat([Buffer.from([0xff, 0xe1, 0, garbage.length + 2]), garbage]);
  const broken = Buffer.concat([jpeg.subarray(0, 2), app1, jpeg.subarray(2)]);
  try {
    assert.equal((await prepareBundleInput(broken)).pages, 1);
  } catch (error) {
    assert(error instanceof RenderError && error.code === 'UNSUPPORTED_INPUT', String(error));
  }
});

test('bundle inputs: flatten keeps form values and signature appearances', async () => {
  const form = await prepareBundleInput(await filledForm(), { flatten: true });
  assert.match(await text(form.bytes), /Ada Lovelace/);
  assert(!(await PDFDocument.load(form.bytes)).catalog.has(PDFName.of('AcroForm')));
  const signed = await prepareBundleInput(await signedPdf(), { flatten: true });
  assert.match(await text(signed.bytes), /Signed by Example/);
  assert.doesNotMatch(deepString(await PDFDocument.load(signed.bytes)), /\/Sig\b/);
  // An invisible signature has nothing to draw, so it flattens without an appearance.
  await prepareBundleInput(await signedPdf(false, [0, 0, 0, 0]), { flatten: true });
});

test('bundle inputs: flatten fails with typed errors for missing appearances and dynamic XFA', async () => {
  await assert.rejects(prepareBundleInput(await signedPdf(false), { flatten: true }), code('FORM_APPEARANCE_MISSING'));
  const d = await PDFDocument.create();
  d.addPage();
  d.catalog.set(PDFName.of('NeedsRendering'), d.context.obj(true));
  d.catalog.set(
    PDFName.of('AcroForm'),
    d.context.obj({ Fields: [], XFA: d.context.register(d.context.stream('<xdp:xdp/>')) }),
  );
  await assert.rejects(prepareBundleInput(await d.save(), { flatten: true }), code('XFA_UNSUPPORTED'));
});

test('bundle inputs: stripActiveContent removes actions and reports them', async () => {
  const result = await prepareBundleInput(await activePdf(), { stripActiveContent: true });
  const out = await PDFDocument.load(result.bytes);
  assert.doesNotMatch(deepString(out), /JavaScript|\/JS\b|Launch|\/AA\b|OpenAction/);
  assert.match(await text(result.bytes), /Active/);
  for (const key of ['A', 'AA', 'OpenAction', 'JavaScript']) assert(result.stripped.includes(key), key);
  const dest = await PDFDocument.create();
  const page = dest.addPage();
  dest.catalog.set(PDFName.of('OpenAction'), dest.context.obj([page.ref, PDFName.of('Fit')]));
  assert.equal((await prepareBundleInput(await dest.save(), { stripActiveContent: true })).pages, 1);
  // Stripping is not flattening: forms are still rejected.
  await assert.rejects(prepareBundleInput(await filledForm(), { stripActiveContent: true }), /forms/);
});

const multipart = async (parts: [string, Blob | string][]) => {
  const form = new FormData();
  for (const [name, value] of parts) form.append(name, value);
  const request = new Request('http://localhost', { method: 'POST', body: form });
  return { payload: Buffer.from(await request.arrayBuffer()), contentType: request.headers.get('content-type')! };
};
const blob = (bytes: Uint8Array) => new Blob([Buffer.from(bytes)]);

test('bundle route assembles the acceptance bundle', needsPdfcpu, async () => {
  const apiKey = 'a'.repeat(32);
  const dir = await mkdtemp(join(tmpdir(), 'bundle-test-'));
  const app = createApp({ apiKey });
  try {
    await writeFile(join(dir, 'plain.pdf'), await plainPdf('Owner protected'));
    execFileSync('pdfcpu', ['encrypt', '--opw', 'secret', join(dir, 'plain.pdf'), join(dir, 'owner.pdf')], { stdio: 'ignore' });
    execFileSync('pdfcpu', ['encrypt', '--upw', 'u', '--opw', 'o', join(dir, 'plain.pdf'), join(dir, 'user.pdf')], { stdio: 'ignore' });
    const owner = await readFile(join(dir, 'owner.pdf'));
    const webp = await sharp({ create: { width: 500, height: 300, channels: 3, background: '#00ff00' } }).webp().toBuffer();
    const body = await multipart([
      ['decrypt', 'true'],
      ['flatten', 'true'],
      ['stripActiveContent', 'true'],
      ['files', blob(await plainPdf())],
      ['files', blob(owner)],
      ['files', blob(await filledForm())],
      ['files', blob(await signedPdf())],
      ['files', blob(await sidewaysPhoto())],
      ['files', blob(webp)],
      ['files', blob(await activePdf())],
    ]);
    const headers = { authorization: `Bearer ${apiKey}`, 'content-type': body.contentType };
    const response = await app.inject({ method: 'POST', url: '/v1/pdf/bundle', headers, payload: body.payload });
    assert.equal(response.statusCode, 200, response.body);
    assert.equal(response.headers['content-type'], 'application/pdf');
    const pdf = await PDFDocument.load(response.rawPayload);
    assert.equal(pdf.getPageCount(), 7);
    const extracted = await text(response.rawPayload);
    for (const value of ['Rendered document', 'Owner protected', 'Ada Lovelace', 'Signed by Example', 'Active'])
      assert.match(extracted, new RegExp(value));
    assert.match(String(response.headers['imprint-stripped-active-content']), /^7:/);

    const send = async (parts: [string, Blob | string][], auth = true) => {
      const b = await multipart(parts);
      return app.inject({
        method: 'POST',
        url: '/v1/pdf/bundle',
        headers: { ...(auth ? { authorization: `Bearer ${apiKey}` } : {}), 'content-type': b.contentType },
        payload: b.payload,
      });
    };
    assert.equal((await send([['files', blob(await plainPdf())]], false)).statusCode, 401);
    const password = await send([['decrypt', 'true'], ['files', blob(await readFile(join(dir, 'user.pdf')))]]);
    assert.equal(password.statusCode, 422);
    assert.equal(password.json().error, 'PDF_PASSWORD_REQUIRED');
    const noDecrypt = await send([['files', blob(owner)]]);
    assert.equal(noDecrypt.statusCode, 422);
    assert.equal(noDecrypt.json().error, 'INVALID_PDF_OPERATION');
    assert.equal((await send([['files', blob(await filledForm())]])).statusCode, 422);
    assert.equal((await send([['colour', 'red'], ['files', blob(await plainPdf())]])).statusCode, 400);
    assert.equal((await send([['decrypt', 'yes'], ['files', blob(await plainPdf())]])).statusCode, 400);
    assert.equal((await send([['decrypt', 'true']])).statusCode, 400);
    // The cap is on total bytes, not count: a 22-part bundle of small files assembles.
    const small = blob(await plainPdf());
    const many = await send(Array.from({ length: 22 }, () => ['files', small] as [string, Blob]));
    assert.equal(many.statusCode, 200, many.body);
    assert.equal((await PDFDocument.load(many.rawPayload)).getPageCount(), 22);
    const tooMany = await send(Array.from({ length: 51 }, () => ['files', small] as [string, Blob]));
    assert.equal(tooMany.statusCode, 413);
    const tenMiB = blob(Buffer.concat([Buffer.from('%PDF-'), Buffer.alloc(10 * 1024 * 1024 - 5)]));
    const overTotal = await send(Array.from({ length: 21 }, () => ['files', tenMiB] as [string, Blob]));
    assert.equal(overTotal.statusCode, 413);
    assert.equal(overTotal.json().error, 'LIMIT_EXCEEDED');
    // The app-wide 16 MiB JSON bodyLimit must not cap multipart bundles: 18 MiB of junk PDFs reaches the scan.
    const junk = blob(Buffer.concat([Buffer.from('%PDF-'), Buffer.alloc(9 * 1024 * 1024)]));
    assert.equal((await send([['files', junk], ['files', junk]])).statusCode, 422);
    const tooBig = await send([['files', blob(Buffer.alloc(10 * 1024 * 1024 + 1))]]);
    assert.equal(tooBig.statusCode, 413);
    assert.equal(
      (await app.inject({ method: 'POST', url: '/v1/pdf/bundle', headers: { authorization: `Bearer ${apiKey}` }, payload: { files: [] } })).statusCode,
      415,
    );
  } finally {
    await app.close();
    await rm(dir, { recursive: true, force: true });
  }
});
