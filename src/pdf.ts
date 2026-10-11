import { PDFDocument, PDFName, PDFDict, StandardFonts, degrees, rgb } from 'pdf-lib';
import { RenderError } from './errors.ts';
export interface PdfInput {
  bytes: Uint8Array;
  pages?: number[];
}
export interface PdfOptions {
  inputs: PdfInput[];
  watermark?: { text: string; opacity?: number; rotation?: number };
  pageNumbers?: boolean;
  attachments?: { name: string; bytes: Uint8Array; mimeType?: string }[];
}
const limit = 32 * 1024 * 1024;
const invalid = (message: string): never => {
  throw new RenderError('INVALID_PDF_OPERATION', message);
};
// Parsing and processing are also isolated by the HTTP worker deadline and heap budget.
export async function processPdf(options: PdfOptions): Promise<Buffer> {
  if (!Array.isArray(options.inputs) || options.inputs.length < 1 || options.inputs.length > 20)
    invalid('Supply 1–20 PDF inputs');
  let total = 0,
    pages = 0;
  const output = await PDFDocument.create();
  for (const input of options.inputs) {
    if (
      !(input.bytes instanceof Uint8Array) ||
      !Buffer.from(input.bytes.subarray(0, 5)).equals(Buffer.from('%PDF-'))
    )
      invalid('Input must be a PDF');
    total += input.bytes.length;
    if (total > limit) throw new RenderError('LIMIT_EXCEEDED', 'PDF input budget exceeded');
    let source: PDFDocument;
    try {
      source = await PDFDocument.load(input.bytes, { throwOnInvalidObject: true, updateMetadata: false });
    } catch {
      invalid('Cannot read PDF; encrypted or malformed inputs are unsupported');
    }
    if (source!.getPageCount() > 1000)
      throw new RenderError('LIMIT_EXCEEDED', 'Source PDF exceeds 1000 pages');
    // Copying pages cannot preserve signatures or interactive forms faithfully. Fail explicitly.
    if (source!.catalog.has(PDFName.of('AcroForm')))
      invalid('Interactive forms and signed PDFs are unsupported');
    for (const [, object] of source!.context.enumerateIndirectObjects())
      if (object instanceof PDFDict) {
        for (const key of ['JavaScript', 'JS', 'Launch', 'OpenAction', 'AA', 'XFA'])
          if (object.has(PDFName.of(key))) invalid('Active PDF content is unsupported');
        if (
          ['JavaScript', 'Launch', 'SubmitForm', 'ImportData', 'GoToR', 'Rendition'].some(
            (action) => object.get(PDFName.of('S')) === PDFName.of(action),
          ) ||
          object.get(PDFName.of('Type')) === PDFName.of('Sig')
        )
          invalid('Active or signed PDF content is unsupported');
      }
    const selected = input.pages ?? Array.from({ length: source!.getPageCount() }, (_, i) => i + 1);
    if (
      !selected.length ||
      selected.length > 1000 ||
      selected.some((n) => !Number.isSafeInteger(n) || n < 1 || n > source!.getPageCount())
    )
      invalid('Pages use one-based indexes within each input');
    pages += selected.length;
    if (pages > 1000) throw new RenderError('LIMIT_EXCEEDED', 'Output exceeds 1000 pages');
    for (const page of await output.copyPages(
      source!,
      selected.map((n) => n - 1),
    ))
      output.addPage(page);
  }
  const watermark = options.watermark;
  if (
    watermark &&
    (!watermark.text ||
      watermark.text.length > 200 ||
      !Number.isFinite(watermark.opacity ?? 0.2) ||
      (watermark.opacity ?? 0.2) < 0 ||
      (watermark.opacity ?? 0.2) > 1 ||
      !Number.isFinite(watermark.rotation ?? 30) ||
      Math.abs(watermark.rotation ?? 30) > 360)
  )
    invalid('Invalid watermark');
  const font = await output.embedFont(StandardFonts.Helvetica);
  if (watermark)
    try {
      font.encodeText(watermark.text);
    } catch {
      invalid('Watermark text must be supported by the built-in Latin font');
    }
  for (const [index, page] of output.getPages().entries()) {
    if (watermark) {
      const size = Math.max(
        8,
        Math.min(48, (page.getWidth() * 0.7) / Math.max(1, font.widthOfTextAtSize(watermark.text, 1))),
      );
      page.drawText(watermark.text, {
        x: page.getWidth() * 0.15,
        y: page.getHeight() / 2,
        size,
        font,
        opacity: watermark.opacity ?? 0.2,
        rotate: degrees(watermark.rotation ?? 30),
        color: rgb(0.4, 0.4, 0.4),
      });
    }
    if (options.pageNumbers) {
      const text = `${index + 1} / ${pages}`;
      page.drawText(text, {
        x: Math.max(0, (page.getWidth() - font.widthOfTextAtSize(text, 10)) / 2),
        y: 16,
        size: 10,
        font,
      });
    }
  }
  if (options.attachments && (options.attachments.length > 20 || !Array.isArray(options.attachments)))
    invalid('At most 20 attachments');
  const names = new Set<string>();
  for (const item of options.attachments ?? []) {
    if (
      !/^[\p{L}\p{N} _.-]{1,100}$/u.test(item.name) ||
      item.name === '.' ||
      item.name === '..' ||
      names.has(item.name) ||
      !(item.bytes instanceof Uint8Array)
    )
      invalid('Attachment names must be unique simple filenames');
    names.add(item.name);
    total += item.bytes.length;
    if (total > limit) throw new RenderError('LIMIT_EXCEEDED', 'PDF attachment budget exceeded');
    if (item.mimeType && !/^[\w.+-]+\/[\w.+-]+$/.test(item.mimeType))
      invalid('Invalid attachment media type');
    await output.attach(item.bytes, item.name, { mimeType: item.mimeType });
  }
  const bytes = Buffer.from(await output.save());
  if (bytes.length > limit) throw new RenderError('LIMIT_EXCEEDED', 'PDF output budget exceeded');
  return bytes;
}
export async function splitPdf(bytes: Uint8Array, ranges: number[][]): Promise<Buffer[]> {
  if (!ranges.length || ranges.length > 20 || ranges.reduce((n, r) => n + r.length, 0) > 1000)
    invalid('Supply at most 20 groups and 1000 selected pages');
  const output: Buffer[] = [];
  let total = 0;
  for (const pages of ranges) {
    const part = await processPdf({ inputs: [{ bytes, pages }] });
    total += part.length;
    if (total > limit) throw new RenderError('LIMIT_EXCEEDED', 'Split output budget exceeded');
    output.push(part);
  }
  return output;
}
