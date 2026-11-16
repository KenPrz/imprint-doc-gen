import { readFile, writeFile } from 'node:fs/promises';
import { PDFArray, PDFBool, PDFDict, PDFDocument, PDFName, PDFRef, PDFStream, type PDFObject } from 'pdf-lib';
import sharp, { type OutputInfo } from 'sharp';
import { RenderError } from './errors.ts';
import { limits } from './package.ts';

export interface BundleInputOptions {
  flatten?: boolean;
  stripActiveContent?: boolean;
  pageSize?: 'A4' | 'Letter';
}
export interface BundleInput {
  bytes: Uint8Array;
  pages: number;
  stripped: string[];
}

const name = (key: string) => PDFName.of(key);
const sizes = { A4: [595.28, 841.89], Letter: [612, 792] } as const;
const margin = 36;
// Same keys and actions as the /v1/pdf scan (src/pdf.ts), applied to nested direct objects as well.
const activeKeys = ['JavaScript', 'JS', 'Launch', 'OpenAction', 'AA', 'XFA'].map(name);
const activeActions = ['JavaScript', 'Launch', 'SubmitForm', 'ImportData', 'GoToR', 'Rendition'].map(name);
const invalid = (message: string): never => {
  throw new RenderError('INVALID_PDF_OPERATION', message);
};

/** Worker entry: reads one staged upload, writes its normalised single-input PDF. */
export async function prepareBundleFile(payload: { path: string; out: string } & BundleInputOptions) {
  const { bytes, pages, stripped } = await prepareBundleInput(await readFile(payload.path), payload);
  await writeFile(payload.out, bytes);
  return { pages, stripped };
}

export async function prepareBundleInput(bytes: Uint8Array, options: BundleInputOptions = {}): Promise<BundleInput> {
  if (Buffer.from(bytes.subarray(0, 5)).toString('latin1') !== '%PDF-')
    return { bytes: await imagePage(bytes, sizes[options.pageSize ?? 'A4']), pages: 1, stripped: [] };
  let source: PDFDocument;
  try {
    source = await PDFDocument.load(bytes, { throwOnInvalidObject: true, updateMetadata: false });
  } catch {
    return invalid('Cannot read PDF; encrypted or malformed inputs are unsupported');
  }
  if (source.getPageCount() > 1000) throw new RenderError('LIMIT_EXCEEDED', 'Source PDF exceeds 1000 pages');
  const stripped = new Set<string>();
  if (options.flatten) flatten(source);
  if (options.stripActiveContent) for (const [, object] of source.context.enumerateIndirectObjects()) strip(object, source, stripped);
  // Transforms leave orphaned objects (signature dictionaries, removed actions). Without transforms, orphans
  // are scanned exactly as /v1/pdf scans them.
  if (options.flatten || options.stripActiveContent) prune(source);
  scan(source);
  const output = await PDFDocument.create();
  for (const page of await output.copyPages(source, source.getPageIndices())) output.addPage(page);
  return { bytes: await output.save(), pages: output.getPageCount(), stripped: [...stripped] };
}

function children(object: PDFObject): PDFObject[] {
  if (object instanceof PDFStream) return [object.dict];
  if (object instanceof PDFDict) return object.values();
  if (object instanceof PDFArray) return object.asArray();
  return [];
}
function walk(object: PDFObject, visit: (dict: PDFDict) => void) {
  if (object instanceof PDFDict) visit(object);
  for (const child of children(object)) if (!(child instanceof PDFRef)) walk(child, visit);
}

function scan(doc: PDFDocument) {
  if (doc.catalog.has(name('AcroForm'))) invalid('Interactive forms and signed PDFs are unsupported');
  for (const [, object] of doc.context.enumerateIndirectObjects())
    walk(object, (dict) => {
      if (activeKeys.some((key) => dict.has(key))) invalid('Active PDF content is unsupported');
      if (activeActions.includes(dict.get(name('S')) as PDFName) || dict.get(name('Type')) === name('Sig'))
        invalid('Active or signed PDF content is unsupported');
    });
}

function strip(object: PDFObject, doc: PDFDocument, stripped: Set<string>) {
  const dangerous = (value: PDFObject | undefined) => {
    const target = value instanceof PDFRef ? doc.context.lookup(value) : value;
    return target instanceof PDFDict && activeActions.includes(target.get(name('S')) as PDFName);
  };
  walk(object, (dict) => {
    for (const [key, value] of dict.entries()) {
      if (activeKeys.includes(key) || dangerous(value)) {
        dict.delete(key);
        stripped.add(key.decodeText());
      } else if (value instanceof PDFArray)
        for (let i = value.size() - 1; i >= 0; i--)
          if (dangerous(value.get(i))) {
            value.remove(i);
            stripped.add(key.decodeText());
          }
    }
  });
}

function prune(doc: PDFDocument) {
  const reachable = new Set<PDFRef>();
  const visit = (object: PDFObject | undefined) => {
    if (object instanceof PDFRef) {
      if (reachable.has(object)) return;
      reachable.add(object);
      object = doc.context.lookup(object);
    }
    if (object) for (const child of children(object)) visit(child);
  };
  visit(doc.context.trailerInfo.Root);
  visit(doc.context.trailerInfo.Info);
  for (const [ref] of doc.context.enumerateIndirectObjects()) if (!reachable.has(ref)) doc.context.delete(ref);
}

function flatten(doc: PDFDocument) {
  const acroForm = doc.catalog.lookupMaybe(name('AcroForm'), PDFDict);
  if (!acroForm) return;
  if (
    acroForm.has(name('XFA')) &&
    (doc.catalog.get(name('NeedsRendering')) === PDFBool.True || !acroForm.lookupMaybe(name('Fields'), PDFArray)?.size())
  )
    throw new RenderError('XFA_UNSUPPORTED', 'Dynamic XFA forms cannot be flattened');
  const form = doc.getForm();
  for (const field of form.getFields())
    for (const widget of field.acroField.getWidgets()) {
      if (widget.dict.lookupMaybe(name('AP'), PDFDict)?.has(name('N'))) continue;
      const { width, height } = widget.getRectangle();
      // An invisible widget (typically an invisible signature) has nothing to draw; give it an empty appearance.
      if (width === 0 || height === 0)
        widget.setNormalAppearance(doc.context.register(doc.context.formXObject([], { BBox: [0, 0, 0, 0] })));
      else
        throw new RenderError('FORM_APPEARANCE_MISSING', `Form field "${field.getName()}" has no appearance to flatten`);
    }
  try {
    form.flatten({ updateFieldAppearances: false });
  } catch {
    invalid('Form fields could not be flattened');
  }
  // pdf-lib 1.17.1 removeField() deletes kid widgets but leaves their refs in /Annots (it removes the
  // appearance ref instead). Drop the dangling refs so the copied pages don't point at missing objects.
  for (const page of doc.getPages()) {
    const annots = page.node.Annots();
    for (let i = (annots?.size() ?? 0) - 1; i >= 0; i--) {
      const annot = annots!.get(i);
      if (annot instanceof PDFRef && !doc.context.lookup(annot)) annots!.remove(i);
    }
  }
  doc.catalog.delete(name('AcroForm'));
  doc.catalog.delete(name('Perms'));
}

async function imagePage(bytes: Uint8Array, [pageWidth, pageHeight]: readonly [number, number]) {
  const options = { limitInputPixels: limits.pixels, failOn: 'warning' as const };
  const [boxWidth, boxHeight] = [pageWidth - 2 * margin, pageHeight - 2 * margin];
  let image: { data: Buffer; info: OutputInfo }, alpha: boolean;
  try {
    const metadata = await sharp(bytes, options).metadata();
    if (!['jpeg', 'png', 'webp', 'gif'].includes(metadata.format ?? ''))
      throw new RenderError('UNSUPPORTED_INPUT', 'Inputs must be PDF, JPEG, PNG, WebP or GIF');
    if ((metadata.width ?? 0) * (metadata.pageHeight ?? metadata.height ?? 0) > limits.pixels)
      throw new RenderError('LIMIT_EXCEEDED', 'Image exceeds the pixel limit');
    alpha = metadata.hasAlpha ?? false;
    // Auto-orient from EXIF, then cap at 300 dpi of the printable area so photos don't bloat the bundle.
    const pipeline = sharp(bytes, options)
      .rotate()
      .resize({ width: Math.round((boxWidth / 72) * 300), height: Math.round((boxHeight / 72) * 300), fit: 'inside', withoutEnlargement: true });
    image = await (alpha ? pipeline.png() : pipeline.jpeg({ quality: 90 })).toBuffer({ resolveWithObject: true });
  } catch (error) {
    if (error instanceof RenderError) throw error;
    if (/pixel limit/i.test(String(error))) throw new RenderError('LIMIT_EXCEEDED', 'Image exceeds the pixel limit');
    throw new RenderError('UNSUPPORTED_INPUT', 'Inputs must be PDF, JPEG, PNG, WebP or GIF');
  }
  const doc = await PDFDocument.create();
  const embedded = alpha ? await doc.embedPng(image.data) : await doc.embedJpg(image.data);
  // ponytail: never upscale past 1 px per point; small images stay small and sharp.
  const scale = Math.min(boxWidth / image.info.width, boxHeight / image.info.height, 1);
  const [width, height] = [image.info.width * scale, image.info.height * scale];
  doc.addPage([pageWidth, pageHeight]).drawImage(embedded, {
    x: (pageWidth - width) / 2,
    y: (pageHeight - height) / 2,
    width,
    height,
  });
  return doc.save();
}
