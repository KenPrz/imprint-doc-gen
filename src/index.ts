import { compilePptx } from './pptx.ts';
import { compileOpenDocument } from './opendocument.ts';
import { renderContext, expressionSettings } from './expressions.ts';
import { compileDocx } from './docx.ts';
import { compileXlsx } from './xlsx.ts';
import { openOfficePackage } from './ooxml.ts';
import type { CompiledTemplate, RenderOptions } from './types.ts';
export { RenderError } from './errors.ts';
export type { CompiledTemplate, RenderOptions, DocumentFormat } from './types.ts';
export async function compile(input: Uint8Array): Promise<CompiledTemplate> {
  const document = await openOfficePackage(input);
  const compiled = await (document.format === 'docx' ? compileDocx(document) : document.format === 'xlsx' ? compileXlsx(document) : document.format === 'pptx' ? compilePptx(document) : compileOpenDocument(document));
  return Object.freeze({ format: compiled.format, render(data: Record<string, unknown>, options: RenderOptions = {}) { return renderContext.run(expressionSettings(options.locale, options.timeZone), () => compiled.render(data, options)); } });
}
export async function render(template: Uint8Array | CompiledTemplate, data: Record<string, unknown>, options?: RenderOptions): Promise<Buffer> {
  const compiled = template instanceof Uint8Array ? await compile(template) : template;
  return compiled.render(data, options);
}

export { inspectTemplate } from './inspect.ts';
export { processPdf, splitPdf } from './pdf.ts';
export { barcode } from './barcode.ts';
export type { PdfInput, PdfOptions } from './pdf.ts';
export type { BarcodeOptions } from './barcode.ts';
export type { ImageOptions } from './types.ts';
