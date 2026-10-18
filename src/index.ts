import { RenderError } from './errors.ts';
import { limits } from './package.ts';
import { updateCharts } from './charts.ts';
export { updateCharts } from './charts.ts';
export type { ChartData } from './charts.ts';
import { compileText, textMagic } from './text-template.ts';
export { compileText, encodeTextTemplate } from './text-template.ts';
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
  if(input.byteLength>limits.archive)throw new RenderError('LIMIT_EXCEEDED','Template exceeds 8 MiB');
  if (Buffer.from(input).subarray(0, textMagic.length).toString() === textMagic) {
    let envelope;
    try { envelope = JSON.parse(new TextDecoder('utf-8', { fatal: true }).decode(input).slice(textMagic.length)); } catch { throw new RenderError('TEMPLATE_SYNTAX','Invalid UTF-8 text template envelope'); }
    if(!envelope || typeof envelope !== 'object' || Object.keys(envelope).sort().join(',') !== 'format,source')throw new RenderError('TEMPLATE_SYNTAX','Invalid text template envelope');
    return compileText(envelope.format, envelope.source);
  }
  const document = await openOfficePackage(input);
  const compiled = await (document.format === 'docx' ? compileDocx(document) : document.format === 'xlsx' ? compileXlsx(document) : document.format === 'pptx' ? compilePptx(document) : compileOpenDocument(document));
  return Object.freeze({ format: compiled.format, render(data: Record<string, unknown>, options: RenderOptions = {}) { return renderContext.run(expressionSettings(options.locale, options.timeZone), async () => { const bytes = await compiled.render(data, options); return options.charts ? updateCharts(bytes, options.charts) : bytes; }); } });
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
