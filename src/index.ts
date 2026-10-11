import { compileDocx } from './docx.ts';
import { compileXlsx } from './xlsx.ts';
import { openOfficePackage } from './ooxml.ts';
import type { CompiledTemplate, RenderOptions } from './types.ts';
export { RenderError } from './errors.ts';
export type { CompiledTemplate, RenderOptions, DocumentFormat } from './types.ts';
export async function compile(input: Uint8Array): Promise<CompiledTemplate> {
  const document = await openOfficePackage(input);
  return document.format === 'docx' ? compileDocx(document) : compileXlsx(document);
}
export async function render(template: Uint8Array | CompiledTemplate, data: Record<string, unknown>, options?: RenderOptions): Promise<Buffer> {
  const compiled = template instanceof Uint8Array ? await compile(template) : template;
  return compiled.render(data, options);
}
