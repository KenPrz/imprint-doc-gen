import { compileDocx } from './docx.ts';
import type { CompiledTemplate, RenderOptions } from './docx.ts';
export { RenderError } from './errors.ts';
export type { CompiledTemplate, RenderOptions } from './docx.ts';
export const compile = compileDocx;
export async function render(template: Uint8Array | CompiledTemplate, data: Record<string, unknown>, options?: RenderOptions): Promise<Buffer> {
  const compiled = template instanceof Uint8Array ? await compile(template) : template;
  return compiled.render(data, options);
}
