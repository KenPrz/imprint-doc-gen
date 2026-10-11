export type DocumentFormat = 'docx' | 'xlsx';
export interface RenderOptions { attachments?: Record<string, Uint8Array> }
export interface CompiledTemplate { readonly format: DocumentFormat; render(data: Record<string, unknown>, options?: RenderOptions): Promise<Buffer> }
export const mediaTypes = {
  docx: 'application/vnd.openxmlformats-officedocument.wordprocessingml.document',
  xlsx: 'application/vnd.openxmlformats-officedocument.spreadsheetml.sheet',
} as const;
