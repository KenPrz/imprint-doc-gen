import type { ChartData } from './charts.ts';
export type DocumentFormat = 'docx' | 'xlsx' | 'pptx' | 'odt' | 'ods' | 'odp' | 'txt' | 'csv';
export interface ImageOptions {
  fit?: 'contain' | 'cover' | 'fill'; width?: number; height?: number;
  rotate?: number; background?: string;
  crop?: { left: number; top: number; width: number; height: number };
}
export interface RenderOptions { restartNumbering?: boolean; charts?: Record<string, ChartData>; attachments?: Record<string, Uint8Array>; images?: Record<string, ImageOptions>; locale?: string; timeZone?: string }
export interface CompiledTemplate { readonly format: DocumentFormat; render(data: Record<string, unknown>, options?: RenderOptions): Promise<Buffer> }
export const mediaTypes = {
  txt: 'text/plain; charset=utf-8', csv: 'text/csv; charset=utf-8',
  pptx: 'application/vnd.openxmlformats-officedocument.presentationml.presentation',
  odt: 'application/vnd.oasis.opendocument.text', ods: 'application/vnd.oasis.opendocument.spreadsheet', odp: 'application/vnd.oasis.opendocument.presentation',
  docx: 'application/vnd.openxmlformats-officedocument.wordprocessingml.document',
  xlsx: 'application/vnd.openxmlformats-officedocument.spreadsheetml.sheet',
} as const;
