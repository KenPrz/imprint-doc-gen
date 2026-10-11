import { decodeBase64 } from './base64.ts';
export { decodeBase64 } from './base64.ts';
import { RenderError } from './errors.ts';
import { renderDocumentInWorker } from './jobs.ts';
import { convertToPdf } from './converter.ts';
import { ServiceError } from './service-error.ts';
import type { DocumentFormat } from './types.ts';

export interface RenderInput {
  data: Record<string, unknown>;
  attachments?: Record<string, string>;
  output?: DocumentFormat | 'pdf';
}
export interface RenderResult { bytes: Buffer; format: DocumentFormat | 'pdf' }

export class RenderService {
  active = 0;
  readonly concurrency: number;
  readonly converterUrl?: string;
  constructor(concurrency = 2, converterUrl?: string) {
    this.concurrency = concurrency; this.converterUrl = converterUrl;
  }
  get available() { return this.active < this.concurrency; }
  async render(template: Buffer, input: RenderInput): Promise<RenderResult> {
    if (!this.available) throw new ServiceError(503, 'BUSY');
    this.active++;
    try {
      const attachments: Record<string, Uint8Array> = Object.create(null);
      for (const [key, value] of Object.entries(input.attachments ?? {})) attachments[key] = decodeBase64(value);
      const result = await renderDocumentInWorker(template, input.data, attachments, 30_000, input.output);
      if (input.output !== 'pdf') return result;
      if (!this.converterUrl) throw new RenderError('CONVERTER_UNAVAILABLE', 'PDF conversion is not configured');
      return { bytes: await convertToPdf(result.bytes, this.converterUrl, `document.${result.format}`), format: 'pdf' };
    } finally { this.active--; }
  }
}
