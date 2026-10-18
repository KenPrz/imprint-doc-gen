import { Metrics } from './metrics.ts';
import { decodeBase64 } from './base64.ts';
export { decodeBase64 } from './base64.ts';
import { RenderError } from './errors.ts';
import { renderDocumentInWorker } from './jobs.ts';
import { convertToPdf } from './converter.ts';
import { ServiceError } from './service-error.ts';
import type { DocumentFormat, RenderOptions } from './types.ts';

export interface RenderInput extends Pick<RenderOptions, 'images' | 'locale' | 'timeZone' | 'restartNumbering' | 'charts'> {
  data: Record<string, unknown>;
  attachments?: Record<string, string>;
  output?: DocumentFormat | 'pdf';
}
export interface RenderResult { bytes: Buffer; format: DocumentFormat | 'pdf' }

export class RenderService {
  active = 0;
  readonly metrics = new Metrics();
  readonly concurrency: number;
  readonly converterUrl?: string;
  constructor(concurrency = 2, converterUrl?: string) {
    if(!Number.isInteger(concurrency)||concurrency<1||concurrency>8)throw new Error('Concurrency must be between 1 and 8');
    this.concurrency = concurrency; this.converterUrl = converterUrl;
  }
  get available() { return this.active < this.concurrency; }
  async render(template: Buffer, input: RenderInput): Promise<RenderResult> {
    if (!this.available) throw new ServiceError(503, 'BUSY');
    this.active++;
    const started = performance.now(); let ok = false, outputBytes = 0;
    try {
      const attachments: Record<string, Uint8Array> = Object.create(null);
      for (const [key, value] of Object.entries(input.attachments ?? {})) attachments[key] = decodeBase64(value);
      const result = await renderDocumentInWorker(template, input.data, attachments, 30_000, input.output, false, input);
      if (input.output !== 'pdf') { ok = true; outputBytes = result.bytes.length; return result; }
      if (result.format === 'txt' || result.format === 'csv') throw new RenderError('UNSUPPORTED_FORMAT', 'Text templates do not support PDF conversion');
      if (!this.converterUrl) throw new RenderError('CONVERTER_UNAVAILABLE', 'PDF conversion is not configured');
      const bytes = await convertToPdf(result.bytes, this.converterUrl, `document.${result.format}`); ok = true; outputBytes = bytes.length; return { bytes, format: 'pdf' };
    } finally { this.active--; this.metrics.render(ok, (performance.now() - started) / 1000, outputBytes); }
  }
}
