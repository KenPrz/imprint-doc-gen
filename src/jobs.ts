import { Worker } from 'node:worker_threads';
import { RenderError } from './errors.ts';
import type { DocumentFormat } from './types.ts';

export interface RenderedDocument { bytes: Buffer; format: DocumentFormat }

export function renderDocumentInWorker(template: Uint8Array, data: Record<string, unknown>, attachments: Record<string, Uint8Array>, timeoutMs = 30_000, output?: DocumentFormat | 'pdf', validateOnly = false): Promise<RenderedDocument> {
  return new Promise((resolve, reject) => {
    const worker = new Worker(new URL(import.meta.url.endsWith('.ts') ? './render-worker.ts' : './render-worker.js', import.meta.url), {
      workerData: { template, data, attachments, output, validateOnly }, resourceLimits: { maxOldGenerationSizeMb: 256, maxYoungGenerationSizeMb: 32 },
      execArgv: [], // Do not inherit the development server's --watch protocol.
    });
    let finished = false;
    const finish = (error?: Error, document?: RenderedDocument) => {
      if (finished) return; finished = true; clearTimeout(timer);
      void worker.terminate();
      if (error) reject(error); else resolve(document!);
    };
    const timer = setTimeout(() => finish(new RenderError('RENDER_TIMEOUT', 'Render exceeded its time budget')), timeoutMs);
    worker.on('message', message => {
      if (message?.type !== 'imprint-render-result') return;
      if (message.ok && ['docx', 'xlsx'].includes(message.format)) finish(undefined, { bytes: Buffer.from(message.bytes), format: message.format });
      else finish(new RenderError(typeof message.code === 'string' ? message.code : 'RENDER_FAILED', message.message, message.part));
    });
    worker.once('error', () => finish(new RenderError('RENDER_FAILED', 'Render worker failed')));
    worker.once('exit', () => { if (!finished) finish(new RenderError('RENDER_FAILED', 'Render worker exited without a result')); });
  });
}

export async function renderInWorker(template: Uint8Array, data: Record<string, unknown>, attachments: Record<string, Uint8Array>, timeoutMs = 30_000): Promise<Buffer> {
  return (await renderDocumentInWorker(template, data, attachments, timeoutMs)).bytes;
}
