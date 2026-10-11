import { Worker } from 'node:worker_threads';
import { RenderError } from './errors.ts';

export function renderInWorker(template: Uint8Array, data: Record<string, unknown>, attachments: Record<string, Uint8Array>, timeoutMs = 30_000): Promise<Buffer> {
  return new Promise((resolve, reject) => {
    const worker = new Worker(new URL(import.meta.url.endsWith('.ts') ? './render-worker.ts' : './render-worker.js', import.meta.url), {
      workerData: { template, data, attachments }, resourceLimits: { maxOldGenerationSizeMb: 256, maxYoungGenerationSizeMb: 32 },
      execArgv: [], // Do not inherit the development server's --watch protocol.
    });
    let finished = false;
    const finish = (error?: Error, bytes?: Buffer) => {
      if (finished) return; finished = true; clearTimeout(timer);
      void worker.terminate();
      if (error) reject(error); else resolve(bytes!);
    };
    const timer = setTimeout(() => finish(new RenderError('RENDER_TIMEOUT', 'Render exceeded its time budget')), timeoutMs);
    worker.on('message', message => {
      if (message?.type !== 'imprint-render-result') return;
      if (message.ok) finish(undefined, Buffer.from(message.bytes));
      else finish(new RenderError(typeof message.code === 'string' ? message.code : 'RENDER_FAILED', message.message, message.part));
    });
    worker.once('error', () => finish(new RenderError('RENDER_FAILED', 'Render worker failed')));
    worker.once('exit', () => { if (!finished) finish(new RenderError('RENDER_FAILED', 'Render worker exited without a result')); });
  });
}
