import { Worker } from 'node:worker_threads';
import { RenderError } from './errors.ts';
export function runTask<T>(
  task: 'inspect' | 'pdf' | 'barcode',
  payload: unknown,
  timeout = 30000,
): Promise<T> {
  return new Promise((resolve, reject) => {
    const worker = new Worker(
      new URL(import.meta.url.endsWith('.ts') ? './task-worker.ts' : './task-worker.js', import.meta.url),
      {
        workerData: { task, payload },
        execArgv: [],
        resourceLimits: { maxOldGenerationSizeMb: 256, maxYoungGenerationSizeMb: 32 },
      },
    );
    let done = false;
    const finish = (error?: Error, value?: T) => {
      if (done) return;
      done = true;
      clearTimeout(timer);
      void worker.terminate();
      if (error) reject(error);
      else resolve(value!);
    };
    const timer = setTimeout(
      () => finish(new RenderError('RENDER_TIMEOUT', 'Operation exceeded its time budget')),
      timeout,
    );
    worker.once('message', (message) =>
      message.ok
        ? finish(undefined, message.value)
        : finish(new RenderError(message.code, message.message, message.part)),
    );
    worker.once('error', () => finish(new RenderError('RENDER_FAILED', 'Operation worker failed')));
    worker.once('exit', () =>
      finish(new RenderError('RENDER_FAILED', 'Operation worker exited without a result')),
    );
  });
}
