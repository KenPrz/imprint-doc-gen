import { parentPort, workerData } from 'node:worker_threads';
import { render, RenderError } from './index.ts';
try {
  const result = await render(workerData.template, workerData.data, { attachments: workerData.attachments });
  parentPort!.postMessage({ type: 'imprint-render-result', ok: true, bytes: result });
} catch (error) {
  parentPort!.postMessage({ type: 'imprint-render-result', ok: false, code: error instanceof RenderError ? error.code : 'RENDER_FAILED', message: error instanceof RenderError ? error.message : 'Document rendering failed', part: error instanceof RenderError ? error.part : undefined });
}
