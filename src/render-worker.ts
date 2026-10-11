import { parentPort, workerData } from 'node:worker_threads';
import { compile, RenderError } from './index.ts';
try {
  const template = await compile(workerData.template);
  if (workerData.output && workerData.output !== 'pdf' && workerData.output !== template.format) throw new RenderError('UNSUPPORTED_FORMAT', 'Output must match the template format, or be PDF');
  const result = workerData.validateOnly ? Buffer.alloc(0) : await template.render(workerData.data, { attachments: workerData.attachments });
  parentPort!.postMessage({ type: 'imprint-render-result', ok: true, bytes: result, format: template.format });
} catch (error) {
  parentPort!.postMessage({ type: 'imprint-render-result', ok: false, code: error instanceof RenderError ? error.code : 'RENDER_FAILED', message: error instanceof RenderError ? error.message : 'Document rendering failed', part: error instanceof RenderError ? error.part : undefined });
}
