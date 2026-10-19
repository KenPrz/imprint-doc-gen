import { parentPort, workerData } from 'node:worker_threads';
import { RenderError } from './errors.ts';
try {
  let value: unknown;
  if (workerData.task === 'inspect')
    value = await (await import('./inspect.ts')).inspectTemplate(workerData.payload);
  else if (workerData.task === 'pdf') {
    const { processPdf, splitPdf } = await import('./pdf.ts');
    value = workerData.payload.ranges
      ? await splitPdf(workerData.payload.bytes, workerData.payload.ranges)
      : await processPdf(workerData.payload);
  } else if (workerData.task === 'barcode')
    value = await (await import('./barcode.ts')).barcode(workerData.payload);
  else if (workerData.task === 'assurance')
    value = await (await import('./assurance-task.ts')).executeAssurance(workerData.payload);
  else throw new RenderError('UNSUPPORTED_OPERATION', 'Unsupported operation');
  parentPort!.postMessage({ ok: true, value });
} catch (error) {
  parentPort!.postMessage({
    ok: false,
    code: error instanceof RenderError ? error.code : 'RENDER_FAILED',
    message: error instanceof RenderError ? error.message : 'Operation failed',
    part: error instanceof RenderError ? error.part : undefined,
  });
}
