import { parentPort, workerData } from 'node:worker_threads';
import { compile, RenderError } from './index.ts';
import { readPackage } from './package.ts';
import { resolveSheetRule, sheetData, templateSheetRule } from './sheets.ts';
try {
  const template = await compile(workerData.template);
  if (workerData.output && workerData.output !== 'pdf' && workerData.output !== template.format) throw new RenderError('UNSUPPORTED_FORMAT', 'Output must match the template format, or be PDF');
  const { sheets: request, maxSheets = 50, ...options } = workerData.options ?? {};
  const textual = template.format === 'txt' || template.format === 'csv';
  const rule = workerData.validateOnly ? null : resolveSheetRule(textual ? null : templateSheetRule(await readPackage(workerData.template)), request);
  if (rule) {
    if (workerData.output !== 'pdf') throw new RenderError('SHEETS_REQUIRE_PDF', 'Paginated templates render only to PDF');
    const sheets: Buffer[] = [];
    for (const data of sheetData(workerData.data, rule, request?.perSheet, maxSheets)) sheets.push(await template.render(data, { ...options, attachments: workerData.attachments }));
    parentPort!.postMessage({ type: 'imprint-render-result', ok: true, bytes: Buffer.alloc(0), sheets, format: template.format });
  } else {
    const result = workerData.validateOnly ? Buffer.alloc(0) : await template.render(workerData.data, { ...options, attachments: workerData.attachments });
    parentPort!.postMessage({ type: 'imprint-render-result', ok: true, bytes: result, format: template.format });
  }
} catch (error) {
  parentPort!.postMessage({ type: 'imprint-render-result', ok: false, code: error instanceof RenderError ? error.code : 'RENDER_FAILED', message: error instanceof RenderError ? error.message : 'Document rendering failed', part: error instanceof RenderError ? error.part : undefined });
}
