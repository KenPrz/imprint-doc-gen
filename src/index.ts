import { RenderError } from './errors.ts';
import { limits } from './package.ts';
import { updateCharts } from './charts.ts';
export { updateCharts } from './charts.ts';
export type { ChartData } from './charts.ts';
import { compileText, textMagic } from './text-template.ts';
export { compileText, encodeTextTemplate } from './text-template.ts';
import { compilePptx } from './pptx.ts';
import { compileOpenDocument } from './opendocument.ts';
import { renderContext, expressionSettings } from './expressions.ts';
import { compileDocx } from './docx.ts';
import { compileXlsx } from './xlsx.ts';
import { openOfficePackage } from './ooxml.ts';
import type { CompiledTemplate, RenderOptions } from './types.ts';
export { RenderError } from './errors.ts';
export type { CompiledTemplate, RenderOptions, DocumentFormat } from './types.ts';
export async function compile(input: Uint8Array): Promise<CompiledTemplate> {
  if(input.byteLength>limits.archive)throw new RenderError('LIMIT_EXCEEDED','Template exceeds 8 MiB');
  if (Buffer.from(input).subarray(0, textMagic.length).toString() === textMagic) {
    let envelope;
    try { envelope = JSON.parse(new TextDecoder('utf-8', { fatal: true }).decode(input).slice(textMagic.length)); } catch { throw new RenderError('TEMPLATE_SYNTAX','Invalid UTF-8 text template envelope'); }
    if(!envelope || typeof envelope !== 'object' || Object.keys(envelope).sort().join(',') !== 'format,source')throw new RenderError('TEMPLATE_SYNTAX','Invalid text template envelope');
    return compileText(envelope.format, envelope.source);
  }
  const document = await openOfficePackage(input);
  const compiled = await (document.format === 'docx' ? compileDocx(document) : document.format === 'xlsx' ? compileXlsx(document) : document.format === 'pptx' ? compilePptx(document) : compileOpenDocument(document));
  return Object.freeze({ format: compiled.format, render(data: Record<string, unknown>, options: RenderOptions = {}) { return renderContext.run(expressionSettings(options.locale, options.timeZone), async () => { const bytes = await compiled.render(data, options); return options.charts ? updateCharts(bytes, options.charts) : bytes; }); } });
}
export async function render(template: Uint8Array | CompiledTemplate, data: Record<string, unknown>, options?: RenderOptions): Promise<Buffer> {
  const compiled = template instanceof Uint8Array ? await compile(template) : template;
  return compiled.render(data, options);
}

export { inspectTemplate } from './inspect.ts';
export { processPdf, splitPdf } from './pdf.ts';
export { barcode } from './barcode.ts';
export type { PdfInput, PdfOptions } from './pdf.ts';
export type { BarcodeOptions } from './barcode.ts';
export type { ImageOptions } from './types.ts';

export { checkContractSchema, validateContract, contractTypeScript, contractSample, evaluateAssertions } from './contracts.ts';
export type { ContractSchema, ContractIssue, BusinessAssertion } from './contracts.ts';
export { traceTemplateInputs, renderWithTrace } from './trace.ts';
export type { PreflightTrace, PreflightTraceEntry, TraceOptions } from './trace.ts';
export type { ExecutionTrace, ExecutionTraceEvent, ExecutionTraceOptions } from './execution-trace.ts';
export { snapshotDocument, compareDocuments } from './document-diff.ts';
export type { DocumentSnapshot, DocumentDiff, DocumentChange, DocumentUnit } from './document-diff.ts';
export { runRegressionSuite } from './regression.ts';
export type { RegressionCase, RegressionResult } from './regression.ts';
export { createReplayBundle, replayBundle, replayEngineFingerprint } from './replay.ts';
export type { ReplayBundle } from './replay.ts';
export { migrateCarboneText, compileCarboneText, inventoryCrystalMetadata } from './template-migration.ts';
export { ComponentRegistry } from './components.ts';
export { validateDelivery, pdfStructureValidator, artifactDigest } from './delivery.ts';
export type { ArtifactValidator, DeliveryEvidence, ValidatedDelivery } from './delivery.ts';
export { createReportSnapshot, verifyReportSnapshot, reportSnapshotHtml, renderReportSnapshot } from './report-view.ts';
export type { ReportDefinition, ReportColumn, ReportFilter, ReportSnapshot } from './report-view.ts';
export { reviewTemplateCandidate, acceptTemplateCandidate } from './template-review.ts';
export type { TemplateReview } from './template-review.ts';
export { comparePageImages } from './visual-diff.ts';
export type { VisualDiff, VisualDiffOptions, VisualMask } from './visual-diff.ts';
export { veraPdfValidator, parseVeraPdfReport } from './verapdf.ts';
export type { VeraPdfProfile } from './verapdf.ts';
export { comparePdfDocuments } from './pdf-regression.ts';
export type { PdfRegressionOptions, PdfRegressionResult } from './pdf-regression.ts';
