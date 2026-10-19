import { inspectTemplate, type Inspection, type TemplateField } from './inspect.ts';
import { validateContract } from './contracts.ts';
import { lookup, checkPath } from './scope.ts';
import { RenderError } from './errors.ts';
import {
  parseExpression,
  formatValue,
  expressionSettings,
  renderContext,
  type Formatter,
} from './expressions.ts';
/** Static evidence only: no claim that a field executed, appeared on a page, or passed filters. */
export interface PreflightTraceEntry extends TemplateField {
  status: 'present' | 'missing' | 'requires-render-context';
  formatters: string[];
  inputType?: string;
  evaluation?: { status: 'evaluated' | 'error'; outputType?: string; errorCode?: string; value?: unknown };
}
export interface PreflightTrace {
  version: 1;
  phase: 'preflight';
  coverage: 'inspector-discovered-fields';
  settings: { locale: string; timeZone: string };
  inspection: Inspection;
  entries: PreflightTraceEntry[];
}
export interface TraceOptions {
  includeValues?: boolean;
  locale?: string;
  timeZone?: string;
}
const valueType = (v: unknown) => (v === null ? 'null' : Array.isArray(v) ? 'array' : typeof v);
/** Evaluates scalar standard formatters; loops, filters, images and output coordinates need rendering context.
 * Values are excluded by default. includeValues explicitly opts into potentially sensitive output data.
 */
export async function traceTemplateInputs(
  bytes: Uint8Array,
  data: Record<string, unknown>,
  options: TraceOptions = {},
): Promise<PreflightTrace> {
  if (!validateContract({ type: 'object' }, data).valid)
    throw new Error('Trace data must be a plain JSON object');
  if (options.includeValues !== undefined && typeof options.includeValues !== 'boolean')
    throw new Error('includeValues must be boolean');
  const settings = expressionSettings(options.locale, options.timeZone);
  const inspection = await inspectTemplate(bytes);
  const entries = renderContext.run(settings, () =>
    inspection.fields.map((field) => {
      let formatters: Formatter[] = [];
      const standard = field.expression.startsWith('{d.');
      if (standard) {
        try {
          formatters = parseExpression(field.expression.slice(1, -1), field.part).formatters;
        } catch {
          /* Inspection already reports unsupported expressions. */
        }
      }
      const entry: PreflightTraceEntry = {
        ...field,
        formatters: formatters.map((f) => f.name),
        status: 'requires-render-context',
      };
      if (field.path.includes('[') || field.path.startsWith('../')) return entry;
      let value: unknown = null;
      try {
        checkPath(field.path, field.part);
        value = lookup([data], field.path, field.part);
        entry.status = 'present';
        entry.inputType = valueType(value);
      } catch {
        entry.status = 'missing';
      }
      if (standard && field.kind !== 'image') {
        try {
          const output = formatValue(value, formatters);
          entry.evaluation = {
            status: 'evaluated',
            outputType: valueType(output),
            ...(options.includeValues ? { value: output } : {}),
          };
        } catch (e) {
          entry.evaluation = {
            status: 'error',
            errorCode: e instanceof RenderError ? e.code : 'EVALUATION_FAILED',
          };
        }
      }
      return entry;
    }),
  );
  return {
    version: 1,
    phase: 'preflight',
    coverage: 'inspector-discovered-fields',
    settings,
    inspection,
    entries,
  };
}

/** Captures actual standard lookup calls; no page geometry or global array-index attribution is inferred. */
export async function renderWithTrace(
  template: import('./types.ts').CompiledTemplate | Uint8Array,
  data: Record<string, unknown>,
  renderOptions: import('./types.ts').RenderOptions = {},
  traceOptions: import('./execution-trace.ts').ExecutionTraceOptions = {},
) {
  const { render } = await import('./index.ts');
  const { captureExecutionTrace } = await import('./execution-trace.ts');
  return captureExecutionTrace(() => render(template, data, renderOptions), traceOptions);
}
