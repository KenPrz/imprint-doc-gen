import { RenderError } from './errors.ts';
import { AsyncLocalStorage } from 'node:async_hooks';
export interface ExecutionTraceOptions {
  includeValues?: boolean;
  maxEvents?: number;
  maxValueLength?: number;
}
export interface ExecutionTraceEvent {
  sequence: number;
  scopedPath: string;
  part: string;
  scopeDepth: number;
  formatters: string[];
  status: 'evaluated' | 'error';
  missingInput: boolean;
  inputType: string;
  outputType?: string;
  errorCode?: string;
  outputValue?: string | number | boolean | null;
  valueOmitted?: boolean;
}
export interface ExecutionTrace {
  version: 1;
  phase: 'execution';
  coverage: 'standard-lookup-operations';
  events: ExecutionTraceEvent[];
  droppedEvents: number;
  truncated: boolean;
}
interface Context {
  trace: ExecutionTrace;
  options: Required<ExecutionTraceOptions>;
  valueCharacters: number;
  active: boolean;
}
const storage = new AsyncLocalStorage<Context>();
const valueType = (v: unknown) => (v === null ? 'null' : Array.isArray(v) ? 'array' : typeof v);
export function recordLookupExecution(event: {
  scopedPath: string;
  part: string;
  scopeDepth: number;
  formatters: string[];
  missingInput: boolean;
  input: unknown;
  output?: unknown;
  errorCode?: string;
}): void {
  const context = storage.getStore();
  if (!context || !context.active) return;
  const { trace, options } = context;
  if (trace.events.length >= options.maxEvents) {
    trace.droppedEvents++;
    trace.truncated = true;
    return;
  }
  const entry: ExecutionTraceEvent = {
    sequence: trace.events.length + 1,
    scopedPath: event.scopedPath,
    part: event.part,
    scopeDepth: event.scopeDepth,
    formatters: event.formatters,
    status: event.errorCode ? 'error' : 'evaluated',
    missingInput: event.missingInput,
    inputType: valueType(event.input),
  };
  if (event.errorCode) entry.errorCode = event.errorCode;
  else {
    entry.outputType = valueType(event.output);
    if (options.includeValues) {
      if (typeof event.output === 'string') {
        if (
          event.output.length <= options.maxValueLength &&
          context.valueCharacters + event.output.length <= 1_000_000
        ) {
          entry.outputValue = event.output;
          context.valueCharacters += event.output.length;
        } else entry.valueOmitted = true;
      } else if (
        event.output === null ||
        typeof event.output === 'boolean' ||
        (typeof event.output === 'number' && Number.isFinite(event.output))
      )
        entry.outputValue = event.output;
      else entry.valueOmitted = true;
    }
  }
  trace.events.push(entry);
}
/** Captures actual standard lookup operations; catches callback failure without exposing arbitrary error text. */
export async function captureExecutionTrace<T>(
  operation: () => Promise<T>,
  options: ExecutionTraceOptions = {},
): Promise<
  | { ok: true; value: T; trace: ExecutionTrace }
  | { ok: false; error: { code: string }; trace: ExecutionTrace }
> {
  const normalized = {
    includeValues: options.includeValues ?? false,
    maxEvents: options.maxEvents ?? 10000,
    maxValueLength: options.maxValueLength ?? 1000,
  };
  if (
    typeof normalized.includeValues !== 'boolean' ||
    !Number.isSafeInteger(normalized.maxEvents) ||
    normalized.maxEvents < 1 ||
    normalized.maxEvents > 100000 ||
    !Number.isSafeInteger(normalized.maxValueLength) ||
    normalized.maxValueLength < 0 ||
    normalized.maxValueLength > 10000
  )
    throw new Error('Invalid execution trace options');
  const trace: ExecutionTrace = {
    version: 1,
    phase: 'execution',
    coverage: 'standard-lookup-operations',
    events: [],
    droppedEvents: 0,
    truncated: false,
  };
  const context: Context = { trace, options: normalized, valueCharacters: 0, active: true };
  return storage.run(context, async () => {
    try {
      return { ok: true as const, value: await operation(), trace };
    } catch (error) {
      return {
        ok: false as const,
        error: {
          code: error instanceof RenderError ? error.code : 'RENDER_FAILED',
        },
        trace,
      };
    } finally {
      context.active = false;
    }
  });
}
