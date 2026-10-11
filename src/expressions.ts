import { AsyncLocalStorage } from 'node:async_hooks';
import { RenderError } from './errors.ts';
import { checkPath, lookup } from './scope.ts';

export interface Formatter {
  name: string;
  args: (string | number | boolean | null)[];
}
export interface ArrayFilter {
  arrayPath: string;
  field: string;
  operator: string;
  value: string | number | boolean | null;
}
export const renderContext = new AsyncLocalStorage<{ locale: string; timeZone: string }>();
const settings = () => renderContext.getStore() ?? { locale: 'en-US', timeZone: 'UTC' };
export function expressionSettings(locale = 'en-US', timeZone = 'UTC') {
  try {
    new Intl.NumberFormat(locale);
    new Intl.DateTimeFormat(locale, { timeZone });
  } catch {
    throw new RenderError('INVALID_LOCALE', 'Invalid locale or time zone');
  }
  return { locale, timeZone };
}
export function parseExpression(tag: string, part: string) {
  if (tag.length > 2048) throw new RenderError('LIMIT_EXCEEDED', 'Expression exceeds 2048 characters', part);
  const segments: string[] = [];
  let start = 0,
    quoted = false,
    escaped = false,
    depth = 0;
  for (let n = 0; n < tag.length; n++) {
    const c = tag[n]!;
    if (quoted) {
      if (escaped) escaped = false;
      else if (c === '\\') escaped = true;
      else if (c === '"') quoted = false;
      continue;
    }
    if (c === '"') quoted = true;
    else if (c === '(') depth++;
    else if (c === ')') depth--;
    else if (c === ':' && depth === 0) {
      segments.push(tag.slice(start, n));
      start = n + 1;
    }
    if (depth < 0 || depth > 1)
      throw new RenderError('UNSUPPORTED_EXPRESSION', 'Nested calls are not supported', part);
  }
  if (quoted || depth !== 0) throw new RenderError('TEMPLATE_SYNTAX', 'Unclosed formatter call', part);
  segments.push(tag.slice(start));
  if (segments.length > 17) throw new RenderError('LIMIT_EXCEEDED', 'At most 16 expression operations', part);
  let path = segments.shift()!;
  const formatters: Formatter[] = [],
    filters: ArrayFilter[] = [];
  for (const segment of segments) {
    const match = /^([a-zA-Z]+)\((.*)\)(.*)$/.exec(segment);
    if (!match) throw new RenderError('UNSUPPORTED_EXPRESSION', 'Expected formatter(arguments)', part);
    let args: Formatter['args'];
    try {
      args = JSON.parse(`[${match[2]}]`);
    } catch {
      throw new RenderError('TEMPLATE_SYNTAX', 'Formatter arguments must be JSON literals', part);
    }
    if (
      args.some((v) => v !== null && !['string', 'number', 'boolean'].includes(typeof v)) ||
      args.length > 3
    )
      throw new RenderError('UNSUPPORTED_EXPRESSION', 'Expected up to three scalar arguments', part);
    const name = match[1]!,
      suffix = match[3]!;
    if (name === 'where') {
      if (
        !/^\[i(?:\+1)?\](?:\.|$)/.test(suffix) ||
        ![2, 3].includes(args.length) ||
        typeof args[0] !== 'string'
      )
        throw new RenderError(
          'UNSUPPORTED_EXPRESSION',
          'where(field,value) or where(field,operator,value) must precede an array iterator',
          part,
        );
      checkPath(args[0], part);
      const operator = args.length === 2 ? 'eq' : String(args[1]);
      if (!['eq', 'ne', 'gt', 'gte', 'lt', 'lte', 'contains'].includes(operator))
        throw new RenderError('UNSUPPORTED_EXPRESSION', 'Unsupported filter operator', part);
      filters.push({ arrayPath: path.slice(2), field: args[0], operator, value: args.at(-1)! });
      path += suffix;
    } else {
      if (
        suffix ||
        !['number', 'currency', 'date', 'upper', 'lower', 'default', 'if', 'dateSerial', 'required'].includes(
          name,
        )
      )
        throw new RenderError('UNSUPPORTED_EXPRESSION', `Unsupported formatter: ${name}`, part);
      const valid =
        name === 'number'
          ? args.length <= 1 &&
            (!args.length || (Number.isInteger(args[0]) && Number(args[0]) >= 0 && Number(args[0]) <= 20))
          : name === 'currency'
            ? args.length === 1 && typeof args[0] === 'string' && /^[A-Z]{3}$/.test(args[0])
            : name === 'date'
              ? args.length <= 1 &&
                (!args.length || ['short', 'medium', 'long', 'full', 'iso'].includes(String(args[0])))
              : name === 'dateSerial'
                ? args.length <= 1 && (!args.length || ['1900', '1904'].includes(String(args[0])))
                : ['upper', 'lower', 'required'].includes(name)
                  ? !args.length
                  : name === 'default'
                    ? args.length === 1
                    : args.length === 2;
      if (!valid) throw new RenderError('UNSUPPORTED_EXPRESSION', `Invalid arguments for ${name}`, part);
      formatters.push({ name, args });
    }
  }
  return { path, formatters, filters };
}
function date(value: unknown): Date {
  if (
    typeof value !== 'string' ||
    !/^\d{4}-\d{2}-\d{2}(?:T\d{2}:\d{2}:\d{2}(?:\.\d{1,3})?(?:Z|[+-]\d{2}:\d{2}))?$/.test(value)
  )
    throw new RenderError('INVALID_DATA', 'Dates must be ISO dates or timestamps with explicit offsets');
  const result = new Date(value);
  if (
    !Number.isFinite(result.getTime()) ||
    (value.length === 10 && result.toISOString().slice(0, 10) !== value)
  )
    throw new RenderError('INVALID_DATA', 'Invalid date');
  return result;
}
export function formatValue(value: unknown, formatters: Formatter[] = []): unknown {
  const { locale, timeZone } = settings();
  for (const { name, args } of formatters) {
    if (name === 'required') {
      if (value === null || value === undefined || value === '')
        throw new RenderError('MISSING_DATA', 'Required field is missing');
      continue;
    }
    if (name === 'default') {
      if (value === null || value === undefined || value === '') value = args[0];
      continue;
    }
    if (name === 'if') {
      if (value !== null && typeof value === 'object')
        throw new RenderError('INVALID_DATA', 'Conditions require scalar data');
      value = value ? args[0] : args[1];
      continue;
    }
    if (value === null || value === undefined || value === '') continue;
    if (name === 'number' || name === 'currency') {
      if (typeof value !== 'number' || !Number.isFinite(value))
        throw new RenderError('INVALID_DATA', `${name} requires a finite JSON number`);
      value = new Intl.NumberFormat(
        locale,
        name === 'currency'
          ? { style: 'currency', currency: String(args[0]) }
          : args.length
            ? { minimumFractionDigits: Number(args[0]), maximumFractionDigits: Number(args[0]) }
            : {},
      ).format(value);
    } else if (name === 'date')
      value =
        args[0] === 'iso'
          ? date(value).toISOString().slice(0, 10)
          : new Intl.DateTimeFormat(locale, {
              dateStyle: (args[0] ?? 'medium') as 'medium',
              timeZone,
            }).format(date(value));
    else if (name === 'dateSerial') {
      const valueDate = date(value),
        epoch = args[0] === '1904' ? Date.UTC(1904, 0, 1) : Date.UTC(1899, 11, 31);
      const serial = (valueDate.getTime() - epoch) / 86400000;
      value = args[0] === '1904' || valueDate.getTime() < Date.UTC(1900, 2, 1) ? serial : serial + 1;
      if (Number(value) < 0)
        throw new RenderError('INVALID_DATA', 'Date precedes the selected spreadsheet epoch');
    } else {
      if (typeof value !== 'string') throw new RenderError('INVALID_DATA', `${name} requires text`);
      value = name === 'upper' ? value.toLocaleUpperCase(locale) : value.toLocaleLowerCase(locale);
    }
  }
  return value;
}
export function filterArray(value: unknown, filters: ArrayFilter[] = [], part: string): unknown {
  if (!Array.isArray(value) || !filters.length) return value;
  if (value.length > 10000)
    throw new RenderError('LIMIT_EXCEEDED', 'Filtered arrays exceed 10000 items', part);
  return value.filter((item) =>
    filters.every((filter) => {
      let actual: unknown;
      try {
        actual = lookup([item], filter.field, part);
      } catch (error) {
        if (!(error instanceof RenderError) || error.code !== 'MISSING_DATA') throw error;
        actual = null;
      }
      const expected = filter.value;
      if (filter.operator === 'eq') return actual === expected;
      if (filter.operator === 'ne') return actual !== expected;
      if (filter.operator === 'contains')
        return typeof actual === 'string' && typeof expected === 'string' && actual.includes(expected);
      if (typeof actual !== typeof expected || !['number', 'string'].includes(typeof actual)) return false;
      const a = actual as number,
        b = expected as number;
      return filter.operator === 'gt'
        ? a > b
        : filter.operator === 'gte'
          ? a >= b
          : filter.operator === 'lt'
            ? a < b
            : a <= b;
    }),
  );
}
