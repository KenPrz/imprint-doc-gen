import { RenderError } from './errors.ts';

export const forbidden = new Set(['__proto__', 'constructor', 'prototype']);
export function checkPath(path: string, part: string) {
  while (path.startsWith('../')) path = path.slice(3);
  if (/^[#\/^]/.test(path)) throw new RenderError('UNSUPPORTED_TEMPLATE', 'Only #each loops and scalar paths are supported', part);
  if (!/^[A-Za-z_$][\w$]*(?:\.[A-Za-z_$][\w$]*)*$/.test(path) || path.split('.').some(p => forbidden.has(p))) throw new RenderError('TEMPLATE_EXPRESSION', `Unsupported data path: ${path}`, part);
}
export type Scope = readonly Record<string, unknown>[];
export function lookup(scope: Scope, path: string, part: string): unknown {
  let depth = scope.length - 1;
  while (path.startsWith('../')) { depth--; path = path.slice(3); }
  if (depth < 0) throw new RenderError('MISSING_DATA', 'Data path escapes the root scope', part);
  let value: unknown = scope[depth];
  for (const key of path.split('.')) {
    if (value === null || typeof value !== 'object' || !Object.hasOwn(value, key)) throw new RenderError('MISSING_DATA', `Missing value: ${path}`, part);
    value = (value as Record<string, unknown>)[key];
  }
  return value;
}
export function scalar(value: unknown, path: string, part: string): string {
  if (value === null) return '';
  if (!['string', 'number', 'boolean'].includes(typeof value) || (typeof value === 'number' && !Number.isFinite(value))) throw new RenderError('INVALID_DATA', `Expected a scalar at ${path}`, part);
  const text = String(value);
  if (/[\u0000-\u0008\u000b\u000c\u000e-\u001f\ud800-\udfff\ufffe\uffff]/u.test(text)) throw new RenderError('INVALID_DATA', `Value contains invalid XML characters: ${path}`, part);
  if (text.length > 1_000_000) throw new RenderError('LIMIT_EXCEEDED', 'A replacement value exceeds its character budget', part);
  if (/[\r\n\t]/.test(text)) throw new RenderError('UNSUPPORTED_DATA', 'Multiline and tabbed replacements require a later rendering milestone', part);
  return text;
}
