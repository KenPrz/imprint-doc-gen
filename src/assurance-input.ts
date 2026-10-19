import { types } from 'node:util';
import { validateData } from './data.ts';
import type { RenderOptions } from './types.ts';
import { RenderError } from './errors.ts';
/** Snapshot options without invoking accessors; attachments remain byte arrays. */
export function cloneRenderOptions(options: RenderOptions = {}): RenderOptions {
  if (
    !options ||
    typeof options !== 'object' ||
    types.isProxy(options) ||
    ![Object.prototype, null].includes(Object.getPrototypeOf(options)) ||
    Object.getOwnPropertySymbols(options).length
  )
    throw new RenderError('INVALID_OPTIONS', 'Expected plain render options');
  const json: Record<string, unknown> = {},
    attachments: Record<string, Uint8Array> = Object.create(null);
  let hasAttachments = false,
    total = 0;
  for (const [key, d] of Object.entries(Object.getOwnPropertyDescriptors(options))) {
    if (
      !d.enumerable ||
      !Object.hasOwn(d, 'value') ||
      !['locale', 'timeZone', 'restartNumbering', 'images', 'charts', 'attachments'].includes(key)
    )
      throw new RenderError('INVALID_OPTIONS', 'Unsupported option or accessor');
    if (key === 'attachments') {
      const map = d.value;
      if (
        !map ||
        typeof map !== 'object' ||
        types.isProxy(map) ||
        ![Object.prototype, null].includes(Object.getPrototypeOf(map)) ||
        Object.getOwnPropertySymbols(map).length
      )
        throw new RenderError('INVALID_OPTIONS', 'Expected attachment map');
      for (const [name, entry] of Object.entries(Object.getOwnPropertyDescriptors(map))) {
        if (
          !entry.enumerable ||
          !Object.hasOwn(entry, 'value') ||
          !(entry.value instanceof Uint8Array) ||
          types.isProxy(entry.value) ||
          ['__proto__', 'constructor', 'prototype'].includes(name)
        )
          throw new RenderError('INVALID_OPTIONS', 'Expected named attachment bytes');
        total += entry.value.byteLength;
        if (total > 32 * 1024 * 1024) throw new RenderError('LIMIT_EXCEEDED', 'Attachment budget exceeded');
        attachments[name] = Buffer.from(entry.value);
      }
      hasAttachments = true;
    } else json[key] = d.value;
  }
  validateData(json);
  if (
    (json.locale !== undefined && typeof json.locale !== 'string') ||
    (json.timeZone !== undefined && typeof json.timeZone !== 'string') ||
    (json.restartNumbering !== undefined && typeof json.restartNumbering !== 'boolean')
  )
    throw new RenderError('INVALID_OPTIONS', 'Invalid scalar render option');
  const copy = structuredClone(json) as RenderOptions;
  return hasAttachments ? { ...copy, attachments } : copy;
}
export function cloneAssuranceJson<T>(value: T): T {
  validateData({ value });
  return structuredClone(value);
}
