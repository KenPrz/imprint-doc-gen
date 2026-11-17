import { RenderError } from './errors.ts';
import type { PackageParts } from './package.ts';
import { forbidden } from './scope.ts';
import { descendants, parseXml } from './xml.ts';

/** Fixed-row form pagination: render the template once per chunk of `array`, each padded to `rowsPerSheet`. */
export interface SheetRule { array: string; rowsPerSheet: number; pad: boolean }
export interface SheetsInput extends Partial<SheetRule> { perSheet?: Record<string, unknown>[] }

const invalid = (message: string): never => { throw new RenderError('INVALID_SHEETS', message); };
const segment = /^[A-Za-z_$][\w$]*$/;
function checkArray(path: string) {
  if (!path.split('.').every(name => segment.test(name) && !forbidden.has(name))) invalid(`Invalid sheet array path "${path}"`);
}
function checkRows(rows: number) {
  if (!Number.isSafeInteger(rows) || rows < 1 || rows > 500) invalid('rowsPerSheet must be an integer from 1 to 500');
}

/** Parses the `imprint.sheets` template property: `<array path>:<rows per sheet>[:pad|nopad]`. */
export function parseSheetRule(text: string): SheetRule {
  const [array = '', rows = '', pad = 'pad', ...rest] = text.split(':').map(field => field.trim());
  if (rest.length || !/^\d+$/.test(rows) || !['pad', 'nopad'].includes(pad)) invalid(`Invalid imprint.sheets value "${text}"; expected <array>:<rows>[:pad|nopad]`);
  checkArray(array); checkRows(Number(rows));
  return { array, rowsPerSheet: Number(rows), pad: pad === 'pad' };
}

/** Reads `imprint.sheets` from OOXML custom properties or ODF user-defined metadata. */
export function templateSheetRule(parts: PackageParts): SheetRule | null {
  const custom = parts.get('docProps/custom.xml'), meta = parts.get('meta.xml');
  const value = custom
    ? descendants(parseXml(custom.toString(), 'docProps/custom.xml'), 'http://schemas.openxmlformats.org/officeDocument/2006/custom-properties', 'property').find(p => p.getAttribute('name') === 'imprint.sheets')?.textContent
    : meta
      ? descendants(parseXml(meta.toString(), 'meta.xml'), 'urn:oasis:names:tc:opendocument:xmlns:meta:1.0', 'user-defined').find(p => p.getAttributeNS('urn:oasis:names:tc:opendocument:xmlns:meta:1.0', 'name') === 'imprint.sheets')?.textContent
      : undefined;
  return value == null ? null : parseSheetRule(value);
}

/** Request fields override the template rule field by field. Returns null when neither declares pagination. */
export function resolveSheetRule(template: SheetRule | null, request?: SheetsInput): SheetRule | null {
  if (!template && !request) return null;
  const rule = { ...template, ...Object.fromEntries(Object.entries(request ?? {}).filter(([key, value]) => key !== 'perSheet' && value !== undefined)) };
  if (rule.array === undefined || rule.rowsPerSheet === undefined) invalid('Sheets need array and rowsPerSheet, from the request or the imprint.sheets template property');
  checkArray(rule.array!); checkRows(rule.rowsPerSheet!);
  return { array: rule.array!, rowsPerSheet: rule.rowsPerSheet!, pad: rule.pad ?? true };
}

const isRecord = (value: unknown): value is Record<string, unknown> => typeof value === 'object' && value !== null && !Array.isArray(value);

/** One data object per sheet: the array replaced by its chunk, plus `sheet.{number,count,isFirst,isLast,notLast}` and perSheet keys. */
export function sheetData(data: Record<string, unknown>, rule: SheetRule, perSheet: Record<string, unknown>[] | undefined, maxSheets: number): Record<string, unknown>[] {
  const names = rule.array.split('.');
  let source: unknown = data;
  for (const name of names) source = isRecord(source) && Object.hasOwn(source, name) ? source[name] : undefined;
  const rows = source ?? [];
  if (!Array.isArray(rows) || !rows.every(isRecord)) invalid(`d.${rule.array} must be an array of objects`);
  const count = Math.max(1, Math.ceil((rows as unknown[]).length / rule.rowsPerSheet));
  if (count > maxSheets) throw new RenderError('LIMIT_EXCEEDED', `Document needs ${count} sheets; the limit is ${maxSheets}`);
  if (perSheet?.length && perSheet.length !== count) throw new RenderError('SHEET_DATA_MISMATCH', `perSheet has ${perSheet.length} entries; the document has ${count} sheets`);
  return Array.from({ length: count }, (_, index) => {
    const chunk: unknown[] = (rows as unknown[]).slice(index * rule.rowsPerSheet, (index + 1) * rule.rowsPerSheet);
    if (rule.pad) while (chunk.length < rule.rowsPerSheet) chunk.push({});
    const sheet = { ...perSheet?.[index], number: index + 1, count, isFirst: index === 0 ? '1' : '', isLast: index === count - 1 ? '1' : '', notLast: index === count - 1 ? '' : '1' };
    // Copy only the objects along the array path; everything else is shared read-only data.
    const root: Record<string, unknown> = { ...data, sheet };
    let target = root;
    for (const name of names.slice(0, -1)) target = target[name] = isRecord(target[name]) ? { ...target[name] } : {};
    target[names.at(-1)!] = chunk;
    return root;
  });
}
