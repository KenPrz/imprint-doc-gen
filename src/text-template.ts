import { renderContext, expressionSettings } from './expressions.ts';
import type { RenderOptions } from './types.ts';
import { RenderError } from './errors.ts';
import { validateData } from './data.ts';
import { standardBindings, standardLookup } from './standard-syntax.ts';
import { scalar } from './scope.ts';
import { limits } from './package.ts';
import type { CompiledTemplate } from './types.ts';
export const textMagic = 'imprint-text-v1\n';
export function encodeTextTemplate(format: 'txt' | 'csv', source: string): Buffer {
  return Buffer.from(textMagic + JSON.stringify({ format, source }));
}
function records(source: string): string[][] {
  const rows: string[][] = [];
  let row: string[] = [],
    field = '',
    quoted = false,
    closed = false;
  for (let i = 0; i < source.length; i++) {
    const c = source[i]!;
    if (quoted) {
      if (c === '"') {
        if (source[i + 1] === '"') {
          field += '"';
          i++;
        } else {
          quoted = false;
          closed = true;
        }
      } else field += c;
    } else if (c === ',' || c === '\r' || c === '\n') {
      row.push(field);
      field = '';
      closed = false;
      if (c !== ',') {
        rows.push(row);
        row = [];
        if (c === '\r' && source[i + 1] === '\n') i++;
      }
    } else if (c === '"' && !field && !closed) quoted = true;
    else {
      if (closed || c === '"') throw new RenderError('TEMPLATE_SYNTAX', 'Invalid CSV quoting');
      field += c;
    }
  }
  if (quoted) throw new RenderError('TEMPLATE_SYNTAX', 'Unclosed CSV quote');
  if (field || row.length || closed) {
    row.push(field);
    rows.push(row);
  }
  return rows;
}
export function compileText(format: 'txt' | 'csv', source: string): CompiledTemplate {
  if (!['txt', 'csv'].includes(format) || typeof source !== 'string')
    throw new RenderError('UNSUPPORTED_FORMAT', 'Expected txt or csv template');
  if (Buffer.byteLength(source) > limits.entry)
    throw new RenderError('LIMIT_EXCEEDED', 'Text template exceeds byte budget');
  const rows = format === 'csv' ? records(source) : source.split(/\r?\n/).map((line) => [line]);
  const plans = rows.map((row) => row.map((cell) => ({ text: cell, tags: standardBindings(cell, 'text') })));
  return Object.freeze({
    format,
    async render(data: Record<string, unknown>, options: RenderOptions = {}) {
      if (options.charts || options.restartNumbering)
        throw new RenderError('UNSUPPORTED_FORMAT', 'Office options do not apply to text templates');
      return renderContext.run(expressionSettings(options.locale, options.timeZone), async () => {
        validateData(data);
        const out: string[] = [];
        let bytes = 0,
          iterations = 0;
        const emit = (row: (typeof plans)[number], item?: Record<string, unknown>, path?: string) => {
          const values = row.map((cell) => {
            let result = '',
              cursor = 0;
            for (const tag of cell.tags) {
              const local = path && tag.path.startsWith(path + '[i].');
              const value = standardLookup(
                local ? [data, item!] : [data],
                local ? tag.path.slice(path!.length + 4) : tag.path,
                'text',
                tag.formatters,
              );
              const replacement = scalar(value, tag.path, 'text', true);
              if (Buffer.byteLength(result) + Buffer.byteLength(replacement) + bytes > limits.entry)
                throw new RenderError('LIMIT_EXCEEDED', 'Text replacement exceeds output budget');
              result += cell.text.slice(cursor, tag.start) + replacement;
              cursor = tag.end;
            }
            result += cell.text.slice(cursor);
            if (format === 'csv') {
              if (/^[\s\u0000-\u001f]*[=+@-]/.test(result)) result = "'" + result;
              return /[",\r\n]/.test(result) ? '"' + result.replaceAll('"', '""') + '"' : result;
            }
            return result;
          });
          const line = values.join(format === 'csv' ? ',' : '');
          bytes += Buffer.byteLength(line) + 2;
          if (bytes > limits.entry)
            throw new RenderError('LIMIT_EXCEEDED', 'Text output exceeds byte budget');
          out.push(line);
        };
        for (let n = 0; n < plans.length; n++) {
          const row = plans[n]!,
            tags = row.flatMap((c) => c.tags),
            loop = tags.find((t) => t.path.includes('[i]'));
          if (!loop) {
            if (tags.some((t) => t.path.includes('[i')))
              throw new RenderError('TEMPLATE_STRUCTURE', 'Unmatched text increment marker');
            emit(row);
            continue;
          }
          const path = loop.path.split('[i]')[0]!;
          if (
            tags.some((t) => t.path.includes('[i') && !t.path.startsWith(path + '[i].')) ||
            tags.some((t) => t.filters?.length)
          )
            throw new RenderError(
              'UNSUPPORTED_TEMPLATE',
              'Text repeats support one unfiltered array per record',
            );
          const marker = plans[++n];
          if (
            !marker ||
            marker.flatMap((c) => c.tags).length !== 1 ||
            marker.filter((c) => c.text.trim()).length !== 1 ||
            marker.find((c) => c.tags.length)?.text.trim() !== `{d.${path}[i+1]}`
          )
            throw new RenderError(
              'TEMPLATE_STRUCTURE',
              'Repeat record must be followed by a standalone {d.array[i+1]} marker',
            );
          const items = standardLookup([data], path, 'text') ?? [];
          if (!Array.isArray(items)) throw new RenderError('INVALID_DATA', 'Expected repeat array');
          for (const item of items) {
            if (++iterations > 10000)
              throw new RenderError('LIMIT_EXCEEDED', 'Text repetition exceeds 10000 items');
            if (!item || typeof item !== 'object' || Array.isArray(item))
              throw new RenderError('INVALID_DATA', 'Expected object items');
            emit(row, item, path);
          }
        }
        return Buffer.from(out.join(format === 'csv' ? '\r\n' : '\n'));
      });
    },
  });
}
