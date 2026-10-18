import type { Document, Element } from '@xmldom/xmldom';
import { descendants } from './xml.ts';
import { formatNs as f } from './format-namespaces.ts';
import { RenderError } from './errors.ts';
export function odsFormulaRewriter(doc: Document) {
  const groups = new WeakMap<Element, Element>();
  const origins = new WeakMap<Element, number>();
  const originalCounts = new WeakMap<Element, number>();
  const tables = descendants(doc, f.table, 'table');
  for (const table of tables) {
    const rows = descendants(table, f.table, 'table-row');
    originalCounts.set(table, rows.length);
    rows.forEach((row, n) => origins.set(row, n + 1));
  }
  return {
    clone(original: Element, clone: Element, group: Element) {
      const a = [original, ...descendants(original, '*', '*')],
        b = [clone, ...descendants(clone, '*', '*')];
      a.forEach((node, n) => {
        const origin = origins.get(node);
        if (origin) {
          origins.set(b[n]!, origin);
          groups.set(b[n]!, group);
        }
      });
    },
    finish() {
      for (const table of tables) {
        const rows = descendants(table, f.table, 'table-row');
        if (rows.length === originalCounts.get(table) && rows.every((row, n) => origins.get(row) === n + 1))
          continue;
        if (!descendants(table, f.table, 'table-cell').some((c) => c.hasAttributeNS(f.table, 'formula')))
          continue;
        if (
          rows.some((r) => r.hasAttributeNS(f.table, 'number-rows-repeated')) ||
          descendants(table, f.table, 'table-cell').some((c) =>
            c.hasAttributeNS(f.table, 'number-columns-repeated'),
          )
        )
          throw new RenderError(
            'UNSUPPORTED_FORMULA',
            'Formula rewriting requires explicit ODS rows and cells',
          );
        const mapped = new Map<number, number[]>();
        rows.forEach((row, n) => {
          const old = origins.get(row)!;
          const list = mapped.get(old) ?? [];
          list.push(n + 1);
          mapped.set(old, list);
        });
        const location = (r: number) =>
          r > originalCounts.get(table)! ? [r + rows.length - originalCounts.get(table)!] : mapped.get(r);
        rows.forEach((row, n) => {
          const origin = origins.get(row)!;
          const copy = n + 1 - mapped.get(origin)![0]!;
          for (const cell of descendants(row, f.table, 'table-cell')) {
            const formula = cell.getAttributeNS(f.table, 'formula');
            if (!formula) continue;
            if (!formula.startsWith('of:='))
              throw new RenderError('UNSUPPORTED_FORMULA', 'Expected OpenFormula of:= formula');
            let result = '',
              cursor = 0;
            // Strings are copied verbatim; brackets in strings are never interpreted as references.
            const token = /"(?:[^"]|"")*"|\[[^\]]*\]/g;
            for (const m of formula.matchAll(token)) {
              result += formula.slice(cursor, m.index);
              cursor = m.index + m[0].length;
              if (m[0].startsWith('"')) {
                result += m[0];
                continue;
              }
              const ref = /^\[\.(\$?[A-Z]+)(\$?)([1-9]\d*)(?::\.(\$?[A-Z]+)(\$?)([1-9]\d*))?\]$/i.exec(m[0]);
              if (!ref)
                throw new RenderError('UNSUPPORTED_FORMULA', 'ODS formulas support local A1 references only');
              const first = location(Number(ref[3])),
                last = ref[6] ? location(Number(ref[6])) : undefined;
              if (!first || (ref[6] && !last)) {
                result += '#REF!';
                continue;
              }
              const targetRow = rows.find((r) => origins.get(r) === Number(ref[6]));
              const own =
                groups.has(row) && targetRow && groups.get(row) === groups.get(targetRow) && !ref[5];
              const start = first[0]! + (ref[2] ? 0 : copy),
                end = last ? (own ? last[0]! : last.at(-1)!) + (ref[5] ? 0 : copy) : undefined;
              if (start < 1 || start > 1048576 || (end !== undefined && (end < start || end > 1048576)))
                throw new RenderError('LIMIT_EXCEEDED', 'ODS reference exceeds row bounds');
              result += `[.${ref[1]}${ref[2]}${start}${end === undefined ? '' : `:.${ref[4]}${ref[5]}${end}`}]`;
            }
            result += formula.slice(cursor);
            if (
              result
                .replace(/"(?:[^"]|"")*"/g, '')
                .replace(/\[[^\]]*\]/g, '')
                .match(/[\[\]]/)
            )
              throw new RenderError('UNSUPPORTED_FORMULA', 'Malformed ODS reference');
            cell.setAttributeNS(f.table, 'table:formula', result);
            for (const name of ['value', 'boolean-value', 'date-value', 'string-value', 'value-type'])
              cell.removeAttributeNS(f.office, name);
            for (const p of descendants(cell, f.text, 'p')) p.parentNode!.removeChild(p);
          }
        });
      }
    },
  };
}
