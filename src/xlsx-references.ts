import { RenderError } from './errors.ts';

export const maxRow = 1_048_576, maxColumn = 16_384;
export interface Address { column: string; row: number; absoluteColumn: boolean; absoluteRow: boolean }
export function address(value: string): Address {
  const match = /^(\$?)([A-Z]{1,3})(\$?)([1-9]\d*)$/i.exec(value);
  if (!match) throw new RenderError('UNSUPPORTED_FORMULA', `Expected an A1 cell reference: ${value}`);
  const result = { column: match[2]!.toUpperCase(), row: Number(match[4]), absoluteColumn: !!match[1], absoluteRow: !!match[3] };
  if (columnNumber(result.column) > maxColumn || result.row > maxRow) throw new RenderError('LIMIT_EXCEEDED', 'Spreadsheet reference exceeds Excel bounds');
  return result;
}
export const columnNumber = (column: string) => [...column.toUpperCase()].reduce((n, c) => n * 26 + c.charCodeAt(0) - 64, 0);
export function columnName(number: number): string { let result = ''; while (number) { number--; result = String.fromCharCode(65 + number % 26) + result; number = Math.floor(number / 26); } return result; }
export const a1 = (ref: Address) => `${ref.absoluteColumn ? '$' : ''}${ref.column}${ref.absoluteRow ? '$' : ''}${ref.row}`;
export interface RowExpansion { row: number; count: number }
export class RowMap {
  readonly expansions: readonly RowExpansion[];
  constructor(expansions: readonly RowExpansion[]) { this.expansions = [...expansions].sort((a, b) => a.row - b.row); }
  map(row: number, edge: 'cell' | 'start' | 'end' = 'cell'): number | undefined {
    let delta = 0;
    for (const expansion of this.expansions) {
      if (expansion.row > row) break;
      if (expansion.row === row) {
        if (!expansion.count && edge === 'cell') return undefined;
        const result = row + delta + (edge === 'end' ? expansion.count - 1 : 0);
        if (result > maxRow) throw new RenderError('LIMIT_EXCEEDED', 'Expanded worksheet exceeds Excel row bounds');
        if (result < 1) return undefined;
        return result;
      }
      delta += expansion.count - 1;
    }
    const result = row + delta;
    if (result < 1 || result > maxRow) throw new RenderError('LIMIT_EXCEEDED', 'Expanded worksheet exceeds Excel row bounds');
    return result;
  }
  intersects(start: number, end = start) { return this.expansions.some(e => e.row >= start && e.row <= end); }
  range(value: string): string | undefined {
    const bounds = value.split(':');
    if (bounds.length > 2) throw new RenderError('UNSUPPORTED_TEMPLATE', `Unsupported worksheet range: ${value}`);
    const first = address(bounds[0]!), last = address(bounds[1] ?? bounds[0]!);
    if (last.row < first.row || columnNumber(last.column) < columnNumber(first.column)) throw new RenderError('INVALID_PACKAGE', `Reversed worksheet range: ${value}`);
    const start = this.map(first.row, 'start'), end = this.map(last.row, 'end');
    if (start === undefined || end === undefined || start > end) return undefined;
    first.row = start; last.row = end;
    return bounds.length === 1 && start === end ? a1(first) : `${a1(first)}:${a1(last)}`;
  }
}
