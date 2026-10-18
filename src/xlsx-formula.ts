import { RenderError } from './errors.ts';
import { a1, address, columnName, columnNumber, maxRow, RowMap, type Address } from './xlsx-references.ts';

interface Reference { kind: 'reference'; start: number; end: number; sheet?: string; qualifier: string; first: Address; last?: Address }
type Expr = { kind: 'literal'; value: Value } | Reference | { kind: 'unary'; op: string; child: Expr } | { kind: 'binary'; op: string; left: Expr; right: Expr } | { kind: 'call'; name: string; args: Expr[] };
interface Token { kind: 'number' | 'string' | 'name' | 'operator' | 'error'; text: string; start: number; end: number }
export interface Formula { source: string; refs: Reference[]; expression: Expr }
export interface ExcelError { error: string }
export type Value = number | string | boolean | null | ExcelError;
const error = (code: string): ExcelError => ({ error: code });
const isError = (value: Value): value is ExcelError => typeof value === 'object' && value !== null;
const functions = new Set(['SUM', 'MIN', 'MAX', 'AVERAGE', 'COUNT', 'IF', 'IFERROR', 'ABS', 'ROUND', 'ROUNDUP', 'ROUNDDOWN', 'LEN', 'LOWER', 'UPPER', 'CONCAT', 'AND', 'OR', 'NOT', 'LEFT', 'RIGHT', 'MID', 'TRIM', 'INT', 'MOD', 'SQRT', 'POWER', 'COUNTA', 'COUNTBLANK']);
export function parseFormula(source: string, allowRangeRoot = false): Formula {
  if (source.length > 8192) throw new RenderError('LIMIT_EXCEEDED', 'Formula exceeds 8192 characters');
  const tokens: (Token | Reference)[] = [], refs: Reference[] = [];
  let cursor = source.startsWith('=') ? 1 : 0;
  while (cursor < source.length) {
    const tail = source.slice(cursor), space = /^\s+/.exec(tail);
    if (space) { cursor += space[0].length; continue; }
    const ref = /^(?:(?:'((?:[^']|'')+)'|([\p{L}_\\][\p{L}\p{N}_.\\]*))!)?(\$?[A-Za-z]{1,3}\$?[1-9]\d*)(?::(\$?[A-Za-z]{1,3}\$?[1-9]\d*))?(?![\p{L}\p{N}_.$!(])/u.exec(tail);
    if (ref) {
      const qualifier = ref[1] !== undefined || ref[2] !== undefined ? ref[0].slice(0, ref[0].length - ref[3]!.length - (ref[4] ? ref[4].length + 1 : 0)) : '';
      const item: Reference = { kind: 'reference', start: cursor, end: cursor + ref[0].length, sheet: ref[1]?.replaceAll("''", "'") ?? ref[2], qualifier, first: address(ref[3]!), last: ref[4] ? address(ref[4]) : undefined };
      tokens.push(item); refs.push(item); cursor = item.end;
    } else {
      const match = /^(?:\d+(?:\.\d*)?|\.\d+)(?:[Ee][+-]?\d+)?/.exec(tail) ?? /^"(?:[^"]|"")*"/.exec(tail) ?? /^#[A-Z0-9/]+[!?]?/.exec(tail) ?? /^[A-Za-z_][A-Za-z_0-9.]*/.exec(tail) ?? /^(?:<=|>=|<>|[+\-*/^%&=<>(),])/.exec(tail);
      if (!match) throw new RenderError('UNSUPPORTED_FORMULA', `Unsupported formula syntax near ${tail.slice(0, 24)}`);
      const text = match[0], kind = /^[\d.]/.test(text) ? 'number' : text.startsWith('"') ? 'string' : text.startsWith('#') ? 'error' : /^[A-Za-z_]/.test(text) ? 'name' : 'operator';
      tokens.push({ kind, text, start: cursor, end: cursor + text.length }); cursor += text.length;
    }
    if (tokens.length > 4096) throw new RenderError('LIMIT_EXCEEDED', 'Formula exceeds its token budget');
  }
  let position = 0, depth = 0;
  const peek = () => { const token = tokens[position]; return token?.kind === 'operator' ? token.text : undefined; };
  const take = (value: string) => { if (peek() !== value) throw new RenderError('UNSUPPORTED_FORMULA', `Expected ${value} in formula`); position++; };
  const precedence: Record<string, number> = { '=': 1, '<>': 1, '<': 1, '>': 1, '<=': 1, '>=': 1, '&': 2, '+': 3, '-': 3, '*': 4, '/': 4, '^': 5 };
  const expression = (minimum = 0): Expr => {
    if (++depth > 128) throw new RenderError('LIMIT_EXCEEDED', 'Formula nesting exceeds 128 levels');
    let left: Expr;
    const token = tokens[position++];
    if (!token) throw new RenderError('UNSUPPORTED_FORMULA', 'Incomplete formula');
    if (token.kind === 'reference') left = token;
    else if (token.kind === 'number') {
      if (!Number.isFinite(Number(token.text))) throw new RenderError('UNSUPPORTED_FORMULA', 'Numeric formula literal is outside the supported finite range');
      left = { kind: 'literal', value: Number(token.text) };
    }
    else if (token.kind === 'string') left = { kind: 'literal', value: token.text.slice(1, -1).replaceAll('""', '"') };
    else if (token.kind === 'error') left = { kind: 'literal', value: error(token.text) };
    else if (token.kind === 'operator' && ['+', '-'].includes(token.text)) left = { kind: 'unary', op: token.text, child: expression(6) };
    else if (token.kind === 'operator' && token.text === '(') { left = expression(); take(')'); }
    else if (token.kind === 'name') {
      const name = token.text.toUpperCase();
      if (['TRUE', 'FALSE'].includes(name) && peek() !== '(') left = { kind: 'literal', value: name === 'TRUE' };
      else {
        if (!functions.has(name)) throw new RenderError('UNSUPPORTED_FORMULA', `Unsupported formula function or named reference: ${token.text}`);
        take('('); const args: Expr[] = [];
        if (peek() !== ')') { args.push(expression()); while (peek() === ',') { position++; args.push(expression()); } }
        take(')');
        if (name === 'IF' && (args.length < 2 || args.length > 3)) throw new RenderError('UNSUPPORTED_FORMULA', 'IF requires two or three arguments');
        const arity: Record<string, number> = { IFERROR: 2, ABS: 1, ROUND: 2, ROUNDUP: 2, ROUNDDOWN: 2, LEN: 1, LOWER: 1, UPPER: 1, NOT: 1, LEFT: 2, RIGHT: 2, MID: 3, TRIM: 1, INT: 1, MOD: 2, SQRT: 1, POWER: 2, COUNTBLANK: 1 };
        if (arity[name] !== undefined && args.length !== arity[name]) throw new RenderError('UNSUPPORTED_FORMULA', `${name} requires ${arity[name]} arguments`);
        left = { kind: 'call', name, args };
      }
    } else throw new RenderError('UNSUPPORTED_FORMULA', 'Unsupported formula expression');
    while (peek() === '%') { position++; left = { kind: 'unary', op: '%', child: left }; }
    while (peek() !== undefined && (precedence[peek()!] ?? -1) >= minimum) {
      const op = peek()!; position++;
      left = { kind: 'binary', op, left, right: expression(precedence[op]! + 1) };
    }
    depth--; return left;
  };
  const tree = expression();
  if (position !== tokens.length) throw new RenderError('UNSUPPORTED_FORMULA', 'Unsupported trailing formula syntax');
  const pending: [Expr, number, boolean][] = [[tree, 1, allowRangeRoot]];
  while (pending.length) {
    const [node, depth, rangeAllowed] = pending.pop()!;
    if (depth > 128) throw new RenderError('LIMIT_EXCEEDED', 'Formula expression depth exceeds 128 levels');
    if (node.kind === 'reference' && node.last) {
      if (!rangeAllowed) throw new RenderError('UNSUPPORTED_FORMULA', 'Array-valued expressions require a later calculator adapter');
      if (node.first.row > node.last.row || columnNumber(node.first.column) > columnNumber(node.last.column)) throw new RenderError('UNSUPPORTED_FORMULA', 'Reversed formula ranges require a later adapter');
    }
    if (node.kind === 'binary') pending.push([node.left, depth + 1, false], [node.right, depth + 1, false]);
    else if (node.kind === 'unary') pending.push([node.child, depth + 1, false]);
    else if (node.kind === 'call') for (const argument of node.args) pending.push([argument, depth + 1, ['SUM','MIN','MAX','AVERAGE','COUNT','CONCAT','AND','OR','COUNTA','COUNTBLANK'].includes(node.name)]);
  }
  return { source, refs, expression: tree };
}
export function rewriteFormula(formula: Formula, maps: Map<string, RowMap>, sheet: string, copy = 0, loopRow?: number): string {
  let result = '', cursor = 0;
  for (const ref of formula.refs) {
    const target = ref.sheet ?? sheet, map = maps.get(target.toLowerCase());
    if (!map) throw new RenderError('UNSUPPORTED_FORMULA', `Unknown worksheet: ${target}`);
    const translate = (value: Address, edge: 'cell' | 'start' | 'end') => {
      const sameBlock = loopRow !== undefined && map.expansions.some(e => loopRow >= e.row && loopRow < e.row + (e.height ?? 1) && value.row >= e.row && value.row < e.row + (e.height ?? 1));
      const own = (loopRow === value.row || sameBlock) && target.toLowerCase() === sheet.toLowerCase() && !value.absoluteRow;
      const row = map.map(value.row, own ? 'cell' : edge);
      if (row === undefined) return undefined;
      const shifted = row + (value.absoluteRow ? 0 : copy);
      return shifted >= 1 && shifted <= maxRow ? a1({ ...value, row: shifted }) : undefined;
    };
    const first = translate(ref.first, ref.last ? 'start' : 'cell'), last = ref.last ? translate(ref.last, 'end') : undefined;
    const deleted = ref.last && first && last && address(first).row > address(last).row;
    result += formula.source.slice(cursor, ref.start) + (!first || (ref.last && !last) || deleted ? '#REF!' : ref.qualifier + first + (last ? ':' + last : ''));
    cursor = ref.end;
  }
  return result + formula.source.slice(cursor);
}
export interface CalculationBudget { operations: number }
export function evaluateFormula(formula: Formula, sheet: string, read: (sheet: string, reference: string) => Value, budget: CalculationBudget): Value {
  const tick = () => { if (++budget.operations > 1_000_000) throw new RenderError('LIMIT_EXCEEDED', 'Formula calculation exceeds its operation budget'); };
  const numeric = (value: Value): number | ExcelError => {
    if (isError(value)) return value;
    if (value === null) return 0;
    if (typeof value === 'number') return value;
    if (typeof value === 'boolean') return Number(value);
    return /^[+-]?(?:\d+(?:\.\d*)?|\.\d+)(?:[eE][+-]?\d+)?$/.test(value.trim()) && Number.isFinite(Number(value)) ? Number(value) : error('#VALUE!');
  };
  const checked = (n: number): Value => Number.isFinite(n) ? n : error('#NUM!');
  const visit = (node: Expr): Value => {
    tick();
    if (node.kind === 'literal') return typeof node.value === 'number' ? checked(node.value) : node.value;
    if (node.kind === 'reference') return node.last ? error('#VALUE!') : read(node.sheet ?? sheet, node.first.column + node.first.row);
    if (node.kind === 'unary') { const n = numeric(visit(node.child)); return typeof n === 'number' ? checked(node.op === '-' ? -n : node.op === '%' ? n / 100 : n) : n; }
    if (node.kind === 'binary') {
      const left = visit(node.left), right = visit(node.right);
      if (isError(left)) return left; if (isError(right)) return right;
      if (node.op === '&') { if (typeof left === 'boolean' || typeof right === 'boolean') throw new RenderError('UNSUPPORTED_FORMULA', 'Boolean concatenation differs across spreadsheet applications'); const text = String(left ?? '') + String(right ?? ''); return text.length <= 32767 ? text : error('#VALUE!'); }
      if (['=', '<>', '<', '>', '<=', '>='].includes(node.op)) {
        // Excel orders numbers before text before booleans; text comparison ignores case.
        const normalize = (v: Value, other: Value) => typeof v === 'string' ? v.toUpperCase() : v ?? (typeof other === 'string' ? '' : typeof other === 'boolean' ? false : 0);
        const a = normalize(left, right) as string | number | boolean, b = normalize(right, left) as string | number | boolean;
        const rank = (v: unknown) => typeof v === 'number' ? 0 : typeof v === 'string' ? 1 : 2;
        const c = rank(a) !== rank(b) ? rank(a) - rank(b) : a === b ? 0 : a < b ? -1 : 1;
        return node.op === '=' ? c === 0 : node.op === '<>' ? c !== 0 : node.op === '<' ? c < 0 : node.op === '>' ? c > 0 : node.op === '<=' ? c <= 0 : c >= 0;
      }
      const a = numeric(left), b = numeric(right); if (typeof a !== 'number') return a; if (typeof b !== 'number') return b;
      if (node.op === '^' && a === 0 && b <= 0) throw new RenderError('UNSUPPORTED_FORMULA', 'Zero raised to a nonpositive power is outside the supported calculation domain');
      if (node.op === '/' && b === 0) return error('#DIV/0!');
      return checked(node.op === '+' ? a + b : node.op === '-' ? a - b : node.op === '*' ? a * b : node.op === '/' ? a / b : a ** b);
    }
    if (['LEFT','RIGHT','MID','TRIM'].includes(node.name)) {
      const values = node.args.map(visit); const failure = values.find(isError); if (failure) return failure;
      const v = values[0], text = typeof v === 'boolean' ? (v ? 'TRUE' : 'FALSE') : String(v ?? '');
      if (node.name === 'TRIM') return text.replace(/ +/g, ' ').replace(/^ | $/g, '');
      const n = numeric(values[1]!); if (typeof n !== 'number') return n;
      if (n < (node.name === 'MID' ? 1 : 0)) return error('#VALUE!');
      if (node.name === 'MID') { const count = numeric(values[2]!); if (typeof count !== 'number') return count; return count < 0 ? error('#VALUE!') : text.slice(Math.trunc(n)-1, Math.trunc(n)-1+Math.trunc(count)); }
      return node.name === 'LEFT' ? text.slice(0,Math.trunc(n)) : Math.trunc(n) === 0 ? '' : text.slice(-Math.trunc(n));
    }
    if (['INT','MOD','SQRT','POWER'].includes(node.name)) {
      const a = numeric(visit(node.args[0]!)); if (typeof a !== 'number') return a;
      if (node.name === 'INT') return Math.floor(a);
      if (node.name === 'SQRT') return a < 0 ? error('#NUM!') : Math.sqrt(a);
      const b = numeric(visit(node.args[1]!)); if (typeof b !== 'number') return b;
      if (node.name === 'MOD') return b === 0 ? error('#DIV/0!') : checked(a - b * Math.floor(a/b));
      if (a === 0 && b < 0) return error('#DIV/0!');
      return a === 0 && b === 0 ? error('#NUM!') : checked(a ** b);
    }
    if (['COUNTA','COUNTBLANK'].includes(node.name)) {
      let count = 0;
      const add = (value: Value) => { if (node.name === 'COUNTA' ? value !== null : value === null || value === '') count++; };
      for (const arg of node.args) {
        if (arg.kind === 'reference' && arg.last) for (let r=arg.first.row; r<=arg.last.row; r++) for (let c=columnNumber(arg.first.column); c<=columnNumber(arg.last.column); c++) { tick(); add(read(arg.sheet ?? sheet,columnName(c)+r)); }
        else add(visit(arg));
      }
      return count;
    }
    if (node.name === 'IF') {
      const value = visit(node.args[0]!);
      const condition = typeof value === 'string' && /^(TRUE|FALSE)$/i.test(value) ? Number(value.toUpperCase() === 'TRUE') : numeric(value);
      return typeof condition === 'number' ? condition ? visit(node.args[1]!) : node.args[2] ? visit(node.args[2]) : false : condition;
    }
    if (node.name === 'IFERROR') { const value = visit(node.args[0]!); return isError(value) ? visit(node.args[1]!) : value; }
    if (['ABS','ROUND','ROUNDUP','ROUNDDOWN','NOT'].includes(node.name)) {
      const value = numeric(visit(node.args[0]!)); if (typeof value !== 'number') return value;
      if (node.name === 'ABS') return Math.abs(value); if (node.name === 'NOT') return !value;
      const precision = numeric(visit(node.args[1]!)); if (typeof precision !== 'number') return precision;
      const digits = Math.trunc(precision); if (Math.abs(digits) > 15) throw new RenderError('UNSUPPORTED_FORMULA', 'Rounding precision is limited to 15 decimal places');
      const shift = (n: number, d: number) => { const [base, exponent='0'] = String(n).split('e'); return Number(`${base}e${Number(exponent)+d}`); };
      const scaled = shift(Math.abs(value), digits), rounded = node.name === 'ROUNDUP' ? Math.ceil(scaled) : node.name === 'ROUNDDOWN' ? Math.floor(scaled) : Math.round(scaled);
      return checked(Math.sign(value) * shift(rounded, -digits));
    }
    if (['LEN','LOWER','UPPER'].includes(node.name)) {
      const value = visit(node.args[0]!); if (isError(value)) return value;
      const text = typeof value === 'boolean' ? value ? 'TRUE' : 'FALSE' : String(value ?? '');
      return node.name === 'LEN' ? text.length : node.name === 'LOWER' ? text.toLowerCase() : text.toUpperCase();
    }
    if (['CONCAT','AND','OR'].includes(node.name)) {
      let text = '', logical = node.name === 'AND', count = 0;
      const accept = (value: Value, reference: boolean): ExcelError | undefined => {
        if (isError(value)) return value;
        if (node.name === 'CONCAT') { text += typeof value === 'boolean' ? value ? 'TRUE' : 'FALSE' : String(value ?? ''); if (text.length > 32767) return error('#VALUE!'); }
        else { if (reference && (value === null || typeof value === 'string')) return; const n = numeric(value); if (typeof n !== 'number') return n; count++; logical = node.name === 'AND' ? logical && !!n : logical || !!n; }
      };
      for (const arg of node.args) {
        if (arg.kind === 'reference' && arg.last) for (let row=arg.first.row; row<=arg.last.row; row++) for (let col=columnNumber(arg.first.column); col<=columnNumber(arg.last.column); col++) { tick(); const failure = accept(read(arg.sheet ?? sheet,columnName(col)+row),true); if (failure) return failure; }
        else { const failure = accept(visit(arg),arg.kind==='reference'); if (failure) return failure; }
      }
      return node.name === 'CONCAT' ? text : count ? logical : error('#VALUE!');
    }
    const values: number[] = [];
    const add = (value: Value, reference: boolean): ExcelError | undefined => {
      if (isError(value)) return node.name === 'COUNT' ? undefined : value;
      if (reference && typeof value !== 'number') return;
      const n = numeric(value); if (typeof n === 'number') values.push(n); else if (node.name !== 'COUNT') return n;
    };
    for (const arg of node.args) {
      if (arg.kind === 'reference' && arg.last) {
        const first = arg.first, last = arg.last;
        if (first.row > last.row || columnNumber(first.column) > columnNumber(last.column)) return error('#REF!');
        for (let row = first.row; row <= last.row; row++) for (let col = columnNumber(first.column); col <= columnNumber(last.column); col++) {
          tick(); const failure = add(read(arg.sheet ?? sheet, columnName(col) + row), true); if (failure) return failure;
        }
      } else { const failure = add(visit(arg), arg.kind === 'reference'); if (failure) return failure; }
    }
    if (node.name === 'COUNT') return values.length;
    if (node.name === 'AVERAGE' && !values.length) return error('#DIV/0!');
    if (!values.length) return 0;
    if (node.name === 'MIN' || node.name === 'MAX') return values.reduce((a, b) => node.name === 'MIN' ? Math.min(a, b) : Math.max(a, b));
    const sum = values.reduce((a, b) => a + b, 0);
    return checked(node.name === 'AVERAGE' ? sum / values.length : sum);
  };
  return visit(formula.expression);
}
