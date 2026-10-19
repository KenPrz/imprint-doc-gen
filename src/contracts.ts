import { types as utilTypes } from 'node:util';
/** A deliberately bounded JSON Schema subset. Unknown keywords fail closed. */
export interface ContractSchema {
  type: 'object' | 'array' | 'string' | 'number' | 'integer' | 'boolean' | 'null';
  properties?: Record<string, ContractSchema>;
  required?: string[];
  additionalProperties?: boolean;
  items?: ContractSchema;
  enum?: unknown[];
  minimum?: number;
  maximum?: number;
  minLength?: number;
  maxLength?: number;
  minItems?: number;
  maxItems?: number;
  description?: string;
}
export interface ContractIssue {
  path: string;
  code: string;
  message: string;
}
export interface BusinessAssertion {
  name: string;
  left: { path: string } | { sum: string; field?: string };
  operator: 'equals' | 'lessThanOrEqual' | 'greaterThanOrEqual';
  right: { path: string } | { value: string | number | boolean | null };
  tolerance?: number;
}
const types = new Set(['object', 'array', 'string', 'number', 'integer', 'boolean', 'null']);
const common = ['type', 'enum', 'description'];
const own = (v: object, k: string) => Object.hasOwn(v, k);
const pointer = (p: string, k: string | number) =>
  `${p}/${String(k).replace(/~/g, '~0').replace(/\//g, '~1')}`;
function json(value: unknown): void {
  let nodes = 0,
    characters = 0;
  const seen = new Set<object>();
  function visit(v: unknown, depth: number): void {
    if (++nodes > 100_000 || depth > 64) throw new Error('Contract JSON budget exceeded');
    if (typeof v === 'string') {
      characters += v.length;
      if (characters > 10_000_000) throw new Error('Contract string budget exceeded');
      return;
    }
    if (v === null || typeof v === 'boolean') return;
    if (typeof v === 'number' && Number.isFinite(v)) return;
    if (typeof v !== 'object' || !v || seen.has(v)) throw new Error('Expected finite acyclic JSON data');
    if (utilTypes.isProxy(v)) throw new Error('JSON proxies are unsupported');
    if (
      !Array.isArray(v) &&
      Object.getPrototypeOf(v) !== Object.prototype &&
      Object.getPrototypeOf(v) !== null
    )
      throw new Error('Expected plain JSON object');
    if (Object.getOwnPropertySymbols(v).length) throw new Error('Symbol JSON keys are unsupported');
    if (
      Object.getOwnPropertyNames(v).some(
        (k) => !(Array.isArray(v) && k === 'length') && !Object.getOwnPropertyDescriptor(v, k)!.enumerable,
      )
    )
      throw new Error('Non-enumerable JSON properties are unsupported');
    seen.add(v);
    if (Array.isArray(v) && Object.keys(v).length !== v.length)
      throw new Error('Sparse or extended arrays are unsupported');
    for (const key of Object.keys(v)) {
      const d = Object.getOwnPropertyDescriptor(v, key)!;
      if (!('value' in d)) throw new Error('JSON accessors are unsupported');
      visit(d.value, depth + 1);
    }
    seen.delete(v);
  }
  visit(value, 0);
}
export function checkContractSchema(schema: ContractSchema): void {
  json(schema);
  let count = 0;
  function visit(s: ContractSchema, depth: number): void {
    if (++count > 2000 || depth > 32) throw new Error('Schema budget exceeded');
    if (!s || typeof s !== 'object' || Array.isArray(s) || !types.has(s.type))
      throw new Error('Schema requires a supported explicit type');
    const allowed = new Set([
      ...common,
      ...(s.type === 'object'
        ? ['properties', 'required', 'additionalProperties']
        : s.type === 'array'
          ? ['items', 'minItems', 'maxItems']
          : s.type === 'string'
            ? ['minLength', 'maxLength']
            : ['number', 'integer'].includes(s.type)
              ? ['minimum', 'maximum']
              : []),
    ]);
    for (const key of Object.keys(s))
      if (!allowed.has(key)) throw new Error(`Unsupported schema keyword: ${key}`);
    if (s.description !== undefined && typeof s.description !== 'string')
      throw new Error('description must be a string');
    if (s.enum !== undefined && (!Array.isArray(s.enum) || s.enum.length === 0))
      throw new Error('enum must be nonempty');
    for (const k of ['minimum', 'maximum', 'minLength', 'maxLength', 'minItems', 'maxItems'] as const)
      if (
        s[k] !== undefined &&
        (typeof s[k] !== 'number' ||
          !Number.isFinite(s[k]) ||
          (!['minimum', 'maximum'].includes(k) && (!Number.isSafeInteger(s[k]) || s[k]! < 0)))
      )
        throw new Error(`Invalid ${k}`);
    for (const [a, b] of [
      ['minimum', 'maximum'],
      ['minLength', 'maxLength'],
      ['minItems', 'maxItems'],
    ] as const)
      if (s[a] !== undefined && s[b] !== undefined && s[a]! > s[b]!)
        throw new Error(`Invalid ${a}/${b} range`);
    if (s.type === 'object') {
      if (
        s.properties !== undefined &&
        (!s.properties || typeof s.properties !== 'object' || Array.isArray(s.properties))
      )
        throw new Error('properties must be an object');
      if (s.additionalProperties !== undefined && typeof s.additionalProperties !== 'boolean')
        throw new Error('additionalProperties must be boolean');
      if (
        s.required !== undefined &&
        (!Array.isArray(s.required) ||
          s.required.some((k) => typeof k !== 'string') ||
          new Set(s.required).size !== s.required.length)
      )
        throw new Error('required must contain unique strings');
      for (const child of Object.values(s.properties ?? {})) visit(child, depth + 1);
    }
    if (s.type === 'array') {
      if (!s.items) throw new Error('Array schema requires items');
      visit(s.items, depth + 1);
    }
  }
  visit(schema, 0);
}
function equal(a: unknown, b: unknown): boolean {
  if (a === b) return true;
  if (!a || !b || typeof a !== 'object' || typeof b !== 'object' || Array.isArray(a) !== Array.isArray(b))
    return false;
  const keys = Object.keys(a);
  return (
    keys.length === Object.keys(b).length &&
    keys.every((k) => own(b, k) && equal((a as any)[k], (b as any)[k]))
  );
}
/** Paths in results are RFC 6901 JSON pointers. An empty path identifies the root. */
export function validateContract(
  schema: ContractSchema,
  data: unknown,
): { valid: boolean; issues: ContractIssue[] } {
  checkContractSchema(schema);
  json(data);
  const issues: ContractIssue[] = [];
  function visit(s: ContractSchema, v: unknown, p: string): void {
    if (issues.length >= 1000) return;
    const add = (code: string, message: string) => issues.push({ path: p, code, message });
    const actual = v === null ? 'null' : Array.isArray(v) ? 'array' : typeof v;
    if (s.type === 'integer' ? !Number.isSafeInteger(v) : actual !== s.type) {
      add('type', `Expected ${s.type}`);
      return;
    }
    if (s.enum && !s.enum.some((e) => equal(e, v))) add('enum', 'Value is outside enum');
    if (typeof v === 'number') {
      if (s.minimum !== undefined && v < s.minimum) add('minimum', 'Value below minimum');
      if (s.maximum !== undefined && v > s.maximum) add('maximum', 'Value above maximum');
    }
    if (typeof v === 'string') {
      const n = Array.from(v).length;
      if (s.minLength !== undefined && n < s.minLength) add('minLength', 'String is too short');
      if (s.maxLength !== undefined && n > s.maxLength) add('maxLength', 'String is too long');
    }
    if (Array.isArray(v)) {
      if (s.minItems !== undefined && v.length < s.minItems) add('minItems', 'Array is too short');
      if (s.maxItems !== undefined && v.length > s.maxItems) add('maxItems', 'Array is too long');
      v.forEach((x, i) => visit(s.items!, x, pointer(p, i)));
    } else if (v !== null && typeof v === 'object') {
      for (const key of s.required ?? [])
        if (!own(v, key))
          issues.push({ path: pointer(p, key), code: 'required', message: 'Required property is missing' });
      for (const key of Object.keys(v)) {
        if (own(s.properties ?? {}, key)) visit(s.properties![key]!, (v as any)[key], pointer(p, key));
        else if (s.additionalProperties === false)
          issues.push({
            path: pointer(p, key),
            code: 'additionalProperties',
            message: 'Unexpected property',
          });
      }
    }
  }
  visit(schema, data, '');
  return { valid: issues.length === 0, issues };
}
export function contractTypeScript(schema: ContractSchema, name = 'TemplateData'): string {
  checkContractSchema(schema);
  if (
    !/^[A-Za-z_$][\w$]*$/.test(name) ||
    new Set(
      'abstract any as asserts async await bigint boolean break case catch class const constructor continue debugger declare default delete do else enum export extends false finally for from function get global if implements import in infer instanceof interface intrinsic is keyof let module namespace never new null number object of package private protected public readonly require return set static string super switch symbol this throw true try type typeof undefined unique unknown using var void while with yield'.split(
        ' ',
      ),
    ).has(name)
  )
    throw new Error('Invalid type name');
  function render(s: ContractSchema): string {
    if (s.enum && s.enum.every((v) => v === null || ['string', 'number', 'boolean'].includes(typeof v))) {
      const base = { ...s };
      delete base.enum;
      return `(${s.enum.map((v) => JSON.stringify(v)).join(' | ')}) & (${render(base)})`;
    }
    if (s.type === 'array') return `Array<${render(s.items!)}>`;
    if (s.type === 'object')
      return `{ ${Object.entries(s.properties ?? {})
        .map(([k, v]) => `${JSON.stringify(k)}${s.required?.includes(k) ? '' : '?'}: ${render(v)};`)
        .join(' ')}${(s.required ?? [])
        .filter((k) => !own(s.properties ?? {}, k))
        .map((k) => `${JSON.stringify(k)}: ${s.additionalProperties === false ? 'never' : 'unknown'};`)
        .join(' ')}${s.additionalProperties !== false ? ' [key: string]: unknown;' : ''} }`;
    return s.type === 'integer' ? 'number' : s.type;
  }
  return `export type ${name} = ${render(schema)};\n`;
}
function at(data: unknown, path: string): unknown {
  if (path === '') return data;
  if (typeof path !== 'string' || !path.startsWith('/') || /~(?![01])/u.test(path))
    throw new Error('Expected an RFC 6901 JSON pointer');
  let v = data;
  for (const segment of path
    .slice(1)
    .split('/')
    .map((s) => s.replace(/~1/g, '/').replace(/~0/g, '~'))) {
    if (Array.isArray(v) && !/^(0|[1-9][0-9]*)$/.test(segment))
      throw new Error('Invalid array pointer index');
    if (!v || typeof v !== 'object' || !own(v, segment)) throw new Error(`Missing assertion input: ${path}`);
    v = (v as any)[segment];
  }
  return v;
}
export function evaluateAssertions(
  data: unknown,
  assertions: BusinessAssertion[],
): { valid: boolean; issues: ContractIssue[] } {
  json(data);
  json(assertions);
  if (!Array.isArray(assertions) || assertions.length > 1000) throw new Error('Assertion budget exceeded');
  const issues: ContractIssue[] = [];
  for (const a of assertions) {
    if (
      !a ||
      typeof a.name !== 'string' ||
      !['equals', 'lessThanOrEqual', 'greaterThanOrEqual'].includes(a.operator) ||
      !a.left ||
      !a.right ||
      Object.keys(a).some((k) => !['name', 'left', 'right', 'operator', 'tolerance'].includes(k))
    )
      throw new Error('Invalid business assertion');
    const tol = a.tolerance ?? 0;
    if (typeof tol !== 'number' || !Number.isFinite(tol) || tol < 0)
      throw new Error('Invalid assertion tolerance');
    const lkeys = Object.keys(a.left),
      rkeys = Object.keys(a.right);
    if (
      !(
        (lkeys.length === 1 && own(a.left, 'path')) ||
        (own(a.left, 'sum') && lkeys.every((k) => ['sum', 'field'].includes(k)))
      ) ||
      rkeys.length !== 1 ||
      !['path', 'value'].includes(rkeys[0]!)
    )
      throw new Error('Invalid assertion operands');
    try {
      let left: unknown;
      if ('path' in a.left) left = at(data, a.left.path);
      else {
        const rows = at(data, a.left.sum);
        if (!Array.isArray(rows)) throw new Error('Sum input must be an array');
        const field = a.left.field;
        left = rows.reduce<number>((sum, row) => {
          const n = field === undefined ? row : at(row, field);
          if (typeof n !== 'number' || !Number.isFinite(n)) throw new Error('Sum requires finite numbers');
          const next = sum + n;
          if (!Number.isFinite(next)) throw new Error('Sum overflow');
          return next;
        }, 0);
      }
      const right = 'path' in a.right ? at(data, a.right.path) : a.right.value;
      const numeric = typeof left === 'number' && typeof right === 'number';
      if (a.operator !== 'equals' && !numeric) throw new Error('Ordering requires numeric operands');
      if (tol !== 0 && !numeric) throw new Error('Tolerance requires numeric operands');
      const pass =
        a.operator === 'equals'
          ? numeric
            ? Math.abs((left as number) - (right as number)) <= tol
            : equal(left, right)
          : a.operator === 'lessThanOrEqual'
            ? (left as number) - (right as number) <= tol
            : (right as number) - (left as number) <= tol;
      if (!pass) issues.push({ path: a.name, code: 'assertion', message: 'Business assertion failed' });
    } catch (e) {
      issues.push({ path: a.name, code: 'assertion-input', message: (e as Error).message });
    }
  }
  return { valid: issues.length === 0, issues };
}
/** Deterministic structural sample, not a realistic or exhaustive dataset generator. */
export function contractSample(schema: ContractSchema): unknown {
  checkContractSchema(schema);
  let nodes = 0;
  function make(s: ContractSchema): unknown {
    if (++nodes > 2000) throw new Error('Sample budget exceeded');
    if (s.enum) {
      const found = s.enum.find((v) => validateContract(s, v).valid);
      if (found === undefined) throw new Error('No valid enum sample');
      return JSON.parse(JSON.stringify(found));
    }
    if (s.type === 'null') return null;
    if (s.type === 'boolean') return false;
    if (s.type === 'string') {
      if ((s.minLength ?? 0) > 10000) throw new Error('Sample string budget exceeded');
      return 'x'.repeat(s.minLength ?? 0);
    }
    if (s.type === 'number' || s.type === 'integer') {
      const minimum = s.minimum ?? Math.min(0, s.maximum ?? 0);
      return s.type === 'integer'
        ? s.minimum !== undefined
          ? Math.ceil(minimum)
          : Math.floor(minimum)
        : minimum;
    }
    if (s.type === 'array') {
      const n = s.minItems ?? 0;
      if (n > 2000) throw new Error('Sample array budget exceeded');
      return Array.from({ length: n }, () => make(s.items!));
    }
    const out: Record<string, unknown> = {};
    for (const key of s.required ?? []) {
      const child = own(s.properties ?? {}, key) ? s.properties![key] : undefined;
      Object.defineProperty(out, key, {
        value: child ? make(child) : null,
        enumerable: true,
        writable: true,
        configurable: true,
      });
    }
    return out;
  }
  const sample = make(schema);
  if (!validateContract(schema, sample).valid)
    throw new Error('Cannot generate a valid sample for this schema');
  return sample;
}
