import { compileText } from './text-template.ts';
import { validateData } from './data.ts';
import type { CompiledTemplate, RenderOptions } from './types.ts';
import { forbidden } from './scope.ts';

export interface MigrationIssue {
  offset: number;
  tag: string;
  reason: string;
}
export interface CarboneMigration {
  source: string;
  originalSource: string;
  requirements: { stringFields: string[]; locale?: 'en-US' };
  compatible: boolean;
  fields: string[];
  issues: MigrationIssue[];
  /** Compatibility is lexical only: compile the actual Office package and compare rendered output. */
  requiredChecks: string[];
}

/** Analyze extracted logical text, never raw Office XML (tags may span runs).
 * Scalar substitutions and explicitly contracted string case conversions are accepted.
 * Case conversion requires en-US render options and string input validation by the caller.
 * Unsupported input is returned unchanged, so callers can display a lossless migration report.
 */
export function migrateCarboneText(
  source: string,
  options: { stringFields?: readonly string[]; locale?: 'en-US' } = {},
): CarboneMigration {
  if (typeof source !== 'string' || source.length > 1_000_000)
    throw new Error('Migration source must be text of at most 1000000 characters');
  if (options.locale !== undefined && options.locale !== 'en-US')
    throw new Error('Case migration requires en-US locale');
  if (
    options.stringFields !== undefined &&
    (!Array.isArray(options.stringFields) ||
      options.stringFields.length > 10000 ||
      options.stringFields.some((f) => typeof f !== 'string'))
  )
    throw new Error('Invalid string field contract');
  const fields = new Set<string>(),
    issues: MigrationIssue[] = [],
    edits: { start: number; end: number; value: string }[] = [];
  const requiredStrings = new Set<string>();
  const tags = /\{[^{}]*\}|[{}]/g;
  let tagCount = 0;
  for (const match of source.matchAll(tags)) {
    if (++tagCount > 10000) throw new Error('Migration tag budget exceeded');
    const tag = match[0],
      offset = match.index;
    const scalar = /^\{d\.([A-Za-z_$][\w$]*(?:\.[A-Za-z_$][\w$]*)*)\}$/.exec(tag);
    if (scalar && !scalar[1]!.split('.').some((key) => forbidden.has(key))) fields.add(scalar[1]!);
    else if (
      /^\{d\.([A-Za-z_$][\w$]*(?:\.[A-Za-z_$][\w$]*)*)((?::(?:upperCase|lowerCase)(?:\(\))?){1,16})\}$/.test(
        tag,
      )
    ) {
      const parts = tag.slice(1, -1).split(':'),
        path = parts.shift()!.slice(2);
      if (
        options.locale === 'en-US' &&
        options.stringFields?.includes(path) &&
        !path.split('.').some((k) => forbidden.has(k))
      ) {
        fields.add(path);
        requiredStrings.add(path);
        edits.push({
          start: offset,
          end: offset + tag.length,
          value: `{d.${path}:${parts.map((f) => (f.startsWith('upper') ? 'upper()' : 'lower()')).join(':')}}`,
        });
      } else
        issues.push({
          offset,
          tag,
          reason:
            'Case mapping requires an explicit stringFields contract and en-US locale; non-string and locale behavior differs.',
        });
    } else
      issues.push({
        offset,
        tag,
        reason:
          'Only scalar {d.path} substitutions are supported; braces, formatters, repeats, aliases, options and complement tags require manual migration.',
      });
    if (fields.size + issues.length > 10000) throw new Error('Migration tag budget exceeded');
  }
  let migrated = source;
  if (issues.length === 0)
    for (const edit of edits.reverse())
      migrated = migrated.slice(0, edit.start) + edit.value + migrated.slice(edit.end);
  return {
    source: migrated,
    originalSource: source,
    requirements: {
      stringFields: [...requiredStrings].sort(),
      ...(requiredStrings.size ? { locale: 'en-US' as const } : {}),
    },
    compatible: issues.length === 0,
    fields: [...fields].sort(),
    issues,
    requiredChecks: [
      'Compile the actual template package',
      'Compare representative rendered output and business values',
      'Review missing, null, type and locale behavior',
    ],
  };
}

export interface CrystalMetadata {
  schema: 'imprint-crystal-inventory-v1';
  name: string;
  fields: string[];
  formulas: { name: string; expression: string }[];
  groups: string[];
  subreports: string[];
}
/** Accepts our explicit interchange schema populated by an authorized external extractor.
 * This is not an SAP export parser, RPT decoder or formula translator.
 */
export function inventoryCrystalMetadata(value: unknown) {
  try {
    validateData({ metadata: value });
  } catch {
    throw new Error('Invalid Crystal interchange metadata; expected plain JSON');
  }
  const object = (v: unknown): v is Record<string, unknown> =>
    !!v && typeof v === 'object' && !Array.isArray(v);
  const text = (v: unknown): v is string => typeof v === 'string' && v.length > 0 && v.length <= 100000;
  const list = (v: unknown): v is string[] =>
    Array.isArray(v) && v.length <= 10000 && Array.from(v).every(text);
  if (
    !object(value) ||
    value.schema !== 'imprint-crystal-inventory-v1' ||
    !text(value.name) ||
    Object.keys(value).some(
      (k) => !['schema', 'name', 'fields', 'formulas', 'groups', 'subreports'].includes(k),
    ) ||
    !list(value.fields) ||
    !list(value.groups) ||
    !list(value.subreports) ||
    !Array.isArray(value.formulas) ||
    value.formulas.length > 10000 ||
    !Array.from(value.formulas).every(
      (f) => object(f) && Object.keys(f).length === 2 && text(f.name) && text(f.expression),
    )
  )
    throw new Error('Invalid Crystal interchange metadata; raw RPT and native exports are not supported');
  let total = (value.name as string).length;
  for (const list of [value.fields, value.groups, value.subreports] as string[][])
    for (const item of list) total += item.length;
  for (const formula of value.formulas as { name: string; expression: string }[])
    total += formula.name.length + formula.expression.length;
  if (total > 10_000_000) throw new Error('Crystal metadata exceeds aggregate budget');
  const metadata = structuredClone(value) as unknown as CrystalMetadata;
  return {
    metadata,
    automaticConversion: false as const,
    tasks: [
      ...metadata.fields.map((name) => ({
        kind: 'field' as const,
        name,
        action: 'Map to a JSON input path',
      })),
      ...metadata.formulas.map((f) => ({
        kind: 'formula' as const,
        name: f.name,
        action: 'Reimplement and compare calculations against Crystal output',
      })),
      ...metadata.groups.map((name) => ({
        kind: 'group' as const,
        name,
        action: 'Rebuild grouping and verify ordering and pagination',
      })),
      ...metadata.subreports.map((name) => ({
        kind: 'subreport' as const,
        name,
        action: 'Extract data contract and rebuild subreport',
      })),
    ],
  };
}

/** Safe TXT execution of the bounded migration subset. This enforces all case
 * conversion contracts on every render; the migration report alone is not an execution gate. */
export function compileCarboneText(
  source: string,
  options: { stringFields?: readonly string[]; locale?: 'en-US' } = {},
): CompiledTemplate {
  const migration = migrateCarboneText(source, options);
  if (!migration.compatible)
    throw new Error('Unsupported Carbone migration; inspect migrateCarboneText issues');
  const compiled = compileText('txt', migration.source);
  const required = [...migration.requirements.stringFields];
  return Object.freeze({
    ...compiled,
    async render(data: Record<string, unknown>, renderOptions: RenderOptions = {}) {
      validateData(data);
      if (required.length && renderOptions.locale !== undefined && renderOptions.locale !== 'en-US')
        throw new Error('Migrated case formatters require en-US locale');
      for (const path of required) {
        let value: unknown = data;
        for (const key of path.split('.')) {
          value =
            value !== null && typeof value === 'object' && Object.hasOwn(value, key)
              ? (value as Record<string, unknown>)[key]
              : undefined;
        }
        if (value !== null && value !== undefined && typeof value !== 'string')
          throw new Error(`Migrated case formatter requires string data at ${path}`);
      }
      return compiled.render(data, required.length ? { ...renderOptions, locale: 'en-US' } : renderOptions);
    },
  });
}
