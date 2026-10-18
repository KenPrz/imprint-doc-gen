import { RenderError } from './errors.ts';
import { validateData } from './data.ts';
import { decodeBase64 } from './base64.ts';
import {
  validateContract,
  contractTypeScript,
  contractSample,
  evaluateAssertions,
  type ContractSchema,
  type BusinessAssertion,
} from './contracts.ts';
import { renderWithTrace } from './trace.ts';
import { compareDocuments, snapshotDocument } from './document-diff.ts';
import { runRegressionSuite, type RegressionCase } from './regression.ts';
import {
  createReportSnapshot,
  reportSnapshotHtml,
  type ReportDefinition,
  type ReportFilter,
} from './report-view.ts';
import { migrateCarboneText, compileCarboneText, inventoryCrystalMetadata } from './template-migration.ts';
import { ComponentRegistry, type DocumentComponent } from './components.ts';
import { assembleDocxComponents } from './docx-components.ts';
import { inspectCrystalArtifact } from './crystal-artifact.ts';
import { reviewTemplateCandidate, acceptTemplateCandidate } from './template-review.ts';
import { createReplayBundle, replayBundle } from './replay.ts';
import { validateDelivery, pdfStructureValidator } from './delivery.ts';
import type { DocumentFormat, RenderOptions } from './types.ts';

export const assuranceOperations = [
  'contract',
  'trace',
  'diff',
  'regression',
  'report',
  'migrate',
  'components',
  'docx-components',
  'crystal-artifact',
  'review',
  'replay-create',
  'replay',
  'delivery',
] as const;
function fail(message: string): never {
  throw new RenderError('ASSURANCE_INPUT', message);
}
function object(value: unknown, required: string[], optional: string[] = []): Record<string, unknown> {
  if (!value || typeof value !== 'object' || Array.isArray(value)) fail('Expected an object');
  const out = value as Record<string, unknown>;
  if (
    required.some((k) => !Object.hasOwn(out, k)) ||
    Object.keys(out).some((k) => ![...required, ...optional].includes(k))
  )
    fail('Missing or unsupported operation properties');
  return out;
}
function bytes(value: unknown, allowEmpty = false): Buffer {
  if (allowEmpty && value === '') return Buffer.alloc(0);
  if (typeof value !== 'string' || value.length > 12 * 1024 * 1024) fail('Expected base64 bytes up to 8 MiB');
  const result = decodeBase64(value);
  if (result.length > 8 * 1024 * 1024) fail('Artifact exceeds 8 MiB');
  return result;
}
function data(value: unknown): Record<string, unknown> {
  validateData(value);
  return value;
}
function text(value: unknown, max = 1_000_000): string {
  if (typeof value !== 'string' || value.length > max) fail('Expected bounded text');
  return value;
}
function strings(value: unknown, max = 10000): string[] {
  if (
    !Array.isArray(value) ||
    value.length > max ||
    value.some((v) => typeof v !== 'string' || v.length > 2048)
  )
    fail('Expected a bounded string array');
  return value;
}
function options(value: unknown): RenderOptions {
  if (value === undefined) return {};
  const out = object(value, [], ['locale', 'timeZone', 'restartNumbering']);
  for (const k of ['locale', 'timeZone']) if (out[k] !== undefined) text(out[k], 200);
  if (out.restartNumbering !== undefined && typeof out.restartNumbering !== 'boolean')
    fail('restartNumbering must be boolean');
  return out as RenderOptions;
}
function mode(value: unknown): 'exact' | 'semantic' {
  if (value === undefined) return 'exact';
  if (value !== 'exact' && value !== 'semantic') fail('Expected exact or semantic comparison');
  return value;
}
function cases(value: unknown): RegressionCase[] {
  if (!Array.isArray(value) || !value.length || value.length > 100) fail('Require 1–100 approved cases');
  return value.map((item) => {
    const c = object(item, ['name', 'data', 'approved'], ['options', 'masks']);
    return {
      name: text(c.name, 200),
      data: data(c.data),
      approved: bytes(c.approved, true),
      options: options(c.options),
      ...(c.masks === undefined ? {} : { masks: strings(c.masks, 100) }),
    };
  });
}
/** JSON-only worker entry point. No caller-selected callbacks, modules, paths or commands. */
export async function executeAssurance(payload: unknown): Promise<unknown> {
  try {
    validateData(payload);
    const op = payload.operation;
    if (typeof op !== 'string' || !assuranceOperations.includes(op as (typeof assuranceOperations)[number]))
      fail('Unsupported assurance operation');
    if (op === 'contract') {
      const p = object(payload, ['operation', 'schema', 'data'], ['assertions']);
      const schema = p.schema as ContractSchema;
      const validation = validateContract(schema, p.data);
      return {
        validation,
        assertions: evaluateAssertions(
          p.data,
          (p.assertions === undefined ? [] : p.assertions) as BusinessAssertion[],
        ),
        typescript: contractTypeScript(schema),
        sample: contractSample(schema),
      };
    }
    if (op === 'trace') {
      const p = object(payload, ['operation', 'template', 'data'], ['options', 'includeValues']);
      if (p.includeValues !== undefined && typeof p.includeValues !== 'boolean')
        fail('includeValues must be boolean');
      const result = await renderWithTrace(bytes(p.template), data(p.data), options(p.options), {
        includeValues: p.includeValues as boolean | undefined,
      });
      return result.ok
        ? { ok: true, artifact: result.value.toString('base64'), trace: result.trace }
        : result;
    }
    if (op === 'diff') {
      const p = object(payload, ['operation', 'before', 'after', 'format'], ['masks']);
      if (!['docx', 'xlsx', 'pptx', 'odt', 'ods', 'odp', 'txt', 'csv'].includes(String(p.format)))
        fail('Unsupported document format');
      const before = await snapshotDocument(bytes(p.before, true), p.format as DocumentFormat);
      const after = await snapshotDocument(bytes(p.after, true), p.format as DocumentFormat);
      return compareDocuments(before, after, p.masks === undefined ? [] : strings(p.masks, 100));
    }
    if (op === 'regression') {
      const p = object(payload, ['operation', 'template', 'cases'], ['mode']);
      return runRegressionSuite(bytes(p.template), cases(p.cases), mode(p.mode));
    }
    if (op === 'report') {
      const p = object(payload, ['operation', 'definition'], ['filters']);
      const snapshot = createReportSnapshot(
        p.definition as ReportDefinition,
        (p.filters === undefined ? [] : p.filters) as ReportFilter[],
      );
      return { snapshot, html: reportSnapshotHtml(snapshot) };
    }
    if (op === 'migrate') {
      const p = object(payload, ['operation', 'sourceType', 'source'], ['stringFields', 'locale', 'data']);
      if (p.sourceType === 'crystal-inventory') {
        if (p.stringFields !== undefined || p.locale !== undefined || p.data !== undefined)
          fail('Carbone options do not apply to Crystal inventory');
        return inventoryCrystalMetadata(p.source);
      }
      if (p.sourceType !== 'carbone-text') fail('Expected carbone-text or crystal-inventory');
      if (p.locale !== undefined && p.locale !== 'en-US') fail('Case migration requires en-US');
      const source = text(p.source),
        migrationOptions = {
          ...(p.stringFields === undefined ? {} : { stringFields: strings(p.stringFields) }),
          ...(p.locale === undefined ? {} : { locale: 'en-US' as const }),
        };
      const migration = migrateCarboneText(source, migrationOptions);
      if (p.data === undefined) return { migration };
      const artifact = await compileCarboneText(source, migrationOptions).render(data(p.data));
      return { migration, artifact: artifact.toString('base64') };
    }
    if (op === 'crystal-artifact') {
      const p = object(payload, ['operation', 'artifact']);
      return inspectCrystalArtifact(bytes(p.artifact));
    }
    if (op === 'docx-components') {
      const p = object(payload, ['operation', 'template', 'components']);
      if (!Array.isArray(p.components) || p.components.length > 64)
        fail('Expected up to 64 pinned DOCX components');
      const components = p.components.map((item) => {
        const c = object(item, ['id', 'version', 'document', 'sha256']);
        return {
          id: text(c.id, 128),
          version: text(c.version, 128),
          document: bytes(c.document),
          sha256: text(c.sha256, 64),
        };
      });
      const assembled = await assembleDocxComponents(bytes(p.template), components);
      return { artifact: assembled.document.toString('base64'), lock: assembled.lock };
    }
    if (op === 'components') {
      const p = object(payload, ['operation', 'components', 'source'], ['impact']);
      if (!Array.isArray(p.components)) fail('Expected component array');
      for (const c of p.components) object(c, ['id', 'version', 'source']);
      const registry = new ComponentRegistry(p.components as DocumentComponent[]),
        expanded = registry.expand(text(p.source));
      if (p.impact === undefined) return expanded;
      const impact = object(p.impact, ['id', 'version', 'templates']);
      const templates = data(impact.templates);
      for (const value of Object.values(templates)) text(value);
      return {
        ...expanded,
        impactedTemplates: registry.impact(
          text(impact.id, 128),
          text(impact.version, 128),
          templates as Record<string, string>,
        ),
      };
    }
    if (op === 'review') {
      const p = object(payload, ['operation', 'base', 'candidate', 'cases'], ['mode']);
      const base = bytes(p.base),
        candidate = bytes(p.candidate);
      const review = await reviewTemplateCandidate(base, candidate, cases(p.cases), mode(p.mode));
      return {
        review,
        ...(review.passed ? { artifact: acceptTemplateCandidate(review, base).toString('base64') } : {}),
      };
    }
    if (op === 'replay-create') {
      const p = object(payload, ['operation', 'template', 'data'], ['options']);
      return createReplayBundle(bytes(p.template), data(p.data), options(p.options));
    }
    if (op === 'replay') {
      const p = object(payload, ['operation', 'bundle']);
      return { artifact: (await replayBundle(p.bundle)).toString('base64') };
    }
    const p = object(payload, ['operation', 'artifact'], ['maxPages']);
    if (p.maxPages !== undefined && (typeof p.maxPages !== 'number' || !Number.isInteger(p.maxPages)))
      fail('maxPages must be an integer');
    const result = await validateDelivery(bytes(p.artifact), [
      pdfStructureValidator(p.maxPages as number | undefined),
    ]);
    return {
      evidence: result.evidence,
      ...(result.artifact ? { artifact: result.artifact.toString('base64') } : {}),
    };
  } catch (error) {
    if (error instanceof RenderError) throw error;
    throw new RenderError(
      'ASSURANCE_INPUT',
      error instanceof Error ? error.message : 'Invalid assurance input',
    );
  }
}
