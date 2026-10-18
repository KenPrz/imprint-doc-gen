import { artifactDigest } from './delivery.ts';
import { inspectTemplate } from './inspect.ts';
import { runRegressionSuite, snapshotRegressionCases } from './regression.ts';
import type { RegressionCase, RegressionResult } from './regression.ts';
import type { Inspection } from './inspect.ts';
import { RenderError } from './errors.ts';
export interface TemplateReview {
  baseSha256: string;
  candidateSha256: string;
  inspection: Inspection;
  mode: 'exact' | 'semantic';
  passed: boolean;
  results: RegressionResult[];
  addedFields: string[];
  removedFields: string[];
}
const reviewed = new WeakMap<TemplateReview, { baseSha256: string; candidate: Buffer; passed: boolean }>();
/** AI providers supply candidate bytes only. No provider-generated code or claims are executed. */
export async function reviewTemplateCandidate(
  base: Uint8Array,
  candidate: Uint8Array,
  cases: RegressionCase[],
  mode: 'exact' | 'semantic' = 'exact',
): Promise<TemplateReview> {
  if (
    !(base instanceof Uint8Array) ||
    !(candidate instanceof Uint8Array) ||
    base.byteLength > 8 * 1024 * 1024 ||
    candidate.byteLength > 8 * 1024 * 1024
  )
    throw new RenderError('REVIEW_INPUT', 'Invalid template bytes');
  if (!Array.isArray(cases) || !cases.length || cases.length > 1000)
    throw new RenderError('REVIEW_INPUT', 'Require 1–1000 approved cases');
  const before = Buffer.from(base),
    after = Buffer.from(candidate);
  // Freeze caller-owned cases before the first await; validate prior to cloning to reject accessors.
  const fixtures = snapshotRegressionCases(cases);
  const baseInspection = await inspectTemplate(before),
    inspection = await inspectTemplate(after);
  if (!baseInspection.valid) throw new RenderError('REVIEW_INPUT', 'Base template is invalid');
  const a = new Set(baseInspection.fields.map((f) => f.path)),
    b = new Set(inspection.fields.map((f) => f.path));
  const suite =
    inspection.valid && baseInspection.format === inspection.format
      ? await runRegressionSuite(after, fixtures, mode)
      : { passed: false, results: [] };
  const review: TemplateReview = {
    baseSha256: artifactDigest(before),
    candidateSha256: artifactDigest(after),
    inspection,
    mode,
    passed: suite.passed,
    results: suite.results,
    addedFields: [...b].filter((p) => !a.has(p)),
    removedFields: [...a].filter((p) => !b.has(p)),
  };
  reviewed.set(review, { baseSha256: review.baseSha256, candidate: after, passed: suite.passed });
  return review;
}
/** In-process handoff only; persistence must re-run review. Does not publish or write files. */
export function acceptTemplateCandidate(review: TemplateReview, currentBase: Uint8Array): Buffer {
  const record = reviewed.get(review);
  if (!record || !record.passed)
    throw new RenderError('REVIEW_REQUIRED', 'A passing review from this process is required');
  if (artifactDigest(currentBase) !== record.baseSha256)
    throw new RenderError('STALE_REVIEW', 'Base template changed after review');
  return Buffer.from(record.candidate);
}
