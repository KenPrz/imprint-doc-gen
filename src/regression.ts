import { types as utilTypes } from 'node:util';
import { cloneAssuranceJson, cloneRenderOptions } from './assurance-input.ts';
import { compile } from './index.ts';
import { compareDocuments, snapshotDocument } from './document-diff.ts';
import type { DocumentDiff } from './document-diff.ts';
import type { RenderOptions } from './types.ts';
export interface RegressionCase {
  name: string;
  data: Record<string, unknown>;
  approved: Uint8Array;
  options?: RenderOptions;
  masks?: string[];
}
export interface RegressionResult {
  name: string;
  passed: boolean;
  diff?: DocumentDiff;
  error?: string;
}
/** Captures each input before an asynchronous render can yield to caller mutation. */
export function snapshotRegressionCases(cases: RegressionCase[]): RegressionCase[] {
  if (
    utilTypes.isProxy(cases) ||
    !Array.isArray(cases) ||
    !cases.length ||
    cases.length > 1000 ||
    Object.getOwnPropertySymbols(cases).length ||
    Object.keys(cases).length !== cases.length
  )
    throw new Error('Expected 1–1000 dense regression cases');
  let budget = 0;
  return Array.from({ length: cases.length }, (_, index) => {
    const entry = Object.getOwnPropertyDescriptor(cases, String(index));
    if (!entry || !('value' in entry)) throw new Error('Regression cases cannot contain accessors');
    const item = entry.value;
    if (
      !item ||
      typeof item !== 'object' ||
      utilTypes.isProxy(item) ||
      ![Object.prototype, null].includes(Object.getPrototypeOf(item)) ||
      Object.getOwnPropertySymbols(item).length
    )
      throw new Error('Expected plain regression cases');
    const descriptors = Object.getOwnPropertyDescriptors(item);
    if (
      Object.entries(descriptors).some(
        ([key, d]) =>
          !d.enumerable || !('value' in d) || !['name', 'data', 'approved', 'options', 'masks'].includes(key),
      )
    )
      throw new Error('Invalid regression case field or accessor');
    const name = descriptors.name?.value,
      data = descriptors.data?.value,
      approved = descriptors.approved?.value;
    if (typeof name !== 'string' || !name.trim() || name.length > 200)
      throw new Error('Case names must be nonempty and unique');
    if (
      utilTypes.isProxy(approved) ||
      !(approved instanceof Uint8Array) ||
      approved.byteLength > 8 * 1024 * 1024
    )
      throw new Error('Invalid approved document');
    const safeData = cloneAssuranceJson(data);
    if (!safeData || typeof safeData !== 'object' || Array.isArray(safeData))
      throw new Error('Expected object case data');
    const options =
      descriptors.options?.value === undefined ? undefined : cloneRenderOptions(descriptors.options.value);
    const masks =
      descriptors.masks?.value === undefined ? undefined : cloneAssuranceJson(descriptors.masks.value);
    if (
      masks !== undefined &&
      (!Array.isArray(masks) || masks.length > 100 || masks.some((m) => typeof m !== 'string'))
    )
      throw new Error('Invalid comparison masks');
    budget +=
      approved.byteLength +
      Buffer.byteLength(JSON.stringify(safeData)) +
      (options ? Buffer.byteLength(JSON.stringify(options)) : 0);
    if (budget > 64 * 1024 * 1024) throw new Error('Regression suite exceeds 64 MiB');
    return {
      name,
      data: safeData,
      approved: Buffer.from(approved),
      ...(options ? { options } : {}),
      ...(masks ? { masks } : {}),
    };
  });
}
/** Runs sequentially with bounded cases. Exact bytes are the default acceptance criterion. */
export async function runRegressionSuite(
  template: Uint8Array,
  cases: RegressionCase[],
  mode: 'exact' | 'semantic' = 'exact',
): Promise<{ passed: boolean; mode: 'exact' | 'semantic'; results: RegressionResult[] }> {
  cases = snapshotRegressionCases(cases);
  if (!Array.isArray(cases) || !cases.length || cases.length > 1000)
    throw new Error('Expected 1–1000 regression cases');
  if (!['exact', 'semantic'].includes(mode)) throw new Error('Invalid comparison mode');
  if (
    new Set(cases.map((c) => c.name)).size !== cases.length ||
    cases.some((c) => typeof c.name !== 'string' || !c.name.trim() || c.name.length > 200)
  )
    throw new Error('Case names must be nonempty and unique');
  if (mode === 'exact' && cases.some((c) => c.masks?.length))
    throw new Error('Masks require explicit semantic comparison');
  if (utilTypes.isProxy(template) || !(template instanceof Uint8Array))
    throw new Error('Invalid regression template');
  const compiled = await compile(Buffer.from(template)),
    results: RegressionResult[] = [];
  for (const item of cases) {
    try {
      const output = await compiled.render(item.data, item.options);
      const diff = compareDocuments(
        await snapshotDocument(item.approved, compiled.format),
        await snapshotDocument(output, compiled.format),
        item.masks,
      );
      results.push({ name: item.name, passed: mode === 'exact' ? diff.binaryEqual : diff.equal, diff });
    } catch (error) {
      results.push({
        name: item.name,
        passed: false,
        error: error instanceof Error ? error.message : 'Regression case failed',
      });
    }
  }
  return { passed: results.every((r) => r.passed), mode, results };
}
