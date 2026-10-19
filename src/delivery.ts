import { createHash } from 'node:crypto';
import { PDFDocument } from 'pdf-lib';
import { RenderError } from './errors.ts';
export interface ValidationCheck {
  passed: boolean;
  details: string;
}
export interface ArtifactValidator {
  id: string;
  version: string;
  validate(bytes: Uint8Array, signal: AbortSignal): Promise<ValidationCheck>;
}
export interface DeliveryEvidence {
  sha256: string;
  size: number;
  status: 'approved' | 'quarantined';
  checks: { id: string; version: string; passed: boolean; details: string }[];
}
export type ValidatedDelivery = { evidence: DeliveryEvidence; artifact?: Buffer };
export const artifactDigest = (bytes: Uint8Array) => createHash('sha256').update(bytes).digest('hex');
/** Validators are trusted application code, not user-supplied JavaScript. */
export async function validateDelivery(
  input: Uint8Array,
  validators: readonly ArtifactValidator[],
  timeoutMs = 30_000,
): Promise<ValidatedDelivery> {
  if (!(input instanceof Uint8Array) || !input.byteLength || input.byteLength > 32 * 1024 * 1024)
    throw new RenderError('DELIVERY_INPUT', 'Expected a nonempty artifact up to 32 MiB');
  if (
    !Array.isArray(validators) ||
    validators.length < 1 ||
    validators.length > 32 ||
    !Number.isInteger(timeoutMs) ||
    timeoutMs < 1 ||
    timeoutMs > 120_000
  )
    throw new RenderError('DELIVERY_POLICY', 'Require 1–32 validators and a 1–120000 ms deadline');
  const ids = new Set<string>();
  const policy = validators.map((v) => {
    if (
      !v ||
      typeof v.id !== 'string' ||
      !/^[a-zA-Z0-9._-]{1,80}$/.test(v.id) ||
      typeof v.version !== 'string' ||
      !v.version.length ||
      v.version.length > 100 ||
      typeof v.validate !== 'function' ||
      ids.has(v.id)
    )
      throw new RenderError('DELIVERY_POLICY', 'Invalid or duplicate validator identity');
    ids.add(v.id);
    return { id: v.id, version: v.version, validate: v.validate.bind(v) };
  });
  const artifact = Buffer.from(input);
  const evidence: DeliveryEvidence = {
    sha256: artifactDigest(artifact),
    size: artifact.length,
    status: 'quarantined',
    checks: [],
  };
  for (const validator of policy) {
    const controller = new AbortController();
    let timer: ReturnType<typeof setTimeout> | undefined;
    let result: ValidationCheck;
    try {
      const check = await Promise.race([
        Promise.resolve().then(() => validator.validate(Buffer.from(artifact), controller.signal)),
        new Promise<never>((_, reject) => {
          timer = setTimeout(() => {
            controller.abort();
            reject(new Error('Validation deadline exceeded'));
          }, timeoutMs);
        }),
      ]);
      if (
        !check ||
        typeof check.passed !== 'boolean' ||
        typeof check.details !== 'string' ||
        check.details.length > 4096
      )
        throw new Error('Invalid validator result');
      result = { passed: check.passed, details: check.details };
    } catch {
      result = {
        passed: false,
        details: controller.signal.aborted
          ? 'Validation deadline exceeded'
          : 'Validator failed or returned an invalid result',
      };
    } finally {
      clearTimeout(timer);
      controller.abort();
    }
    evidence.checks.push({ id: validator.id, version: validator.version, ...result });
  }
  if (evidence.checks.every((c) => c.passed)) {
    evidence.status = 'approved';
    return { evidence, artifact };
  }
  return { evidence };
}
/** Structural parsing only; deliberately does not claim PDF/A or PDF/UA conformance. */
export function pdfStructureValidator(maxPages = 1000): ArtifactValidator {
  if (!Number.isInteger(maxPages) || maxPages < 1 || maxPages > 10000)
    throw new RenderError('DELIVERY_POLICY', 'Invalid PDF page limit');
  return {
    id: 'pdf-structure',
    version: 'pdf-lib-1.17.1',
    async validate(bytes) {
      if (!Buffer.from(bytes.subarray(0, 5)).equals(Buffer.from('%PDF-')))
        return { passed: false, details: 'Missing PDF header' };
      const pdf = await PDFDocument.load(bytes, { throwOnInvalidObject: true, updateMetadata: false });
      const pages = pdf.getPageCount();
      return { passed: pages > 0 && pages <= maxPages, details: `${pages} pages; structural check only` };
    },
  };
}
