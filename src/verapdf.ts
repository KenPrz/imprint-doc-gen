import { execFile } from 'node:child_process';
import { promisify } from 'node:util';
import { mkdtemp, writeFile, rm } from 'node:fs/promises';
import { tmpdir } from 'node:os';
import { join, isAbsolute } from 'node:path';
import { parseXml, descendants } from './xml.ts';
import type { ArtifactValidator, ValidationCheck } from './delivery.ts';
const exec = promisify(execFile);
export type VeraPdfProfile =
  '1a' | '1b' | '2a' | '2b' | '2u' | '3a' | '3b' | '3u' | '4' | '4e' | '4f' | 'ua1' | 'ua2';
const profiles = new Set(['1a', '1b', '2a', '2b', '2u', '3a', '3b', '3u', '4', '4e', '4f', 'ua1', 'ua2']);
export function parseVeraPdfReport(
  xml: string,
  profile: VeraPdfProfile,
  version: string,
  expectedInput?: { path: string; size: number },
): ValidationCheck {
  if (!profiles.has(profile) || typeof xml !== 'string' || Buffer.byteLength(xml) > 4 * 1024 * 1024)
    throw new Error('Invalid veraPDF report');
  const doc = parseXml(xml, 'veraPDF report');
  const root = doc.documentElement;
  if (root?.tagName !== 'report' || root.namespaceURI) throw new Error('Invalid veraPDF root');
  const one = (parent: import('@xmldom/xmldom').Element, name: string) => {
    const nodes = Array.from(parent.childNodes).filter(
      (n) => n.nodeType === 1 && n.nodeName === name,
    ) as import('@xmldom/xmldom').Element[];
    if (nodes.length !== 1 || nodes[0].namespaceURI) throw new Error('Missing or duplicate veraPDF ' + name);
    return nodes[0];
  };
  const build = one(root, 'buildInformation'),
    jobs = one(root, 'jobs'),
    job = one(jobs, 'job'),
    report = one(job, 'validationReport'),
    summary = one(root, 'batchSummary');
  if (
    descendants(root, '*', 'job').length !== 1 ||
    descendants(root, '*', 'validationReport').length !== 1 ||
    descendants(root, '*', 'batchSummary').length !== 1
  )
    throw new Error('Unexpected veraPDF jobs');
  const releases = descendants(build, '*', 'releaseDetails').filter((r) => r.getAttribute('id') === 'core');
  if (releases.length !== 1 || releases[0].namespaceURI || releases[0].getAttribute('version') !== version)
    throw new Error('Unexpected veraPDF core version');
  const expected = profile.startsWith('ua')
    ? 'PDF/UA-' + profile.slice(2) + ' validation profile'
    : 'PDF/A-' + profile + ' validation profile';
  if (report.getAttribute('profileName') !== expected)
    throw new Error('Unexpected veraPDF validation profile');
  const item = one(job, 'item'),
    name = one(item, 'name').textContent;
  if (!name || !/^[1-9][0-9]*$/.test(item.getAttribute('size') ?? ''))
    throw new Error('Missing veraPDF input identity');
  if (
    expectedInput &&
    (name !== expectedInput.path || item.getAttribute('size') !== String(expectedInput.size))
  )
    throw new Error('veraPDF reported a different input artifact');
  const counts = one(summary, 'validationReports'),
    details = one(report, 'details');
  const passed =
    report.getAttribute('isCompliant') === 'true' &&
    report.getAttribute('jobEndStatus') === 'normal' &&
    summary.getAttribute('totalJobs') === '1' &&
    ['failedToParse', 'encrypted', 'outOfMemory', 'veraExceptions'].every(
      (k) => summary.getAttribute(k) === '0',
    ) &&
    counts.getAttribute('compliant') === '1' &&
    counts.getAttribute('nonCompliant') === '0' &&
    counts.getAttribute('failedJobs') === '0' &&
    counts.textContent?.trim() === '1' &&
    details.getAttribute('failedRules') === '0' &&
    details.getAttribute('failedChecks') === '0' &&
    ['passedRules', 'passedChecks'].every((k) => /^[1-9][0-9]*$/.test(details.getAttribute(k) ?? ''));
  return {
    passed,
    details: `veraPDF ${version}: ${profile} ${passed ? 'passed' : 'failed'}; machine-verifiable checks only`,
  };
}
/** Trusted, locally installed executable. No shell, profiles or executables from HTTP requests. */
export function veraPdfValidator(config: {
  executable: string;
  version: string;
  profile: VeraPdfProfile;
  timeoutMs?: number;
}): ArtifactValidator {
  const { executable, version, profile, timeoutMs = 25000 } = config;
  if (
    typeof executable !== 'string' ||
    !isAbsolute(executable) ||
    !profiles.has(profile) ||
    typeof version !== 'string' ||
    !/^[0-9]+\.[0-9]+\.[0-9]+(?:-[A-Za-z0-9.-]+)?$/.test(version) ||
    !Number.isInteger(timeoutMs) ||
    timeoutMs < 1 ||
    timeoutMs > 120000
  )
    throw new Error('Invalid veraPDF configuration');
  return {
    id: `verapdf-${profile}`,
    version,
    async validate(bytes, signal) {
      const folder = await mkdtemp(join(tmpdir(), 'imprint-verapdf-'));
      try {
        const file = join(folder, 'artifact.pdf');
        await writeFile(file, bytes, { mode: 0o600 });
        try {
          const { stdout } = await exec(
            executable,
            ['--format', 'xml', '--maxfailuresdisplayed', '1', '-f', profile, file],
            { signal, timeout: timeoutMs, maxBuffer: 4 * 1024 * 1024, encoding: 'utf8' },
          );
          return parseVeraPdfReport(stdout, profile, version, { path: file, size: bytes.byteLength });
        } catch (error) {
          const failure = error as { code?: unknown; stdout?: unknown };
          if (failure.code === 1 && typeof failure.stdout === 'string') {
            const result = parseVeraPdfReport(failure.stdout, profile, version, {
              path: file,
              size: bytes.byteLength,
            });
            if (!result.passed) return result;
          }
          throw error;
        }
      } finally {
        await rm(folder, { recursive: true, force: true });
      }
    },
  };
}
