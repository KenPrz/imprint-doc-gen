/** Downloads pinned public reports into a caller-selected scratch folder. No report
 * files are redistributed. Expected metadata was independently read with olefile. */
import { mkdir, readFile, writeFile } from 'node:fs/promises';
import { resolve, join } from 'node:path';
import { createHash } from 'node:crypto';
import assert from 'node:assert/strict';
import { inspectCrystalArtifact } from '../../src/crystal-artifact.ts';

const destination = process.argv[2];
if (!destination)
  throw new Error('Usage: node scripts/assurance/crystal-sample-check.ts <scratch-directory>');
const dir = resolve(destination);
await mkdir(dir, { recursive: true });
const oracle = JSON.parse(
  await readFile(new URL('../../tests/fixtures/crystal/directory-oracle.json', import.meta.url), 'utf8'),
);
for (const sample of oracle.samples) {
  const url = `https://raw.githubusercontent.com/kleinjoshuaa/Crystal-Reports/${oracle.sourceCommit}/${encodeURIComponent(sample.filename)}`;
  const response = await fetch(url, { signal: AbortSignal.timeout(30_000) });
  if (!response.ok) throw new Error(`Sample download failed: ${response.status}`);
  const data = Buffer.from(await response.arrayBuffer());
  assert.equal(data.length, sample.bytes);
  assert.equal(createHash('sha256').update(data).digest('hex'), sample.sha256);
  await writeFile(join(dir, sample.filename), data);
  const actual = inspectCrystalArtifact(data);
  const normalize = (entries: { path: string[]; kind: string; declaredBytes: number }[]) =>
    entries
      .map(({ path, kind, declaredBytes }) => ({ path, kind, declaredBytes }))
      .sort((a, b) => JSON.stringify(a.path).localeCompare(JSON.stringify(b.path), 'en'));
  assert.deepEqual(normalize(actual.entries), normalize(sample.entries));
  assert.equal(actual.crystalCandidate, true);
  assert.equal(actual.automaticConversion, false);
  console.log(
    `${sample.filename}: ${actual.entries.length} directory entries matched independent olefile oracle; SHA-256 verified`,
  );
}
