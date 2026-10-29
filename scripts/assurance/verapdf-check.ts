import { mkdtemp, writeFile, readFile, rm, mkdir } from 'node:fs/promises';
import { tmpdir } from 'node:os';
import { join } from 'node:path';
import assert from 'node:assert/strict';
import { execFileSync } from 'node:child_process';
import { PDFDocument } from 'pdf-lib';
import { veraPdfValidator } from '../../src/verapdf.ts';
import { artifactDigest, validateDelivery } from '../../src/delivery.ts';
// Use the published stable release; upstream prunes development images from GHCR.
const image = 'verapdf/cli:v1.30.2@sha256:d5ee329657cf9bc4b2400392dd54c7d0a0ce9980ff6fa2da5590eebeec007cdb';
// Acquisition failures should be visible, and downloads must not consume a document's validation deadline.
execFileSync('docker', ['pull', image], { stdio: 'inherit', timeout: 120000 });
const directory = await mkdtemp(join(tmpdir(), 'imprint-verapdf-check-'));
try {
  const executable = join(directory, 'verapdf');
  await writeFile(
    executable,
    `#!/bin/sh\nfor argument do file="$argument"; done\nfolder=$(dirname "$file")\nexec docker run --rm --network none --read-only --tmpfs /tmp --memory 512m --cpus 1 --pids-limit 128 --user "$(id -u):$(id -g)" -v "$folder:$folder:ro" ${image} "$@"\n`,
    { mode: 0o700 },
  );
  const validator = veraPdfValidator({ executable, profile: '1b', version: '1.30.2' });
  const fixture = await readFile('tests/fixtures/assurance/pdfa-1b-pass.pdf');
  const provenance = JSON.parse(await readFile('tests/fixtures/assurance/provenance.json', 'utf8'));
  assert.equal(artifactDigest(fixture), provenance.sha256);
  const positive = await validateDelivery(fixture, [validator]);
  assert.equal(positive.evidence.status, 'approved', JSON.stringify(positive.evidence));
  const pdf = await PDFDocument.create();
  pdf.addPage().drawText('Deliberately not PDF/A');
  const negative = await validateDelivery(await pdf.save(), [validator]);
  assert.equal(negative.evidence.status, 'quarantined');
  assert.equal(negative.artifact, undefined);
  assert.match(negative.evidence.checks[0].details, /1b failed/);
  const broken = await validateDelivery(Buffer.from('%PDF-broken'), [validator]);
  assert.equal(broken.evidence.status, 'quarantined');
  await mkdir('artifacts/assurance', { recursive: true });
  await writeFile(
    'artifacts/assurance/verapdf-validation.json',
    JSON.stringify(
      { image, positive: positive.evidence, negative: negative.evidence, malformed: broken.evidence },
      null,
      2,
    ),
  );
  console.log('Independent veraPDF positive, negative and malformed checks passed');
} finally {
  await rm(directory, { recursive: true, force: true });
}
