import { test } from 'node:test';
import assert from 'node:assert/strict';
import { mkdtemp, mkdir, writeFile, readFile, rm } from 'node:fs/promises';
import { join } from 'node:path';
import { tmpdir } from 'node:os';
import { spawnSync } from 'node:child_process';
test('container publication gate requires complete acceptance or exact version-bound owner deferral', async () => {
  const folder = await mkdtemp(join(tmpdir(), 'imprint-release-gate-'));
  try {
    await mkdir(join(folder, 'scripts', 'assurance'), { recursive: true });
    await mkdir(join(folder, 'docs'));
    await writeFile(join(folder, 'package.json'), JSON.stringify({ version: '1.0.0' }));
    const script = join(folder, 'scripts', 'assurance', 'release-gate.mjs');
    await writeFile(script, await readFile('scripts/assurance/release-gate.mjs'));
    for (const plan of [
      {},
      { releaseApproved: false, remainingAcceptance: [] },
      { releaseApproved: 'true', remainingAcceptance: [] },
      { releaseApproved: true, remainingAcceptance: ['customer baselines'] },
      { releaseApproved: true, remainingAcceptance: null },
    ]) {
      await writeFile(join(folder, 'docs', 'assurance-plan.json'), JSON.stringify(plan));
      assert.equal(spawnSync(process.execPath, [script]).status, 1);
    }
    await writeFile(
      join(folder, 'docs', 'assurance-plan.json'),
      JSON.stringify({ releaseApproved: true, remainingAcceptance: [] }),
    );
    assert.equal(spawnSync(process.execPath, [script]).status, 0);
    const pending = ['desktop Office acceptance'];
    const approval = {
      version: '1.0.0',
      date: '2026-10-19',
      approvedBy: 'repository owner',
      disposition: 'release-with-documented-limitations',
      reason: 'Explicit owner release instruction after disclosure',
      deferredAcceptance: pending,
    };
    for (const change of [
      { version: '1.0.1' },
      { deferredAcceptance: [] },
      { reason: '' },
      { approvedBy: 'automation' },
      { disposition: 'passed' },
      { date: '' },
    ]) {
      await writeFile(
        join(folder, 'docs', 'assurance-plan.json'),
        JSON.stringify({
          releaseApproved: true,
          remainingAcceptance: pending,
          releaseApproval: { ...approval, ...change },
        }),
      );
      assert.equal(spawnSync(process.execPath, [script]).status, 1);
    }
    await writeFile(
      join(folder, 'docs', 'assurance-plan.json'),
      JSON.stringify({ releaseApproved: true, remainingAcceptance: pending, releaseApproval: approval }),
    );
    const accepted = spawnSync(process.execPath, [script], { encoding: 'utf8' });
    assert.equal(accepted.status, 0);
    assert.match(accepted.stdout, /documented acceptance limitations/);
  } finally {
    await rm(folder, { recursive: true, force: true });
  }
});
