import { readFileSync } from 'node:fs';
const plan = JSON.parse(readFileSync(new URL('../../docs/assurance-plan.json', import.meta.url), 'utf8'));
const pkg = JSON.parse(readFileSync(new URL('../../package.json', import.meta.url), 'utf8'));
const pending = plan.remainingAcceptance;
const approval = plan.releaseApproval;
const validPending =
  Array.isArray(pending) && pending.every((item) => typeof item === 'string' && item.trim().length > 0);
// An explicit, version-bound owner decision may defer acceptance; never claim it passed.
const explicitDeferral =
  approval?.version === pkg.version &&
  approval?.disposition === 'release-with-documented-limitations' &&
  approval?.approvedBy === 'repository owner' &&
  typeof approval?.reason === 'string' &&
  approval.reason.trim().length >= 20 &&
  /^\d{4}-\d{2}-\d{2}$/.test(approval?.date ?? '') &&
  JSON.stringify(approval?.deferredAcceptance) === JSON.stringify(pending);
if (plan.releaseApproved !== true || !validPending || (pending.length > 0 && !explicitDeferral)) {
  console.error(
    'Container release blocked: require completed acceptance or explicit version-bound owner deferral. See docs/assurance-plan.json.',
  );
  process.exitCode = 1;
} else if (pending.length) {
  console.log(
    `Release ${pkg.version} approved with ${pending.length} documented acceptance limitations. Required CI jobs must still pass.`,
  );
} else console.log('Assurance acceptance completed. Required CI jobs must also pass.');
