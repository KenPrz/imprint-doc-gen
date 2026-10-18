import { test } from 'node:test';
import assert from 'node:assert/strict';
import { reviewTemplateCandidate, acceptTemplateCandidate } from '../src/template-review.ts';
import { encodeTextTemplate } from '../src/text-template.ts';
const base = encodeTextTemplate('txt', 'Hi {d.name}');
test('candidate edits require independently approved outputs; stale/forged reviews fail', async () => {
  const candidate = encodeTextTemplate('txt', 'Hello {d.name}');
  const review = await reviewTemplateCandidate(base, candidate, [
    { name: 'unicode', data: { name: '李' }, approved: Buffer.from('Hello 李') },
  ]);
  assert.equal(review.passed, true);
  assert.deepEqual(acceptTemplateCandidate(review, base), candidate);
  assert.throws(() => acceptTemplateCandidate({ ...review }, base), /passing review/);
  assert.throws(() => acceptTemplateCandidate(review, candidate), /changed/);
});
test('failed candidate cannot be accepted by changing public report', async () => {
  const review = await reviewTemplateCandidate(base, base, [
    { name: 'bad', data: { name: 'Alice' }, approved: Buffer.from('different') },
  ]);
  assert.equal(review.passed, false);
  review.passed = true;
  assert.throws(() => acceptTemplateCandidate(review, base));
});
test('candidate inspection reports field changes; empty or bad test suites rejected', async () => {
  const next = encodeTextTemplate('txt', 'Hello {d.other}');
  const review = await reviewTemplateCandidate(base, next, [
    { name: 'new', data: { other: 'Bob' }, approved: Buffer.from('Hello Bob') },
  ]);
  assert.deepEqual(review.addedFields, ['other']);
  assert.deepEqual(review.removedFields, ['name']);
  await assert.rejects(reviewTemplateCandidate(base, next, []));
  await assert.rejects(
    reviewTemplateCandidate(base, next, [
      {
        name: 'a',
        data: {
          get other() {
            throw new Error('not executed');
          },
        },
        approved: Buffer.from('Hello Bob'),
      },
    ]),
  );
});
test('candidate bytes and fixture data cannot change in flight or after review', async () => {
  const candidate = encodeTextTemplate('txt', 'Hello {d.name}'),
    expected = Buffer.from(candidate);
  const data = { name: 'Bob' },
    approved = Buffer.from('Hello Bob');
  const promise = reviewTemplateCandidate(base, candidate, [{ name: 'snapshot', data, approved }]);
  candidate.fill(0);
  approved.fill(0);
  data.name = 'Mallory';
  const review = await promise;
  assert.equal(review.passed, true);
  assert.deepEqual(acceptTemplateCandidate(review, base), expected);
  const returned = acceptTemplateCandidate(review, base);
  returned.fill(0);
  assert.deepEqual(acceptTemplateCandidate(review, base), expected);
});
test('invalid or format-changing candidate remains unapproved', async () => {
  for (const candidate of [Buffer.from('bad zip'), encodeTextTemplate('csv', '{d.name}')]) {
    const r = await reviewTemplateCandidate(base, candidate, [
      { name: 'x', data: { name: 'X' }, approved: Buffer.from('X') },
    ]);
    assert.equal(r.passed, false);
    assert.throws(() => acceptTemplateCandidate(r, base));
  }
});
