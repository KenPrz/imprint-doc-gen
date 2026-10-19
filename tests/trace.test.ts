import test from 'node:test';
import assert from 'node:assert/strict';
import { traceTemplateInputs } from '../src/trace.ts';
import { encodeTextTemplate } from '../src/text-template.ts';
test('preflight trace reports static sources and formatters without revealing values', async () => {
  const bytes = encodeTextTemplate('txt', 'Name: {d.name:upper()} Missing: {d.missing}');
  const trace = await traceTemplateInputs(bytes, { name: 'private-name' });
  assert.equal(trace.phase, 'preflight');
  assert(trace.inspection.valid);
  assert.equal(trace.entries[0]!.status, 'present');
  assert.deepEqual(trace.entries[0]!.formatters, ['upper']);
  assert.equal(trace.entries[1]!.status, 'missing');
  assert(!JSON.stringify(trace).includes('private-name'));
});
test('trace preserves invalid-template diagnostics and marks loops contextual', async () => {
  const invalid = await traceTemplateInputs(Buffer.from('bad'), {});
  assert(!invalid.inspection.valid);
  const trace = await traceTemplateInputs(encodeTextTemplate('txt', '{d.rows[i].name}\n{d.rows[i+1].name}'), {
    rows: [{ name: 'x' }],
  });
  assert(trace.entries.length > 0);
  assert(trace.entries.every((e) => e.status === 'requires-render-context'));
});
test('formatter preflight uses renderer semantics with explicit sensitive-output opt in', async () => {
  const bytes = encodeTextTemplate(
    'txt',
    '{d.name:upper()} {d.missing:default("fallback")} {d.required:required()} {d.amount:number(2)}',
  );
  const data = { name: 'private', amount: 1234.5 };
  const redacted = await traceTemplateInputs(bytes, data);
  assert(!JSON.stringify(redacted).includes('PRIVATE'));
  assert.equal(redacted.entries[2]!.evaluation!.errorCode, 'MISSING_DATA');
  const actual = await traceTemplateInputs(bytes, data, { includeValues: true, locale: 'de-DE' });
  assert.deepEqual(
    actual.entries.map((e) => e.evaluation?.value),
    ['PRIVATE', 'fallback', undefined, '1.234,50'],
  );
  assert.equal(actual.entries[2]!.status, 'missing');
  assert.equal(actual.entries[2]!.evaluation!.status, 'error');
  const invalid = await traceTemplateInputs(encodeTextTemplate('txt', '{d.n:number()}'), { n: 'secret' });
  assert.equal(invalid.entries[0]!.evaluation!.errorCode, 'INVALID_DATA');
  assert(!JSON.stringify(invalid).includes('secret'));
  await assert.rejects(traceTemplateInputs(bytes, data, { locale: '_invalid' }));
  await assert.rejects(traceTemplateInputs(bytes, data, { includeValues: 'yes' } as any));
});
test('split Office tags are inspected as a logical paragraph and trace never executes accessors', async () => {
  const { readFile } = await import('node:fs/promises');
  const { readPackage, writePackage } = await import('../src/package.ts');
  const parts = await readPackage(await readFile('examples/invoice/invoice.docx'));
  const original = parts.get('word/document.xml')!.toString();
  assert(original.includes('{d.company}'));
  parts.set(
    'word/document.xml',
    Buffer.from(original.replace('{d.company}', '{d.com</w:t></w:r><w:r><w:t>pany}')),
  );
  const trace = await traceTemplateInputs(await writePackage(parts), { company: 'split' });
  assert(trace.entries.some((e) => e.path === 'company' && e.status === 'present'));
  let calls = 0;
  const data = {
    get company() {
      calls++;
      return 'private';
    },
  };
  await assert.rejects(traceTemplateInputs(encodeTextTemplate('txt', '{d.company}'), data));
  assert.equal(calls, 0);
  await assert.rejects(traceTemplateInputs(encodeTextTemplate('txt', '{d.company}'), [] as any));
});
test('execution trace records real repeated lookups, formatter results and failures', async () => {
  const { renderWithTrace } = await import('../src/trace.ts');
  const template = encodeTextTemplate('txt', '{d.rows[i].name:upper()}\n{d.rows[i+1]}');
  const result = await renderWithTrace(
    template,
    { rows: [{ name: 'a' }, { name: 'b' }] },
    {},
    { includeValues: true },
  );
  assert(result.ok);
  assert.equal(result.value.toString(), 'A\nB');
  assert.equal(result.trace.phase, 'execution');
  assert(result.trace.events.some((e) => e.outputValue === 'A'));
  assert(result.trace.events.some((e) => e.outputValue === 'B'));
  assert(result.trace.events.some((e) => e.scopeDepth === 1));
  const failure = await renderWithTrace(encodeTextTemplate('txt', '{d.name:required()}'), {});
  assert(!failure.ok);
  assert.equal(failure.error.code, 'MISSING_DATA');
  assert.equal(failure.trace.events[0]!.status, 'error');
  assert(failure.trace.events[0]!.missingInput);
});
test('execution traces isolate concurrent and nested contexts and visibly bound evidence', async () => {
  const { renderWithTrace } = await import('../src/trace.ts');
  const template = encodeTextTemplate('txt', '{d.name:upper()} {d.name}');
  const results = await Promise.all(
    Array.from({ length: 25 }, (_, i) =>
      renderWithTrace(template, { name: `private-${i}` }, {}, { includeValues: true }),
    ),
  );
  results.forEach((result, i) => {
    assert(result.ok);
    assert.deepEqual(
      result.trace.events.map((e) => e.outputValue),
      [`PRIVATE-${i}`, `private-${i}`],
    );
  });
  const redacted = await renderWithTrace(template, { name: 'sensitive' });
  assert(!JSON.stringify(redacted.trace).includes('sensitive'));
  const bounded = await renderWithTrace(
    template,
    { name: 'long-value' },
    {},
    { maxEvents: 1, includeValues: true, maxValueLength: 2 },
  );
  assert(bounded.ok);
  assert(bounded.trace.truncated);
  assert.equal(bounded.trace.droppedEvents, 1);
  assert.equal(bounded.trace.events.length, 1);
  assert(bounded.trace.events[0]!.valueOmitted);
  await assert.rejects(renderWithTrace(template, {}, {}, { maxEvents: -1 }));
  const { captureExecutionTrace } = await import('../src/execution-trace.ts');
  const outer = await captureExecutionTrace(async () => {
    const inner = await renderWithTrace(template, { name: 'nested' });
    assert(inner.ok);
    return true;
  });
  assert(outer.ok);
  assert.equal(outer.trace.events.length, 0);
});
test('completed execution traces do not mutate from detached asynchronous operations', async () => {
  const { captureExecutionTrace } = await import('../src/execution-trace.ts');
  const { standardLookup } = await import('../src/standard-syntax.ts');
  let finish!: () => void;
  const completed = new Promise<void>((resolve) => {
    finish = resolve;
  });
  const result = await captureExecutionTrace(async () => {
    setImmediate(() => {
      standardLookup([{ name: 'late' }], 'name', 'text');
      finish();
    });
    return 'done';
  });
  assert(result.ok);
  assert.equal(result.trace.events.length, 0);
  await completed;
  assert.equal(result.trace.events.length, 0);
});
