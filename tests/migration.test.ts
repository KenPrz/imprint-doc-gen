import test from 'node:test';
import assert from 'node:assert/strict';
import {
  migrateCarboneText,
  inventoryCrystalMetadata,
  compileCarboneText,
} from '../src/template-migration.ts';
import { ComponentRegistry } from '../src/components.ts';
import { compileText } from '../src/text-template.ts';

test('Carbone scalar migration retains source and renders official basic example semantics', async () => {
  const source = 'Hello {d.firstname} {d.lastname} !';
  const result = migrateCarboneText(source);
  assert.equal(result.compatible, true);
  assert.equal(result.source, source);
  assert.deepEqual(result.fields, ['firstname', 'lastname']);
  assert.equal(
    (await compileText('txt', result.source).render({ firstname: 'John', lastname: 'Doe' })).toString(),
    'Hello John Doe !',
  );
  assert.deepEqual(migrateCarboneText('{d.a.b} {d.a.b}').fields, ['a.b']);
});

for (const tag of [
  '{d.rows[i].name}',
  '{d.rows[i+1]}',
  '{d.name:upperCase}',
  '{c.value}',
  '{#alias}',
  '{o.lang=en}',
  '{t(hello)}',
  '{d.__proto__.name}',
  '{d.constructor}',
  '{d.name',
  '{d.name}}',
  '{{d.name}}',
  '{d.rows[0]}',
  '{d.parent..name}',
  '{d. name}',
]) {
  test(`Migration flags unsupported source without silently changing it: ${tag}`, () => {
    const source = `Prefix ${tag} suffix`,
      result = migrateCarboneText(source);
    assert.equal(result.compatible, false);
    assert.equal(result.source, source);
    assert(result.issues.length > 0);
    for (const issue of result.issues)
      assert.equal(source.slice(issue.offset, issue.offset + issue.tag.length), issue.tag);
  });
}
test('Migration rejects source/tag budget overflow and accepts empty literal', () => {
  assert.throws(() => migrateCarboneText('a'.repeat(1000001)), /1000000/);
  assert.throws(() => migrateCarboneText('{}'.repeat(10001)), /budget/);
  assert.equal(migrateCarboneText('').compatible, true);
});

test('Crystal inventory requires explicit interchange and produces all manual conversion tasks', () => {
  const metadata = {
    schema: 'imprint-crystal-inventory-v1',
    name: 'Invoice',
    fields: ['Orders.Total'],
    formulas: [{ name: 'Tax', expression: '{Orders.Total} * 0.12' }],
    groups: ['Customer'],
    subreports: ['Payments'],
  };
  const report = inventoryCrystalMetadata(metadata);
  assert.equal(report.automaticConversion, false);
  assert.deepEqual(
    report.tasks.map((t) => t.kind),
    ['field', 'formula', 'group', 'subreport'],
  );
  metadata.fields.push('Added');
  assert.equal(report.metadata.fields.length, 1);
  for (const invalid of [
    Buffer.from('RPT'),
    'native text export',
    {},
    { ...metadata, raw: 'ignored' },
    { ...metadata, fields: new Array(2) },
    { ...metadata, formulas: [{ name: 'x' }] },
    { ...metadata, formulas: new Array(2) },
  ]) {
    assert.throws(() => inventoryCrystalMetadata(invalid), /Invalid Crystal/);
  }
});
const components = [
  { id: 'address', version: '1', source: '123 Main Street' },
  { id: 'address', version: '2', source: '456 Main Street' },
  { id: 'header', version: '1', source: 'Company: {{> address@1 }} / {d.name}' },
];
test('Components expand pinned nested versions preserving fields and deterministic lock', async () => {
  const registry = new ComponentRegistry(components);
  const expanded = registry.expand('{{>header@1}}\n{{>address@1}}');
  assert.equal(expanded.source, 'Company: 123 Main Street / {d.name}\n123 Main Street');
  assert.deepEqual(
    expanded.lock.map((c) => `${c.id}@${c.version}`),
    ['address@1', 'header@1'],
  );
  assert(expanded.lock.every((c) => /^[a-f0-9]{64}$/.test(c.sha256)));
  assert.equal(
    (await compileText('txt', expanded.source).render({ name: 'Example' })).toString(),
    'Company: 123 Main Street / Example\n123 Main Street',
  );
  assert.deepEqual(registry.expand('{{>address@1}}').lock, expanded.lock.slice(0, 1));
});
test('Components impact is transitive, exact-version and sorted', () => {
  const registry = new ComponentRegistry(components);
  const templates = { z: '{{>header@1}}', a: '{{>address@1}}', other: '{{>address@2}}', static: 'Text' };
  assert.deepEqual(registry.impact('address', '1', templates), ['a', 'z']);
  assert.deepEqual(registry.impact('address', '2', templates), ['other']);
  assert.throws(() => registry.impact('address', '3', templates), /Missing/);
});
test('Registry snapshots supplied component objects', () => {
  const items = [{ id: 'x', version: '1', source: 'before' }],
    registry = new ComponentRegistry(items);
  items[0]!.source = 'after';
  assert.equal(registry.expand('{{>x@1}}').source, 'before');
});
test('Components reject cycles, missing, duplicate and malformed identities/references', () => {
  for (const items of [
    [{ id: 'x', version: '1', source: '{{>x@1}}' }],
    [
      { id: 'x', version: '1', source: '{{>y@1}}' },
      { id: 'y', version: '1', source: '{{>x@1}}' },
    ],
    [{ id: 'x', version: '1', source: '{{>y@1}}' }],
    [components[0]!, components[0]!],
    [{ id: '../x', version: '1', source: '' }],
    [{ id: 'x', version: 'latest*', source: '' }],
    [{ id: 'x', version: '1', source: '{{> address }}' }],
  ])
    assert.throws(() => new ComponentRegistry(items));
  const registry = new ComponentRegistry([]);
  for (const text of ['{{>missing@1}}', '{{>x@}}', '{{>x@1']) assert.throws(() => registry.expand(text));
});
test('Component graph and expansion budgets stop deeply nested and exponential work', () => {
  assert.throws(
    () =>
      new ComponentRegistry(
        Array.from({ length: 66 }, (_, i) => ({
          id: `n${i}`,
          version: '1',
          source: i === 65 ? 'end' : `{{>n${i + 1}@1}}`,
        })),
      ),
    /depth/,
  );
  const registry = new ComponentRegistry(
    Array.from({ length: 15 }, (_, i) => ({
      id: `n${i}`,
      version: '1',
      source: i === 14 ? 'end' : `{{>n${i + 1}@1}}`.repeat(2),
    })),
  );
  assert.throws(() => registry.expand('{{>n0@1}}'), /count/);
  const large = new ComponentRegistry([{ id: 'x', version: '1', source: 'x'.repeat(600000) }]);
  assert.throws(() => large.expand('{{>x@1}}{{>x@1}}'), /budget/);
});

test('Opt-in string case migration matches independent Carbone oracle with explicit contract', async () => {
  const { readFile } = await import('node:fs/promises');
  const oracle = JSON.parse(
    await readFile(new URL('./fixtures/migration/carbone-oracle.json', import.meta.url), 'utf8'),
  );
  for (const sample of oracle.cases) {
    const isCase = sample.source.includes(':');
    if (isCase && typeof sample.data.name !== 'string') continue;
    const result = migrateCarboneText(sample.source, { stringFields: ['name'], locale: 'en-US' });
    assert(result.compatible, sample.name);
    assert.equal(result.originalSource, sample.source);
    assert.equal(
      (await compileText('txt', result.source).render(sample.data, { locale: 'en-US' })).toString(),
      sample.output,
      sample.name,
    );
    if (isCase) assert.deepEqual(result.requirements, { stringFields: ['name'], locale: 'en-US' });
  }
});
test('Case migration is atomic, bounded and never assumes compatible coercion or locale', () => {
  assert.equal(migrateCarboneText('{d.name:upperCase}', { stringFields: ['name'] }).compatible, false);
  assert.equal(migrateCarboneText('{d.name:upperCase}', { locale: 'en-US' }).compatible, false);
  assert.equal(
    migrateCarboneText('{d.name:upperCase()}', { stringFields: ['name'], locale: 'en-US' }).source,
    '{d.name:upper()}',
  );
  const source = '{d.name:upperCase} {d.value:formatN(2)}';
  assert.equal(migrateCarboneText(source, { stringFields: ['name'], locale: 'en-US' }).source, source);
  assert.equal(
    migrateCarboneText(`{d.name${':upperCase'.repeat(17)}}`, { stringFields: ['name'], locale: 'en-US' })
      .compatible,
    false,
  );
});

test('Compiled Carbone migration enforces actual string contracts and locale on every render', async () => {
  const names = ['person.name'];
  const compiled = compileCarboneText('{d.person.name:upperCase}', { stringFields: names, locale: 'en-US' });
  names.length = 0;
  assert.equal((await compiled.render({ person: { name: 'i' } })).toString(), 'I');
  for (const name of [123, false, {}, []]) {
    await assert.rejects(compiled.render({ person: { name } }), /requires string/);
  }
  await assert.rejects(compiled.render({ person: { name: 'i' } }, { locale: 'tr' }), /en-US/);
  for (const data of [{}, { person: {} }, { person: { name: null } }])
    assert.equal((await compiled.render(data)).toString(), '');
  assert.throws(() => compileCarboneText('{d.name:formatN(2)}'), /Unsupported/);
  assert.throws(() => compileCarboneText('{d.name:upperCase}'), /Unsupported/);
});

test('Crystal interchange rejects executable getters without invoking them', () => {
  let accessed = false;
  const metadata = Object.defineProperty({}, 'name', {
    enumerable: true,
    get() {
      accessed = true;
      return 'Report';
    },
  });
  assert.throws(() => inventoryCrystalMetadata(metadata), /Invalid Crystal/);
  assert.equal(accessed, false);
});
