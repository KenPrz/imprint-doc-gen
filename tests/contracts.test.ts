import test from 'node:test';
import assert from 'node:assert/strict';
import {
  checkContractSchema,
  validateContract,
  contractTypeScript,
  evaluateAssertions,
  type ContractSchema,
} from '../src/contracts.ts';
const schema: ContractSchema = {
  type: 'object',
  required: ['name', 'lines'],
  additionalProperties: false,
  properties: {
    name: { type: 'string', minLength: 1, maxLength: 2 },
    lines: { type: 'array', items: { type: 'number', minimum: 0 }, maxItems: 3 },
  },
};
test('contracts validate structure, Unicode lengths, numeric bounds and escaped pointers', () => {
  assert(validateContract(schema, { name: '😀', lines: [0, 1] }).valid);
  const bad = validateContract(schema, { name: 'abc', lines: [-1], extra: true });
  assert.deepEqual(
    bad.issues.map((i) => i.code),
    ['maxLength', 'minimum', 'additionalProperties'],
  );
  assert(!validateContract(schema, { name: 'a' }).valid);
  assert.equal(validateContract({ type: 'object', required: ['a/b~c'] }, {}).issues[0]!.path, '/a~1b~0c');
  assert(!validateContract({ type: 'integer' }, 1.5).valid);
});
test('contracts reject unsupported schemas, unsafe JSON and bounded recursion', () => {
  for (const s of [
    { type: 'string', pattern: '.*' },
    { type: 'array' },
    { type: 'string', minLength: -1 },
    { type: 'number', minimum: 5, maximum: 2 },
    { type: 'object', required: ['a', 'a'] },
    { type: 'object', additionalProperties: { type: 'number' } },
  ])
    assert.throws(() => checkContractSchema(s as any));
  for (const d of [NaN, Infinity, undefined, () => 1, new Date(), Array(2)])
    assert.throws(() => validateContract(schema, d));
  const cycle: any = {};
  cycle.x = cycle;
  assert.throws(() => validateContract(schema, cycle));
  let calls = 0;
  const getter = {
    get secret() {
      calls++;
      return 1;
    },
  };
  assert.throws(() => validateContract(schema, getter));
  assert.equal(calls, 0);
  let deep: ContractSchema = { type: 'string' };
  for (let i = 0; i < 40; i++) deep = { type: 'array', items: deep };
  assert.throws(() => checkContractSchema(deep));
});
test('enum uses structural equality and object keys cannot escape generated types', () => {
  assert(validateContract({ type: 'object', enum: [{ a: 1, b: 2 }] }, { b: 2, a: 1 }).valid);
  assert(!validateContract({ type: 'number', enum: [1, 2] }, 3).valid);
  const code = contractTypeScript(schema);
  assert.match(code, /"name": string/);
  assert.match(code, /"lines": Array<number>/);
  assert.throws(() => contractTypeScript(schema, 'x; process.exit()'));
  const unusual = contractTypeScript({ type: 'object', properties: { 'a"\n;': { type: 'boolean' } } });
  assert(unusual.includes(JSON.stringify('a"\n;')));
});
test('business assertions sum safely, use explicit numeric tolerance and fail closed', () => {
  const a: any = {
    name: 'total',
    left: { sum: '/rows', field: '/cost' },
    operator: 'equals',
    right: { path: '/total' },
    tolerance: 1e-9,
  };
  assert(evaluateAssertions({ rows: [{ cost: 0.1 }, { cost: 0.2 }], total: 0.3 }, [a]).valid);
  assert(!evaluateAssertions({ rows: [{ cost: 2 }], total: 3 }, [a]).valid);
  assert(!evaluateAssertions({ rows: [{}], total: 3 }, [a]).valid);
  assert(!evaluateAssertions({ rows: [{ cost: 1e308 }, { cost: 1e308 }], total: 3 }, [a]).valid);
  assert.throws(() => evaluateAssertions({}, [{ ...a, tolerance: -1 }]));
  assert.throws(() => evaluateAssertions({}, [{ ...a, left: { eval: 'process.exit()' } }]));
  assert(
    evaluateAssertions({ n: 3 }, [
      { name: 'max', left: { path: '/n' }, operator: 'lessThanOrEqual', right: { value: 4 } },
    ]).valid,
  );
  assert(
    !evaluateAssertions({ n: '3' }, [
      { name: 'max', left: { path: '/n' }, operator: 'lessThanOrEqual', right: { value: 4 } },
    ]).valid,
  );
  assert(
    evaluateAssertions({ 'a/b': null }, [
      { name: 'null', left: { path: '/a~1b' }, operator: 'equals', right: { value: null } },
    ]).valid,
  );
});
test('schema sample is independently validated and impossible or excessive samples fail', async () => {
  const { contractSample } = await import('../src/contracts.ts');
  assert(validateContract(schema, contractSample(schema)).valid);
  for (const s of [
    { type: 'integer', minimum: 0.5, maximum: 0.7 },
    { type: 'number', enum: ['x'] },
    { type: 'string', minLength: 100000 },
    { type: 'object', required: ['x'], additionalProperties: false },
  ])
    assert.throws(() => contractSample(s as any));
  assert.equal(contractSample({ type: 'integer', maximum: -2.2 }), -3);
});
test('prototype-named properties remain own data and enums intersect their declared type', () => {
  const s = JSON.parse(
    '{"type":"object","properties":{"__proto__":{"type":"string"}},"required":["__proto__"],"additionalProperties":false}',
  );
  assert(validateContract(s, JSON.parse('{"__proto__":"ok"}')).valid);
  assert(!validateContract(s, {}).valid);
  assert.match(contractTypeScript({ type: 'number', enum: ['bad', 1] }), / & \(number\)/);
});
test('schema keyword misuse fails closed across every supported kind', () => {
  const bad: any[] = [
    true,
    false,
    null,
    {},
    [],
    { type: ['string', 'null'] },
    { type: 'bogus' },
    { type: 'string', description: 2 },
    { type: 'string', enum: [] },
    { type: 'string', enum: {} },
    ...['minLength', 'maxLength'].flatMap((k) =>
      [-1, 0.2, '2', null, true].map((v) => ({ type: 'string', [k]: v })),
    ),
    ...['minItems', 'maxItems'].flatMap((k) =>
      [-1, 0.2, '2', null, true].map((v) => ({ type: 'array', items: { type: 'null' }, [k]: v })),
    ),
    ...['minimum', 'maximum'].flatMap((k) => ['2', null, true].map((v) => ({ type: 'number', [k]: v }))),
    { type: 'object', properties: [] },
    { type: 'object', properties: null },
    { type: 'object', properties: { x: true } },
    { type: 'object', required: 'x' },
    { type: 'object', required: [2] },
    { type: 'object', additionalProperties: 0 },
    { type: 'array', items: [] },
    { type: 'array', items: false },
    ...[
      '$ref',
      '$defs',
      'allOf',
      'anyOf',
      'oneOf',
      'not',
      'if',
      'then',
      'else',
      'format',
      'pattern',
      'uniqueItems',
      'default',
      'const',
    ].map((k) => ({ type: 'string', [k]: true })),
  ];
  for (const s of bad) assert.throws(() => checkContractSchema(s), JSON.stringify(s));
});
test('deterministic independent boundary matrix verifies primitive and collection constraints', () => {
  const values: unknown[] = [
    null,
    false,
    true,
    -3,
    -1,
    -0,
    0.5,
    1,
    2,
    3,
    '',
    'a',
    '😀',
    'ab',
    'abc',
    [],
    [1],
    [1, 2],
    [1, 'x'],
    {},
    { x: 1 },
    { x: 'x' },
    { x: 1, y: 2 },
  ];
  const cases: Array<[ContractSchema, (v: any) => boolean]> = [
    [{ type: 'number', minimum: -1, maximum: 2 }, (v) => typeof v === 'number' && v >= -1 && v <= 2],
    [
      { type: 'integer', minimum: -1, maximum: 2 },
      (v) => typeof v === 'number' && Number.isSafeInteger(v) && v >= -1 && v <= 2,
    ],
    [
      { type: 'string', minLength: 1, maxLength: 2 },
      (v) => typeof v === 'string' && [...v].length >= 1 && [...v].length <= 2,
    ],
    [{ type: 'boolean' }, (v) => typeof v === 'boolean'],
    [{ type: 'null' }, (v) => v === null],
    [
      { type: 'array', items: { type: 'number' }, minItems: 1, maxItems: 2 },
      (v) => Array.isArray(v) && v.length >= 1 && v.length <= 2 && v.every((x) => typeof x === 'number'),
    ],
    [
      { type: 'object', required: ['x'], properties: { x: { type: 'number' } }, additionalProperties: false },
      (v) =>
        v !== null &&
        typeof v === 'object' &&
        !Array.isArray(v) &&
        Object.keys(v).length === 1 &&
        typeof v.x === 'number',
    ],
  ];
  for (const [s, expected] of cases)
    for (const v of values) assert.equal(validateContract(s, v).valid, expected(v), JSON.stringify({ s, v }));
});
test('validation rejects accessors and proxies before executing traps; input remains unchanged', () => {
  let traps = 0;
  const proxy = new Proxy(
    {},
    {
      getPrototypeOf() {
        traps++;
        throw new Error('trap');
      },
    },
  );
  assert.throws(() => validateContract({ type: 'object' }, proxy));
  assert.equal(traps, 0);
  const hidden = {};
  Object.defineProperty(hidden, 'value', {
    get() {
      traps++;
      return 1;
    },
  });
  assert.throws(() => validateContract({ type: 'object' }, hidden));
  assert.equal(traps, 0);
  const value = Object.freeze({ name: 'a', lines: Object.freeze([1, 2]) });
  assert(validateContract(schema, value).valid);
  const a: any = { name: 'length', left: { path: '/rows/length' }, operator: 'equals', right: { value: 0 } };
  assert(!evaluateAssertions({ rows: [] }, [a]).valid);
  assert(!evaluateAssertions({ rows: [1] }, [{ ...a, left: { path: '/rows/00' } }]).valid);
});
test('TypeScript declarations reject reserved type names', () => {
  for (const name of [
    'string',
    'number',
    'boolean',
    'null',
    'true',
    'false',
    'undefined',
    'await',
    'import',
    'void',
    'never',
    'any',
    'unknown',
  ])
    assert.throws(() => contractTypeScript({ type: 'string' }, name));
});
