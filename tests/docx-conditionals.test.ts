import test from 'node:test';
import assert from 'node:assert/strict';
import { compileStory, renderStory } from '../src/docx-story.ts';
import { standardBindings } from '../src/standard-syntax.ts';
import { parseXml, descendants, ns } from '../src/xml.ts';

const p = (text: string) => `<w:p><w:r><w:t>${text}</w:t></w:r></w:p>`;
const open = '{d.flag:ifNEM():showBegin}',
  close = '{d.flag:showEnd}';
const source = (body: string) => `<w:document xmlns:w="${ns.word}"><w:body>${body}</w:body></w:document>`;
async function output(body: string, data: Record<string, unknown>) {
  return renderStory(
    compileStory('word/document.xml', source(body)),
    data,
    { iterations: 0 },
    async () => [],
    () => '1',
  );
}
function paragraphs(xml: string) {
  return descendants(parseXml(xml, 'test'), ns.word, 'p').map((n) =>
    descendants(n, ns.word, 't')
      .map((t) => t.textContent ?? '')
      .join(''),
  );
}
// These emptiness and retained-closing-paragraph results were observed from
// Carbone 3.8.2 on the real PMS confirmation-letter template, including false/0.
const cases: [string, unknown, boolean][] = [
  ['null', null, false],
  ['missing', undefined, false],
  ['empty string', '', false],
  ['empty array', [], false],
  ['empty object', {}, false],
  ['zero', 0, true],
  ['false', false, true],
  ['true', true, true],
  ['whitespace', ' ', true],
  ['text', 'yes', true],
  ['array', [null], true],
  ['object', { a: null }, true],
];
for (const [name, flag, shown] of cases)
  test(`DOCX ifNEM block: ${name}`, async () => {
    const data = flag === undefined ? { name: 'Ada' } : { flag, name: 'Ada' };
    const result = await output(
      p('Before') + p(open + 'Guarantee {d.name}') + p('Signature') + p(close) + p('After'),
      data,
    );
    assert.deepEqual(
      paragraphs(result),
      shown ? ['Before', 'Guarantee Ada', 'Signature', '', 'After'] : ['Before', '', 'After'],
    );
    assert(!result.includes('showBegin') && !result.includes('showEnd'));
  });
test('conditional split runs preserve formatting and following data; suppressed children are not evaluated', async () => {
  const body =
    '<w:p><w:r><w:rPr><w:b/></w:rPr><w:t>{d.flag:ifNEM():show</w:t></w:r><w:r><w:t>Begin}Guarantee</w:t></w:r></w:p>' +
    p('{d.bad:number()}') +
    p(close) +
    p('{d.tail}');
  const hidden = await output(body, { flag: null, bad: 'invalid', tail: 'retained' });
  assert.deepEqual(paragraphs(hidden), ['', 'retained']);
  await assert.rejects(output(body, { flag: true, bad: 'invalid' }), /number requires/);
  const shown = await output(body, { flag: true, bad: 3, tail: 'retained' });
  assert(shown.includes('<w:b/>'));
  assert.deepEqual(paragraphs(shown), ['Guarantee', '3', '', 'retained']);
});
test('sequential blocks and table content retain root data scope', async () => {
  const body =
    p(open + 'Section') +
    '<w:tbl><w:tr><w:tc>' +
    p('{d.name}') +
    '</w:tc></w:tr></w:tbl>' +
    p(close) +
    p('{d.other:ifNEM():showBegin}Other') +
    p('{d.other:showEnd}');
  assert.deepEqual(paragraphs(await output(body, { flag: 1, name: 'Root', other: '' })), [
    'Section',
    'Root',
    '',
    '',
  ]);
});
const invalid: [string, string][] = [
  ['unclosed', p(open + 'Text')],
  ['unmatched', p(close)],
  ['different field', p(open + 'Text') + p('{d.other:showEnd}')],
  ['inline', p(open + 'Text' + close)],
  ['prefix', p('Before ' + open + 'Text') + p(close)],
  ['closing content', p(open + 'Text') + p(close + 'After')],
  ['nested', p(open + 'Text') + p('{d.other:ifNEM():showBegin}Inner') + p('{d.other:showEnd}') + p(close)],
  ['cross-container', p(open + 'Text') + '<w:tbl><w:tr><w:tc>' + p(close) + '</w:tc></w:tr></w:tbl>'],
  ['repeat mixing', p(open + 'Text') + p(close) + p('{d.rows[i].name}') + p('{d.rows[i+1].name}')],
  [
    'native repeat mixing',
    p(open + 'Text') + p(close) + p('{{#each rows}}') + p('{{name}}') + p('{{/each}}'),
  ],
  [
    'range anchors',
    p(open + 'Text') + '<w:p><w:bookmarkStart w:id="1"/><w:r><w:t>Range</w:t></w:r></w:p>' + p(close),
  ],
  ['unsupported condition', p('{d.flag:ifEQ(1):showBegin}Text') + p(close)],
  ['missing predicate', p('{d.flag:showBegin}Text') + p(close)],
  ['condition on end', p(open + 'Text') + p('{d.flag:ifNEM():showEnd}')],
];
for (const [name, body] of invalid)
  test(`conditional rejects ${name}`, () =>
    assert.throws(() => compileStory('word/document.xml', source(body))));
test('structural condition is rejected in other formats and ordinary scalar formatting', () => {
  for (const part of ['text', 'xl/worksheets/sheet1.xml', 'ppt/slides/slide1.xml', 'content.xml'])
    assert.throws(() => standardBindings(open, part));
  assert.throws(() => standardBindings('{d.flag:ifNEM()}', 'word/document.xml'));
});
