import test from 'node:test';
import assert from 'node:assert/strict';
import { createHash } from 'node:crypto';
import { assembleDocxComponents, type DocxComponent } from '../src/docx-components.ts';
import { writePackage, readPackage } from '../src/package.ts';
import { ns, parseXml, descendants } from '../src/xml.ts';
const paragraph = (s: string) => `<w:p><w:r><w:t>${s}</w:t></w:r></w:p>`;
async function fixture(body: string, extras: Record<string, string | Buffer> = {}) {
  return writePackage(
    new Map(
      Object.entries({
        '[Content_Types].xml': `<Types xmlns="${ns.contentTypes}"><Default Extension="rels" ContentType="application/vnd.openxmlformats-package.relationships+xml"/><Default Extension="xml" ContentType="application/xml"/><Override PartName="/word/document.xml" ContentType="application/vnd.openxmlformats-officedocument.wordprocessingml.document.main+xml"/></Types>`,
        '_rels/.rels': `<Relationships xmlns="${ns.packageRelationship}"><Relationship Id="root" Type="${ns.relationship}/officeDocument" Target="word/document.xml"/></Relationships>`,
        'word/document.xml': `<w:document xmlns:w="${ns.word}" xmlns:r="${ns.relationship}" xmlns:a="${ns.drawing}" xmlns:wp="${ns.wordDrawing}"><w:body>${body}<w:sectPr/></w:body></w:document>`,
        ...extras,
      }).map(([key, value]) => [key, Buffer.from(value)]),
    ),
  );
}
async function component(body: string, extras: Record<string, string | Buffer> = {}): Promise<DocxComponent> {
  const document = await fixture(body, extras);
  return {
    id: 'clause',
    version: '1',
    document,
    sha256: createHash('sha256').update(document).digest('hex'),
  };
}
const styles = (s: string) => `<w:styles xmlns:w="${ns.word}">${s}</w:styles>`;
const relationships = (s: string) => `<Relationships xmlns="${ns.packageRelationship}">${s}</Relationships>`;
const anchor = paragraph('{{>clause@1}}');
test('DOCX assembly inserts repeated styled paragraphs and tables, preserving pinned lock and isolating style IDs', async () => {
  const item = await component(
    `<w:p><w:pPr><w:pStyle w:val="Heading"/></w:pPr><w:r><w:rPr><w:b/></w:rPr><w:t>Clause</w:t></w:r></w:p><w:tbl><w:tr><w:tc>${paragraph('Cell')}</w:tc></w:tr></w:tbl>`,
    {
      'word/styles.xml': styles(
        '<w:style w:type="paragraph" w:default="1" w:styleId="Normal"><w:rPr><w:color w:val="FF0000"/></w:rPr></w:style><w:style w:type="paragraph" w:styleId="Heading"><w:basedOn w:val="Normal"/></w:style>',
      ),
    },
  );
  const host = await fixture(anchor + anchor, { 'word/styles.xml': styles('<w:style w:styleId="Normal"/>') });
  const first = await assembleDocxComponents(host, [item]),
    second = await assembleDocxComponents(host, [item]);
  assert.deepEqual(first, second);
  assert.deepEqual(first.lock, [{ id: 'clause', version: '1', sha256: item.sha256, occurrences: 2 }]);
  const parts = await readPackage(first.document),
    document = parseXml(parts.get('word/document.xml')!.toString(), 'test');
  assert.equal(descendants(document, ns.word, 'tbl').length, 2);
  assert.equal(descendants(document, ns.word, 'b').length, 2);
  assert.deepEqual(
    descendants(document, ns.word, 'pStyle').map((e) => e.getAttributeNS(ns.word, 'val')),
    ['imprint_1_1', 'imprint_1_0', 'imprint_2_1', 'imprint_2_0'],
  );
  assert.equal(descendants(document, ns.word, 'sectPr').length, 1);
  assert.equal(
    descendants(parseXml(parts.get('word/styles.xml')!.toString(), 's'), ns.word, 'style').length,
    5,
  );
});
test('DOCX assembly copies actual PNG bytes with distinct occurrence relationships and drawing IDs', async () => {
  const png = Buffer.from(
    'iVBORw0KGgoAAAANSUhEUgAAABAAAAAQCAIAAACQkWg2AAAACXBIWXMAAAPoAAAD6AG1e1JrAAAAGUlEQVQokWNgSJtJGmIY1TBzNJTShmnSAAD42/8BAFSM9QAAAABJRU5ErkJggg==',
    'base64',
  );
  const item = await component(
    '<w:p><w:r><w:drawing><wp:inline><wp:docPr id="1" name="pixel"/><a:graphic><a:blip r:embed="pic"/></a:graphic></wp:inline></w:drawing></w:r></w:p>',
    {
      'word/_rels/document.xml.rels': relationships(
        `<Relationship Id="pic" Type="${ns.relationship}/image" Target="media/pixel.png"/>`,
      ),
      'word/media/pixel.png': png,
    },
  );
  const output = await assembleDocxComponents(await fixture(anchor + anchor), [item]),
    parts = await readPackage(output.document);
  const doc = parseXml(parts.get('word/document.xml')!.toString(), 'test');
  assert.deepEqual(
    descendants(doc, ns.wordDrawing, 'docPr').map((e) => e.getAttribute('id')),
    ['1', '2'],
  );
  const images = [...parts].filter(([name]) => name.startsWith('word/media/'));
  assert.equal(images.length, 2);
  assert.ok(images.every(([, data]) => data.equals(png)));
  assert.equal(
    new Set(descendants(doc, ns.drawing, 'blip').map((e) => e.getAttributeNS(ns.relationship, 'embed'))).size,
    2,
  );
});
for (const [name, body, extras, error] of [
  ['nested references', anchor, {}, /Nested/],
  [
    'missing numbering',
    '<w:p><w:pPr><w:numPr><w:numId w:val="99"/></w:numPr></w:pPr></w:p>',
    {},
    /Missing DOCX part/,
  ],
  ['fields', '<w:p><w:r><w:fldChar/></w:r></w:p>', {}, /fldChar/],
  ['notes', '<w:p><w:r><w:footnoteReference w:id="1"/></w:r></w:p>', {}, /footnoteReference/],
  ['theme', '<w:p><w:r><w:rPr><w:rFonts w:asciiTheme="minorHAnsi"/></w:rPr></w:r></w:p>', {}, /theme/],
  [
    'defaults',
    paragraph('x'),
    { 'word/styles.xml': styles('<w:docDefaults><w:rPrDefault/></w:docDefaults>') },
    /defaults/,
  ],
  ['missing styles', '<w:p><w:pPr><w:pStyle w:val="absent"/></w:pPr></w:p>', {}, /Missing.*style/],
  [
    'style cycles',
    paragraph('x'),
    {
      'word/styles.xml': styles(
        '<w:style w:styleId="a"><w:basedOn w:val="b"/></w:style><w:style w:styleId="b"><w:basedOn w:val="a"/></w:style>',
      ),
    },
    /Cyclic/,
  ],
  [
    'external relationship',
    paragraph('x'),
    {
      'word/_rels/document.xml.rels': relationships(
        `<Relationship Id="a" TargetMode="External" Type="${ns.relationship}/image" Target="https://example.com/x"/>`,
      ),
    },
    /External/,
  ],
  [
    'missing image',
    '<w:p><w:r><a:blip r:embed="a"/></w:r></w:p>',
    {
      'word/_rels/document.xml.rels': relationships(
        `<Relationship Id="a" Type="${ns.relationship}/image" Target="media/absent.png"/>`,
      ),
    },
    /Missing.*image/,
  ],
  [
    'hyperlinks',
    '<w:p><w:hyperlink r:id="a"/></w:p>',
    {
      'word/_rels/document.xml.rels': relationships(
        `<Relationship Id="a" Type="${ns.relationship}/hyperlink" Target="other.xml"/>`,
      ),
    },
    /Only embedded/,
  ],
  ['internal section', '<w:p><w:pPr><w:sectPr/></w:pPr></w:p>', {}, /internal sections/],
] as const)
  test(`DOCX assembly rejects ${name}`, async () => {
    const item = await component(body, extras);
    await assert.rejects(assembleDocxComponents(await fixture(anchor), [item]), error);
  });
test('DOCX assembly rejects checksum mismatch, duplicate identity, malformed marker, missing entry and nested marker locations', async () => {
  const item = await component(paragraph('x')),
    host = await fixture(anchor);
  await assert.rejects(assembleDocxComponents(host, [{ ...item, sha256: '0'.repeat(64) }]), /checksum/);
  await assert.rejects(assembleDocxComponents(host, [item, item]), /Duplicate/);
  await assert.rejects(assembleDocxComponents(host, []), /Missing/);
  for (const body of [
    paragraph('before {{>clause@1}}'),
    `<w:tbl><w:tr><w:tc>${anchor}</w:tc></w:tr></w:tbl>`,
    paragraph('{{>clause}}'),
  ])
    await assert.rejects(assembleDocxComponents(await fixture(body), [item]), /whole top-level/);
});
test('DOCX assembly enforces expansion budget and leaves input buffers unchanged', async () => {
  const item = await component(paragraph('x')),
    host = await fixture(anchor.repeat(129)),
    before = Buffer.from(host),
    source = Buffer.from(item.document);
  await assert.rejects(assembleDocxComponents(host, [item]), /128 occurrences/);
  assert.deepEqual(host, before);
  assert.deepEqual(item.document, source);
});
test('independent python-docx fixture assembles styled headings, tables, numbered lists and real drawings', async () => {
  const { readFile } = await import('node:fs/promises');
  const host = await readFile(new URL('./fixtures/docx-components/host.docx', import.meta.url));
  const document = await readFile(new URL('./fixtures/docx-components/clause.docx', import.meta.url));
  const output = await assembleDocxComponents(host, [
    { id: 'clause', version: '1', document, sha256: createHash('sha256').update(document).digest('hex') },
  ]);
  const parts = await readPackage(output.document),
    doc = parseXml(parts.get('word/document.xml')!.toString(), 'real');
  assert.equal(descendants(doc, ns.word, 'tbl').length, 2);
  assert.equal(descendants(doc, ns.wordDrawing, 'docPr').length, 2);
  assert.equal(descendants(doc, ns.word, 't').filter((e) => e.textContent === 'First requirement').length, 2);
  const stylesDoc = parseXml(parts.get('word/styles.xml')!.toString(), 'styles');
  const nums = descendants(
    parseXml(parts.get('word/numbering.xml')!.toString(), 'numbering'),
    ns.word,
    'num',
  );
  const ids = new Set(nums.map((n) => n.getAttributeNS(ns.word, 'numId')));
  assert.equal(ids.size, nums.length);
  for (const ref of descendants(stylesDoc, ns.word, 'numId'))
    assert.ok(ids.has(ref.getAttributeNS(ns.word, 'val')));
});
test('DOCX assembly rejects corrupt images before returning a package', async () => {
  const item = await component('<w:p><w:r><a:blip r:embed="pic"/></w:r></w:p>', {
    'word/_rels/document.xml.rels': relationships(
      `<Relationship Id="pic" Type="${ns.relationship}/image" Target="media/pixel.png"/>`,
    ),
    'word/media/pixel.png': Buffer.from('89504e470d0a1a0a', 'hex'),
  });
  await assert.rejects(assembleDocxComponents(await fixture(anchor), [item]));
});
test('DOCX assembly rejects missing abstract numbering and duplicate styles', async () => {
  const body = '<w:p><w:pPr><w:numPr><w:numId w:val="1"/></w:numPr></w:pPr></w:p>';
  const item = await component(body, {
    'word/numbering.xml': `<w:numbering xmlns:w="${ns.word}"><w:num w:numId="1"><w:abstractNumId w:val="7"/></w:num></w:numbering>`,
  });
  await assert.rejects(assembleDocxComponents(await fixture(anchor), [item]), /Missing.*abstract/);
  const duplicate = await component(paragraph('x'), {
    'word/styles.xml': styles('<w:style w:styleId="a"/><w:style w:styleId="a"/>'),
  });
  await assert.rejects(assembleDocxComponents(await fixture(anchor), [duplicate]), /duplicate.*style/);
});
