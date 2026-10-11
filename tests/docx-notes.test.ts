import test from 'node:test';
import assert from 'node:assert/strict';
import { readFile } from 'node:fs/promises';
import { readPackage, writePackage } from '../src/package.ts';
import { render } from '../src/index.ts';
import { ns } from '../src/xml.ts';
test('Document note and comment stories render fields while preserving numbering definitions', async () => {
  const parts = await readPackage(await readFile('examples/invoice/invoice.docx'));
  const relationships = 'word/_rels/document.xml.rels';
  for (const kind of ['footnotes', 'endnotes', 'comments']) {
    const singular = kind === 'comments' ? 'comment' : kind.slice(0, -1);
    parts.set(
      `word/${kind}.xml`,
      Buffer.from(
        `<w:${kind} xmlns:w="${ns.word}"><w:${singular} w:id="1"><w:p><w:r><w:t>{d.note}</w:t></w:r></w:p></w:${singular}></w:${kind}>`,
      ),
    );
    parts.set(
      relationships,
      Buffer.from(
        parts
          .get(relationships)!
          .toString()
          .replace(
            '</Relationships>',
            `<Relationship Id="${kind}" Type="${ns.relationship}/${kind}" Target="${kind}.xml"/></Relationships>`,
          ),
      ),
    );
  }
  const numbering = Buffer.from(
    `<w:numbering xmlns:w="${ns.word}"><w:abstractNum w:abstractNumId="0"><w:lvl w:ilvl="0"><w:numFmt w:val="decimal"/><w:lvlText w:val="%1."/></w:lvl></w:abstractNum><w:num w:numId="1"><w:abstractNumId w:val="0"/></w:num></w:numbering>`,
  );
  parts.set('word/numbering.xml', numbering);
  const result = await readPackage(await render(await writePackage(parts), { note: 'Example note' }));
  for (const kind of ['footnotes', 'endnotes', 'comments'])
    assert(result.get(`word/${kind}.xml`)!.toString().includes('Example note'));
  assert.deepEqual(result.get('word/numbering.xml'), numbering);
});
