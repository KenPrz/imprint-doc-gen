import test from 'node:test';
import assert from 'node:assert/strict';
import { fixture, sample, blue, type ExtraFormat } from '../scripts/formats/fixtures.ts';
import { render, compile } from '../src/index.ts';
import { readPackage } from '../src/package.ts';
import { validateRelationships } from '../src/ooxml.ts';
import { ns, parseXml, descendants } from '../src/xml.ts';
import { formatNs as f } from '../src/format-namespaces.ts';
for (const format of ['pptx', 'odt', 'ods', 'odp'] as ExtraFormat[])
  for (const level of [1, 20, 40, 60, 80, 100])
    test(`${format} level ${level}: independent content and package checks`, async () => {
      const input = await fixture(format, level),
        result = await readPackage(await render(input, sample(level), { attachments: { logo: blue } }));
      assert.equal((await compile(input)).format, format);
      if (format === 'pptx') validateRelationships(result);
      const xml = [...result]
        .filter(([name]) => name.endsWith('.xml'))
        .map(([, b]) => b.toString())
        .join('');
      assert(xml.includes(`Example document ${level}`));
      assert(!xml.includes('{d.'));
      if (level >= 11) assert([...result.values()].some((bytes) => bytes.equals(blue)));
      if (format === 'pptx')
        assert.equal(
          descendants(parseXml(result.get('ppt/presentation.xml')!.toString(), 'p'), f.presentation, 'sldId')
            .length,
          level >= 61 ? sample(level).items.length : 1,
        );
      else assert.equal(result.get('mimetype')!.toString().startsWith('application/vnd.oasis'), true);
    });
test('new formats preserve multiline replacements and reject malformed ODF packages', async () => {
  for (const format of ['pptx', 'odt', 'ods', 'odp'] as ExtraFormat[]) {
    const bytes = await render(await fixture(format, 1), { ...sample(1), customer: 'Line one\nLine two' }),
      parts = await readPackage(bytes);
    const xml = [...parts]
      .filter(([path]) => path.endsWith('.xml'))
      .map(([, b]) => b.toString())
      .join('');
    assert(xml.includes(format === 'pptx' ? 'a:br' : 'text:line-break'));
  }
  const { writePackage } = await import('../src/package.ts');
  const parts = await readPackage(await fixture('odt', 1));
  parts.set('mimetype', Buffer.from('application/vnd.oasis.opendocument.spreadsheet'));
  await assert.rejects(compile(await writePackage(parts)), /media type/);
});
