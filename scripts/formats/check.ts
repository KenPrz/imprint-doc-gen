import { mkdir, writeFile, mkdtemp, readFile, rm } from 'node:fs/promises';
import { tmpdir } from 'node:os';
import { join } from 'node:path';
import { execFileSync } from 'node:child_process';
import assert from 'node:assert/strict';
import { fixture, sample, blue, type ExtraFormat } from './fixtures.ts';
import { render } from '../../src/index.ts';
import { resolvePart, relsPath, readPackage } from '../../src/package.ts';
import { validateRelationships } from '../../src/ooxml.ts';
import { parseXml, descendants, ns } from '../../src/xml.ts';
import { formatNs as f } from '../../src/format-namespaces.ts';
import { convertToPdf } from '../../src/converter.ts';
const directory = 'artifacts/formats';
await mkdir(directory, { recursive: true });
const cases: unknown[] = [];
const text = async (bytes: Uint8Array, format: ExtraFormat) => {
  const parts = await readPackage(bytes);
  if (format === 'pptx') validateRelationships(parts);
  let paths = ['content.xml'];
  if (format === 'pptx') {
    const main = 'ppt/presentation.xml',
      rels = descendants(
        parseXml(parts.get(relsPath(main))!.toString(), main),
        ns.packageRelationship,
        'Relationship',
      );
    paths = descendants(parseXml(parts.get(main)!.toString(), main), f.presentation, 'sldId').map((n) =>
      resolvePart(
        main,
        rels
          .find((r) => r.getAttribute('Id') === n.getAttributeNS(ns.relationship, 'id'))!
          .getAttribute('Target')!,
      ),
    );
  }
  return paths
    .flatMap((path) =>
      descendants(
        parseXml(parts.get(path)!.toString(), path),
        format === 'pptx' ? ns.drawing : f.text,
        'p',
      ).map((p) => p.textContent ?? ''),
    )
    .join('\n');
};
for (const format of ['pptx', 'odt', 'ods', 'odp'] as ExtraFormat[])
  for (let level = 1; level <= 100; level++) {
    const template = await fixture(format, level),
      expected = await fixture(format, level, true),
      output = await render(template, sample(level), { attachments: { logo: blue } });
    assert.equal(
      await text(output, format),
      await text(expected, format),
      `${format} ${level}: independent text oracle`,
    );
    const xml = [...(await readPackage(output))]
      .filter(([p]) => p.endsWith('.xml'))
      .map(([, b]) => b.toString())
      .join('');
    assert(!xml.includes('{d.'));
    let pdf = false;
    if ((process.env.GOTENBERG_URL || process.env.IMPRINT_URL) && [1, 40, 100].includes(level)) {
      const temporary = await mkdtemp(join(tmpdir(), 'imprint-formats-'));
      try {
        const texts: string[] = [];
        for (const [name, bytes] of [
          ['actual', output],
          ['expected', expected],
        ] as const) {
          const converted = process.env.GOTENBERG_URL
            ? await convertToPdf(bytes, process.env.GOTENBERG_URL, `input.${format}`)
            : await apiPdf(bytes);
          await writeFile(join(temporary, `${name}.pdf`), converted);
          execFileSync('pdftotext', [
            '-layout',
            join(temporary, `${name}.pdf`),
            join(temporary, `${name}.txt`),
          ]);
          texts.push((await readFile(join(temporary, `${name}.txt`), 'utf8')).replace(/\s+/g, ' ').trim());
          await writeFile(join(directory, `${format}-${level}-${name}.pdf`), converted);
        }
        assert.equal(texts[0], texts[1], `${format} ${level}: converted text oracle`);
        assert(texts[0]!.includes(sample(level).title));
        pdf = true;
      } finally {
        await rm(temporary, { recursive: true, force: true });
      }
    }
    cases.push({ format, level, passed: true, pdf });
    if (level % 25 === 0) console.log(`${format}: ${level}/100`);
  }
await writeFile(
  join(directory, 'report.json'),
  JSON.stringify({ generatedAt: new Date().toISOString(), cases }, null, 2) + '\n',
);
console.log('400 format cases passed');

async function apiPdf(bytes: Uint8Array) {
  const response = await fetch(new URL('/v1/render', process.env.IMPRINT_URL), {
    method: 'POST',
    headers: { authorization: `Bearer ${process.env.IMPRINT_API_KEY}`, 'content-type': 'application/json' },
    body: JSON.stringify({ template: Buffer.from(bytes).toString('base64'), data: {}, output: 'pdf' }),
    signal: AbortSignal.timeout(100000),
  });
  if (!response.ok) throw new Error(`PDF API returned ${response.status}: ${await response.text()}`);
  return Buffer.from(await response.arrayBuffer());
}
