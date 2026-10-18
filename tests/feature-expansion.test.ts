import test from 'node:test';
import assert from 'node:assert/strict';
import { readFile, mkdir, writeFile } from 'node:fs/promises';
import { compile, compileText, encodeTextTemplate, render, updateCharts } from '../src/index.ts';
import { readPackage, writePackage } from '../src/package.ts';
import { ns, parseXml, descendants } from '../src/xml.ts';
import { normalizeImage } from '../src/render-package.ts';
import { parseFormula, evaluateFormula } from '../src/xlsx-formula.ts';
import { renderXmlTemplate } from '../src/xml-template.ts';
import { formatNs as f } from '../src/format-namespaces.ts';
import { workerMetricsServer } from '../src/worker-metrics.ts';
import { RenderService } from '../src/render-service.ts';
import { createApp } from '../src/app.ts';
const p = (text: string) => `<w:p><w:r><w:t>${text}</w:t></w:r></w:p>`;
async function docx(body: string) {
  const parts = await readPackage(await readFile('examples/invoice/invoice.docx'));
  parts.set(
    'word/document.xml',
    Buffer.from(`<w:document xmlns:w="${ns.word}"><w:body>${body}</w:body></w:document>`),
  );
  return parts;
}
test('CSV: quoted records, Unicode, repeated rows, formula injection and worker/API path', async () => {
  const source = 'Name,Note\n{d.rows[i].name},{d.rows[i].note}\n{d.rows[i+1]},';
  const data = {
    rows: [
      { name: 'Zoë, Jr', note: 'a"b\nc' },
      { name: '=SUM(A1)', note: '@evil' },
    ],
  };
  const bytes = await render(encodeTextTemplate('csv', source), data);
  assert.equal(bytes.toString(), 'Name,Note\r\n"Zoë, Jr","a""b\nc"\r\n\'=SUM(A1),\'@evil');
  assert.equal((await render(encodeTextTemplate('csv', source), { rows: [] })).toString(), 'Name,Note');
  assert.throws(() => compileText('csv', '"unclosed'), { code: 'TEMPLATE_SYNTAX' });
  const app = await createApp({ apiKey: 'feature-expansion-test-api-key-32' });
  try {
    const response = await app.inject({
      method: 'POST',
      url: '/v1/render',
      headers: { authorization: 'Bearer feature-expansion-test-api-key-32' },
      payload: { template: encodeTextTemplate('csv', source).toString('base64'), data },
    });
    assert.equal(response.statusCode, 200, response.body);
    assert.equal(response.body, bytes.toString());
  } finally {
    await app.close();
  }
});
test('TXT preserves literal text and rejects malformed repeats and oversized data', async () => {
  await assert.rejects(
    compileText('txt', '{d.value}{d.value}').render({ value: 'x'.repeat(5 * 1024 * 1024) }),
    { code: 'LIMIT_EXCEEDED' },
  );
  assert.equal(
    (
      await compileText('txt', 'Hello {d.name}\n{d.rows[i].name}\n{d.rows[i+1]}').render({
        name: '世界',
        rows: [{ name: 'A' }, { name: 'B' }],
      })
    ).toString(),
    'Hello 世界\nA\nB',
  );
  await assert.rejects(compileText('txt', '{d.rows[i].name}').render({ rows: [] }), {
    code: 'TEMPLATE_STRUCTURE',
  });
});
test('New formula families: independent expected scalars, ranges and error propagation', () => {
  const cases: [string, unknown][] = [
    ['LEFT("abcd",2)', 'ab'],
    ['RIGHT("abcd",0)', ''],
    ['RIGHT("abcd",0.5)', ''],
    ['MID("abcd",2,2)', 'bc'],
    ['TRIM("  a   b  ")', 'a b'],
    ['INT(-2.3)', -3],
    ['MOD(-3,2)', 1],
    ['MOD(3,-2)', -1],
    ['SQRT(9)', 3],
    ['POWER(2,3)', 8],
    ['MOD(1,0)', { error: '#DIV/0!' }],
    ['MID("a",0,1)', { error: '#VALUE!' }],
    ['SQRT(-1)', { error: '#NUM!' }],
    ['COUNTA(A1:A4)', 3],
    ['COUNTBLANK(A1:A4)', 2],
  ];
  const values = [null, '', 4, { error: '#N/A' }] as const;
  for (const [source, expected] of cases)
    assert.deepEqual(
      evaluateFormula(parseFormula(source), 'Sheet1', (_s, r) => values[Number(r.slice(1)) - 1]!, {
        operations: 0,
      }),
      expected,
      source,
    );
});
test('SVG passive geometry rasterizes; active and external SVG fail closed', async () => {
  const svg =
    '<svg xmlns="http://www.w3.org/2000/svg" width="40" height="20"><rect width="40" height="20" fill="#ff0000"/></svg>';
  const image = await normalizeImage(Buffer.from(svg));
  assert.equal(image.extension, 'png');
  assert.equal(image.bytes.subarray(1, 4).toString(), 'PNG');
  for (const body of [
    '<script>alert(1)</script>',
    '<image href="file:///etc/passwd"/>',
    '<rect style="fill:url(http://example.com/a)"/>',
  ])
    await assert.rejects(
      normalizeImage(Buffer.from(svg.replace('<rect width="40" height="20" fill="#ff0000"/>', body))),
    );
});
test('DOCX repeated footnotes, endnotes, classic comments and restarted numbering have distinct IDs', async () => {
  const parts = await docx(
    p('{{#each rows}}') +
      `<w:p><w:commentRangeStart w:id="1"/><w:r><w:t>Comment anchor</w:t></w:r><w:commentRangeEnd w:id="1"/><w:r><w:commentReference w:id="1"/></w:r></w:p><w:p><w:pPr><w:numPr><w:ilvl w:val="0"/><w:numId w:val="1"/></w:numPr></w:pPr><w:r><w:t>{{name}}</w:t></w:r><w:r><w:footnoteReference w:id="1"/></w:r><w:r><w:t xml:space="preserve"> end</w:t></w:r><w:r><w:endnoteReference w:id="1"/></w:r></w:p>` +
      p('{{/each}}'),
  );
  for (const kind of ['footnote', 'endnote', 'comment']) {
    parts.set(
      `word/${kind}s.xml`,
      Buffer.from(
        `<w:${kind}s xmlns:w="${ns.word}">${kind === 'comment' ? '' : `<w:${kind} w:type="separator" w:id="-1"><w:p><w:r><w:separator/></w:r></w:p></w:${kind}><w:${kind} w:type="continuationSeparator" w:id="0"><w:p><w:r><w:continuationSeparator/></w:r></w:p></w:${kind}>`}<w:${kind} w:id="1" ${kind === 'comment' ? 'w:author="Test"' : ''}>${kind === 'comment' ? p('{{note}}') : `<w:p><w:r><w:${kind}Ref/></w:r><w:r><w:t>{{note}}</w:t></w:r></w:p>`}</w:${kind}></w:${kind}s>`,
      ),
    );
    parts.set(
      'word/_rels/document.xml.rels',
      Buffer.from(
        parts
          .get('word/_rels/document.xml.rels')!
          .toString()
          .replace(
            '</Relationships>',
            `<Relationship Id="${kind}" Type="${ns.relationship}/${kind}s" Target="${kind}s.xml"/></Relationships>`,
          ),
      ),
    );
  }
  for (const kind of ['footnotes', 'endnotes', 'comments', 'numbering'])
    parts.set(
      '[Content_Types].xml',
      Buffer.from(
        parts
          .get('[Content_Types].xml')!
          .toString()
          .replace(
            '</Types>',
            `<Override PartName="/word/${kind}.xml" ContentType="application/vnd.openxmlformats-officedocument.wordprocessingml.${kind}+xml"/></Types>`,
          ),
      ),
    );
  parts.set(
    'word/_rels/document.xml.rels',
    Buffer.from(
      parts
        .get('word/_rels/document.xml.rels')!
        .toString()
        .replace(
          '</Relationships>',
          `<Relationship Id="numbering" Type="${ns.relationship}/numbering" Target="numbering.xml"/></Relationships>`,
        ),
    ),
  );
  parts.set(
    'word/numbering.xml',
    Buffer.from(
      `<w:numbering xmlns:w="${ns.word}"><w:abstractNum w:abstractNumId="0"><w:lvl w:ilvl="0"><w:start w:val="1"/><w:numFmt w:val="decimal"/><w:lvlText w:val="%1."/></w:lvl></w:abstractNum><w:num w:numId="1"><w:abstractNumId w:val="0"/></w:num></w:numbering>`,
    ),
  );
  const data = {
    rows: [
      { name: 'One', note: 'First note' },
      { name: 'Two', note: 'Second note' },
    ],
  };
  await assert.rejects(render(await writePackage(parts), data), { code: 'UNSUPPORTED_TEMPLATE' });
  for (const mode of ['notes', 'comments']) {
    const variant = new Map(parts);
    let source = variant.get('word/document.xml')!.toString();
    if (mode === 'notes') source = source.replace(/<w:comment(?:Reference|RangeStart|RangeEnd)[^>]*\/>/g, '');
    else source = source.replace(/<w:(?:footnote|endnote)Reference[^>]*\/>/g, '');
    variant.set('word/document.xml', Buffer.from(source));
    // Unused definitions use root fields, so they remain valid preserved content.
    for (const kind of mode === 'notes' ? ['comment'] : ['footnote', 'endnote'])
      variant.set(
        `word/${kind}s.xml`,
        Buffer.from(variant.get(`word/${kind}s.xml`)!.toString().replace('{{note}}', 'Unreferenced')),
      );
    const result = await readPackage(
      await render(await writePackage(variant), data, { restartNumbering: true }),
    );
    const empty = await readPackage(await render(await writePackage(variant), { rows: [] }));
    for (const kind of mode === 'notes' ? ['footnote', 'endnote'] : ['comment'])
      assert.equal(
        descendants(parseXml(empty.get(`word/${kind}s.xml`)!.toString(), kind), ns.word, kind).filter(
          (n) => Number(n.getAttributeNS(ns.word, 'id')) > 0,
        ).length,
        0,
      );
    const main = parseXml(result.get('word/document.xml')!.toString(), 'main');
    for (const kind of mode === 'notes' ? ['footnote', 'endnote'] : ['comment']) {
      const refs = descendants(main, ns.word, kind + 'Reference').map((n) => n.getAttributeNS(ns.word, 'id'));
      const defs = descendants(
        parseXml(result.get(`word/${kind}s.xml`)!.toString(), kind),
        ns.word,
        kind,
      ).filter((n) => Number(n.getAttributeNS(ns.word, 'id')) > 0);
      assert.equal(new Set(refs).size, 2);
      assert.deepEqual(
        defs.map((n) => n.textContent),
        ['First note', 'Second note'],
      );
      assert.deepEqual(
        defs.map((n) => n.getAttributeNS(ns.word, 'id')),
        refs,
      );
    }
    assert.equal(
      new Set(descendants(main, ns.word, 'numId').map((n) => n.getAttributeNS(ns.word, 'val'))).size,
      2,
    );
    await mkdir('artifacts/feature-expansion', { recursive: true });
    await writeFile(`artifacts/feature-expansion/repeated-${mode}.docx`, await writePackage(result));
  }
});
test('Inline DOCX images create relationships and rasterized media', async () => {
  const parts = await docx(p('Before') + p('{{inline:logo:40:20}}') + p('After'));
  const bytes = await render(
    await writePackage(parts),
    { logo: 'logo' },
    {
      attachments: {
        logo: Buffer.from(
          '<svg xmlns="http://www.w3.org/2000/svg" width="40" height="20"><rect width="40" height="20" fill="#00ff00"/></svg>',
        ),
      },
    },
  );
  const result = await readPackage(bytes);
  assert.match(result.get('word/document.xml')!.toString(), /wp:inline/);
  assert.match(result.get('word/_rels/document.xml.rels')!.toString(), /imprintImage/);
  await mkdir('artifacts/feature-expansion', { recursive: true });
  await writeFile('artifacts/feature-expansion/inline.docx', bytes);
});
test('ODS repeats rewrite local and absolute references and clear stale caches', async () => {
  const cell = (text: string, formula = '') =>
    `<table:table-cell ${formula ? `table:formula="${formula}" office:value-type="float" office:value="999"` : ''}><text:p>${text}</text:p></table:table-cell>`;
  const source = `<office:document-content xmlns:office="${f.office}" xmlns:table="${f.table}" xmlns:text="${f.text}"><office:body><office:spreadsheet><table:table table:name="Sheet1"><table:table-row>${cell('{d.rows[i].n}')}${cell('999', 'of:=[.A1]+[.A$1]')}</table:table-row><table:table-row>${cell('{d.rows[i+1].n}')}</table:table-row><table:table-row>${cell('999', 'of:=SUM([.A1:.A1])')}</table:table-row></table:table></office:spreadsheet></office:body></office:document-content>`;
  const result = parseXml(
    (
      await renderXmlTemplate(source, 'content.xml', { rows: [{ n: 2 }, { n: 3 }] }, { odsFormulas: true })
    ).toString(),
    'result',
  );
  const formulas = descendants(result, f.table, 'table-cell').filter((c) =>
    c.hasAttributeNS(f.table, 'formula'),
  );
  assert.deepEqual(
    formulas.map((c) => c.getAttributeNS(f.table, 'formula')),
    ['of:=[.A1]+[.A$1]', 'of:=[.A2]+[.A$1]', 'of:=SUM([.A1:.A2])'],
  );
  assert(formulas.every((c) => !c.hasAttributeNS(f.office, 'value')));
  await assert.rejects(
    renderXmlTemplate(
      source.replace('[.A1]+[.A$1]', '[Sheet2.A1]'),
      'content.xml',
      { rows: [{ n: 1 }] },
      { odsFormulas: true },
    ),
    { code: 'UNSUPPORTED_FORMULA' },
  );
});
test('Dynamic chart data updates series, categories and values without dangling workbook links', async () => {
  const parts = await docx(p('Chart')),
    c = 'http://schemas.openxmlformats.org/drawingml/2006/chart';
  const part = 'word/charts/chart1.xml';
  parts.set(
    part,
    Buffer.from(
      `<c:chartSpace xmlns:c="${c}"><c:chart><c:plotArea><c:barChart><c:ser><c:idx val="0"/><c:order val="0"/><c:tx><c:v>Old</c:v></c:tx><c:cat><c:strLit/></c:cat><c:val><c:numLit/></c:val></c:ser></c:barChart></c:plotArea></c:chart></c:chartSpace>`,
    ),
  );
  const updates = {
    [part]: {
      categories: ['A', 'B'],
      series: [
        { name: 'Revenue', values: [2, 4] },
        { name: 'Costs', values: [1, 3] },
      ],
    },
  };
  const output = await readPackage(await updateCharts(await writePackage(parts), updates));
  const chart = parseXml(output.get(part)!.toString(), part);
  assert.equal(descendants(chart, c, 'ser').length, 2);
  assert.deepEqual(
    descendants(chart, c, 'v').map((n) => n.textContent),
    ['Revenue', 'A', 'B', '2', '4', 'Costs', 'A', 'B', '1', '3'],
  );
  await assert.rejects(
    updateCharts(await writePackage(parts), {
      [part]: { categories: ['A'], series: [{ name: 'x', values: [NaN] }] },
    }),
    { code: 'INVALID_DATA' },
  );
  parts.set(
    part,
    Buffer.from(parts.get(part)!.toString().replace('<c:chart>', '<c:externalData/><c:chart>')),
  );
  await assert.rejects(updateCharts(await writePackage(parts), updates), { code: 'UNSUPPORTED_TEMPLATE' });
});
test('Worker metrics authenticate, expose independent counters and report store failure', async () => {
  const renderer = new RenderService();
  renderer.metrics.render(true, 0.5, 123);
  let fail = false;
  const token = 'metrics-test-token-32-characters';
  const server = workerMetricsServer(
    renderer,
    () => {
      if (fail) throw Error();
      return 3;
    },
    token,
  );
  await new Promise<void>((resolve) => server.listen(0, '127.0.0.1', resolve));
  const port = (server.address() as { port: number }).port;
  try {
    const url = `http://127.0.0.1:${port}/metrics`;
    assert.equal((await fetch(url)).status, 401);
    const response = await fetch(url, { headers: { authorization: 'Bearer ' + token } });
    assert.match(await response.text(), /imprint_output_bytes_total 123/);
    fail = true;
    assert.equal((await fetch(url, { headers: { authorization: 'Bearer ' + token } })).status, 503);
  } finally {
    await new Promise<void>((resolve) => server.close(() => resolve()));
  }
});
test('Text envelopes reject invalid UTF-8, unsupported options, and inspect fields', async () => {
  const { inspectTemplate } = await import('../src/inspect.ts');
  assert.equal((await inspectTemplate(encodeTextTemplate('txt', '{d.name}'))).fields[0]?.path, 'name');
  await assert.rejects(compile(Buffer.concat([Buffer.from('imprint-text-v1\n'), Buffer.from([255])])), {
    code: 'TEMPLATE_SYNTAX',
  });
  await assert.rejects(compileText('txt', 'a').render({}, { restartNumbering: true }), {
    code: 'UNSUPPORTED_FORMAT',
  });
  assert.equal(
    (
      await compileText('txt', '{d.amount:number(2)}').render({ amount: 1234.5 }, { locale: 'de-DE' })
    ).toString(),
    '1.234,50',
  );
});
test('Repeated note failures and missing inline images fail without producing invalid references', async () => {
  const bad = await docx(
    p('{{#each rows}}') + '<w:p><w:r><w:footnoteReference w:id="99"/></w:r></w:p>' + p('{{/each}}'),
  );
  await assert.rejects(render(await writePackage(bad), { rows: [{}] }), { code: 'INVALID_PACKAGE' });
  const inline = await docx(p('{{inline:missing:20:20}}'));
  await assert.rejects(render(await writePackage(inline), {}), { code: 'MISSING_ATTACHMENT' });
});
test('ODS formulas preserve references beyond explicit rows and ignore bracket text literals', async () => {
  const source = `<office:document-content xmlns:office="${f.office}" xmlns:table="${f.table}" xmlns:text="${f.text}"><table:table><table:table-row><table:table-cell table:formula='of:=[.A100]&amp;"[not a reference]"'/></table:table-row></table:table></office:document-content>`;
  const out = await renderXmlTemplate(source, 'content.xml', {}, { odsFormulas: true });
  assert.match(out.toString(), /A100/);
  assert.match(out.toString(), /not a reference/);
});
test('Comment ranges enclosing note references fail before rendering', async () => {
  const parts = await docx(
    p('{{#each rows}}') +
      '<w:p><w:commentRangeStart w:id="1"/><w:r><w:footnoteReference w:id="1"/></w:r><w:commentRangeEnd w:id="1"/></w:p>' +
      p('{{/each}}'),
  );
  await assert.rejects(compile(await writePackage(parts)), { code: 'UNSUPPORTED_TEMPLATE' });
});
test('Independently generated literal chart workbook remains an editable package', async () => {
  const input = await readFile('tests/fixtures/literal-chart.xlsx');
  const bytes = await render(
    input,
    {},
    {
      charts: {
        'xl/charts/chart1.xml': {
          categories: ['January', 'February'],
          series: [
            { name: 'Revenue', values: [120, 150] },
            { name: 'Costs', values: [80, 90] },
          ],
        },
      },
    },
  );
  const parts = await readPackage(bytes);
  assert.match(parts.get('xl/charts/chart1.xml')!.toString(), /Revenue/);
  assert(parts.has('xl/worksheets/sheet1.xml'));
  await mkdir('artifacts/feature-expansion', { recursive: true });
  await writeFile('artifacts/feature-expansion/chart.xlsx', bytes);
});
