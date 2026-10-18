// Additional native package for independent LibreOffice recalculation.
import { mkdir, writeFile, readFile } from 'node:fs/promises';
import { fixture } from './formats/fixtures.ts';
import { readPackage, writePackage } from '../src/package.ts';
import { render } from '../src/index.ts';
import { formatNs as f } from '../src/format-namespaces.ts';
const parts = await readPackage(await fixture('ods', 1));
const cell = (text: string, formula = '') =>
  `<table:table-cell ${formula ? `table:formula="${formula}" office:value-type="float" office:value="999"` : ''}><text:p>${text}</text:p></table:table-cell>`;
parts.set(
  'content.xml',
  Buffer.from(
    `<office:document-content xmlns:office="${f.office}" xmlns:table="${f.table}" xmlns:text="${f.text}" xmlns:of="urn:oasis:names:tc:opendocument:xmlns:of:1.2" office:version="1.3"><office:body><office:spreadsheet><table:table table:name="Report"><table:table-row>${cell('{d.rows[i].n}')}${cell('999', 'of:=[.A1]*2')}</table:table-row><table:table-row>${cell('{d.rows[i+1].n}')}</table:table-row><table:table-row>${cell('999', 'of:=SUM([.A1:.A1])')}</table:table-row></table:table></office:spreadsheet></office:body></office:document-content>`,
  ),
);
await mkdir('artifacts/feature-expansion', { recursive: true });
await writeFile(
  'artifacts/feature-expansion/formulas.ods',
  await render(await writePackage(parts), { rows: [{ n: 2 }, { n: 3 }] }),
);
console.log('Wrote formulas.ods: expected A1=2, B1=4, A2=3, B2=6, A3=5 after recalculation.');
