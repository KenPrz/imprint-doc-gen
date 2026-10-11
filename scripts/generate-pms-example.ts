import { readFile, writeFile } from 'node:fs/promises';
import { buildDocx } from './corpus/docx.ts';
import { buildXlsx } from './corpus/xlsx.ts';
import { NS, esc } from './corpus/common.ts';
import { writePackage } from '../src/package.ts';
import sharp from 'sharp';

const fixture = JSON.parse(await readFile(new URL('../corpus/docx/level-001/input.json', import.meta.url), 'utf8'));
const output = new URL('../examples/pms/', import.meta.url);
const docx = buildDocx(11, fixture, false);
const xml = docx.get('word/document.xml')!.toString();
const drawing = xml.match(/<w:r><w:drawing>[\s\S]*?<\/w:drawing><\/w:r>/)![0].replace(/\{\{image:[^}]+\}\}/g, '{{image:logo}}');
const p = (value: string, bold = false) => `<w:p><w:r>${bold ? '<w:rPr><w:b/><w:sz w:val="28"/></w:rPr>' : ''}<w:t xml:space="preserve">${esc(value)}</w:t></w:r></w:p>`;
const row = (values: string[]) => `<w:tr>${values.map(value => `<w:tc><w:tcPr><w:tcW w:w="2300" w:type="dxa"/></w:tcPr>${p(value)}</w:tc>`).join('')}</w:tr>`;
const body = `<w:p>${drawing}</w:p>` + p('{{hotel}}', true) + p('Guest invoice {{invoiceNumber}}', true) + p('Guest: {{guest}} | Room: {{room}}') + p('Stay: {{arrival}} to {{departure}}') +
  '<w:tbl><w:tblPr><w:tblW w:w="0" w:type="auto"/><w:tblBorders><w:bottom w:val="single" w:sz="4"/><w:insideH w:val="single" w:sz="4"/></w:tblBorders></w:tblPr><w:tblGrid>' + '<w:gridCol w:w="2300"/>'.repeat(4) + '</w:tblGrid>' +
  row(['Description', 'Quantity', 'Unit price', 'Amount']) + row(['{{#each items}}{{description}}', '{{quantity}}', '{{price}}', '{{amount}}{{/each}}']) + '</w:tbl>' + p('Total: {{currency}} {{total}}', true) + p('Thank you for staying with us.');
docx.set('word/document.xml', Buffer.from(xml.replace(/<w:body>[\s\S]*?<w:sectPr/, `<w:body>${body}<w:sectPr`)));
await writeFile(new URL('invoice.docx', output), await writePackage(docx));
const xlsx = buildXlsx(21, fixture, false);
const cell = (ref: string, value: string) => `<c r="${ref}" t="inlineStr"><is><t>${esc(value)}</t></is></c>`;
const sheetRow = (n: number, values: string[]) => `<row r="${n}">${values.map((value, i) => cell(`${String.fromCharCode(65 + i)}${n}`, value)).join('')}</row>`;
let rows = sheetRow(1, ['{{hotel}}']) + sheetRow(2, ['Invoice', '{{invoiceNumber}}']) + sheetRow(3, ['Guest', '{{guest}}']) + sheetRow(4, ['Room', '{{room}}']) + sheetRow(5, ['Arrival', '{{arrival}}']) + sheetRow(6, ['Departure', '{{departure}}']) + sheetRow(7, ['Currency', '{{currency}}']);
rows += sheetRow(10, ['Description', 'Quantity', 'Unit price', 'Amount']);
rows += `<row r="11">${cell('A11', '{{#each items}}{{description}}')}${cell('B11', '{{quantity}}')}${cell('C11', '{{price}}')}<c r="D11"><f>B11*C11</f><v>0</v></c>${cell('E11', '{{/each}}')}</row>`;
rows += `<row r="12">${cell('A12', 'TOTAL')}<c r="D12"><f>SUM(D11:D11)</f><v>0</v></c></row>`;
const sheet = xlsx.get('xl/worksheets/sheet1.xml')!.toString();
xlsx.set('xl/worksheets/sheet1.xml', Buffer.from(`<worksheet xmlns="${NS.s}" xmlns:r="${NS.r}"><dimension ref="A1:H12"/><cols><col min="1" max="1" width="32" customWidth="1"/><col min="2" max="4" width="20" customWidth="1"/></cols><sheetData>${rows}</sheetData>${sheet.match(/<drawing[^>]*\/>/)![0]}</worksheet>`));
for (const [name, bytes] of xlsx) if (name.startsWith('xl/drawings/') && name.endsWith('.xml')) {
  xlsx.set(name, Buffer.from(bytes.toString().replace(/\{\{image:[^}]+\}\}/g, '{{image:logo}}').replace(/<xdr:row>\d+<\/xdr:row>/g, '<xdr:row>0</xdr:row>').replace(/<xdr:col>\d+<\/xdr:col>/g, '<xdr:col>5</xdr:col>')));
}
// Remove unused synthetic shared strings; the example uses inline strings throughout.
xlsx.delete('xl/sharedStrings.xml');
for (const name of ['xl/_rels/workbook.xml.rels', '[Content_Types].xml']) xlsx.set(name, Buffer.from(xlsx.get(name)!.toString().replace(/<Relationship\b[^>]*Type="[^"]*\/sharedStrings"[^>]*\/>/g, '').replace(/<Override\b[^>]*PartName="\/xl\/sharedStrings.xml"[^>]*\/>/g, '')));
await writeFile(new URL('invoice.xlsx', output), await writePackage(xlsx));
await writeFile(new URL('logo.png', output), await sharp({ create: { width: 240, height: 80, channels: 3, background: '#224c65' } }).png().toBuffer());
