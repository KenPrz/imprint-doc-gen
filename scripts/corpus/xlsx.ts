import { amount, type Data } from './recipes.ts';
import { NS, esc, put, png, contentTypes, rels, perturb, type Parts } from './common.ts';

export function buildXlsx(level: number, data: Data, rendered: boolean): Parts {
  const parts: Parts = new Map();
  const [basic, shared, pictures, layout, formulas, repeats, multi, ranges, disorder, stress] = Array.from({ length: 10 }, (_, i) => amount(level, i + 1));
  const sharedStrings: string[] = [];
  let sharedReferences = 0;
  const binding = (key: string, value: string) => rendered ? value : `{{${key}}}`;
  const rich = (value: string) => {
    const chars = Array.from(value), count = Math.min(chars.length, Math.max(1, shared));
    const runs = [];
    for (let i = 0; i < count; i++) runs.push(`<r><rPr><b/><sz val="11"/><rFont val="Calibri"/></rPr><t xml:space="preserve">${esc(chars.slice(Math.floor(i * chars.length / count), Math.floor((i + 1) * chars.length / count)).join(''))}</t></r>`);
    return runs.join('');
  };
  const cell = (ref: string, value: string | number | boolean, style = 0, formula?: string, forcePlain = false): string => {
    const attrs = `r="${ref}" s="${style}"`;
    if (formula) return `<c ${attrs}><f>${esc(formula)}</f><v>${value}</v></c>`;
    if (typeof value === 'number') return `<c ${attrs}><v>${value}</v></c>`;
    if (typeof value === 'boolean') return `<c ${attrs} t="b"><v>${value ? 1 : 0}</v></c>`;
    if (shared && !forcePlain) {
      const content = `<si>${rich(value)}</si>`;
      let index = sharedStrings.indexOf(content);
      if (index < 0) { index = sharedStrings.length; sharedStrings.push(content); }
      sharedReferences++;
      return `<c ${attrs} t="s"><v>${index}</v></c>`;
    }
    return `<c ${attrs} t="inlineStr"><is><t xml:space="preserve">${esc(value)}</t></is></c>`;
  };
  const row = (n: number, cells: string, height?: number) => `<row r="${n}"${height ? ` ht="${height}" customHeight="1"` : ''}>${cells}</row>`;
  const endRow = repeats ? 39 + (rendered ? data.items.length : 1) : 30;
  const sheetCount = multi ? 2 + multi : formulas ? 2 : 1;
  const types: Record<string, string> = { 'xl/workbook.xml': 'spreadsheetml.sheet.main', 'xl/styles.xml': 'spreadsheetml.styles' };
  const bookRels = [{ id: 'rId50', type: 'styles', target: 'styles.xml' }];
  const sheetNames = Array.from({ length: sheetCount }, (_, i) => i === 0 ? 'Report' : i === 1 ? 'Lookup Data' : `Region ${i - 1} 山田`);
  for (let sheetIndex = 0; sheetIndex < sheetCount; sheetIndex++) {
    const name = `xl/worksheets/sheet${sheetIndex + 1}.xml`;
    types[name] = 'spreadsheetml.worksheet';
    bookRels.push({ id: `rId${sheetIndex + 1}`, type: 'worksheet', target: `worksheets/sheet${sheetIndex + 1}.xml` });
    if (sheetIndex > 0) {
      put(parts, name, `<worksheet xmlns="${NS.s}"><dimension ref="A1:B2"/><sheetData>${row(1, cell('A1', sheetNames[sheetIndex]!) + cell('B1', 7))}${row(2, cell('A2', binding('customer', data.customer)) + cell('B2', binding('title', data.title)))}</sheetData></worksheet>`);
      continue;
    }
    let rows = row(1, cell('A1', `Imprint Doc Gen XLSX level ${level}`) + cell('B1', binding('title', data.title))) + row(2, cell('A2', binding('customer', data.customer)) + cell('B2', rendered ? data.count : '{{count}}'));
    rows += row(3, cell('A3', binding('spaced', data.spaced)) + cell('B3', rendered ? data.enabled : '{{enabled}}'));
    rows += row(4, cell('A4', binding('literal', data.literal)) + cell('B4', 12.5, layout ? 2 : 0));
    for (let i = 0; i < basic; i++) rows += row(5 + i, cell(`A${5 + i}`, binding('customer', data.customer), layout ? 1 : 0));
    for (let i = 0; i < layout; i++) rows += row(15 + i, cell(`A${15 + i}`, binding('title', data.title), 1), 24 + i);
    for (let i = 0; i < formulas; i++) {
      const n = 25 + i;
      rows += row(n, cell(`A${n}`, i + 1) + cell(`B${n}`, 2) + cell(`C${n}`, (i + 1) * 2 + 7, 0, `A${n}*$B${n}+'Lookup Data'!$B$1`) + cell(`D${n}`, i + 8, 0, `$A${n}+'Lookup Data'!B$1`));
    }
    if (repeats) {
      rows += row(39, ['Name', 'Quantity', 'Price', 'Total'].map((v, i) => cell(`${String.fromCharCode(65 + i)}39`, v, 1)).join(''));
      const items = rendered ? data.items : data.items.slice(0, 1);
      items.forEach((item, i) => {
        const n = 40 + i;
        rows += row(n, cell(`A${n}`, rendered ? item.name : '{{#each items}}{{name}}') + cell(`B${n}`, rendered ? item.quantity : '{{quantity}}') + cell(`C${n}`, rendered ? item.price : '{{price}}', 2) + cell(`D${n}`, rendered ? item.quantity * item.price : 0, 2, `B${n}*C${n}`) + (rendered ? '' : cell(`E${n}`, '{{/each}}', 0, undefined, true)));
      });
    }
    for (let i = 0; i < stress * 40; i++) {
      const n = 100 + (rendered && repeats ? data.items.length - 1 : 0) + i;
      rows += row(n, cell(`A${n}`, `Stress ${i + 1}`) + cell(`B${n}`, binding('customer', data.customer)) + cell(`C${n}`, i));
    }
    const actualLastRow = stress ? 100 + (rendered && repeats ? data.items.length - 1 : 0) + stress * 40 - 1 : Math.max(repeats ? endRow : 0, 4 + basic, layout ? 14 + layout : 0, formulas ? 24 + formulas : 0);
    let suffix = '';
    if (ranges) suffix += `<autoFilter ref="A39:D${endRow}"/>`;
    if (layout) suffix += `<mergeCells count="${layout}">${Array.from({ length: layout }, (_, i) => `<mergeCell ref="A${15 + i}:C${15 + i}"/>`).join('')}</mergeCells>`;
    if (ranges >= 2) suffix += `<conditionalFormatting sqref="B40:B${endRow}"><cfRule type="cellIs" priority="1" operator="greaterThan"><formula>5</formula></cfRule></conditionalFormatting>`;
    if (ranges >= 3) suffix += `<dataValidations count="${ranges - 2}">${Array.from({ length: ranges - 2 }, (_, i) => `<dataValidation type="whole" operator="between" allowBlank="1" sqref="${String.fromCharCode(66 + i)}40:${String.fromCharCode(66 + i)}${endRow}"><formula1>0</formula1><formula2>10000</formula2></dataValidation>`).join('')}</dataValidations>`;
    suffix += '<pageMargins left="0.7" right="0.7" top="0.75" bottom="0.75" header="0.3" footer="0.3"/>';
    if (pictures) suffix += '<drawing r:id="rId70"/>';
    if (ranges >= 4) suffix += '<tableParts count="1"><tablePart r:id="rId71"/></tableParts>';
    put(parts, name, `<worksheet xmlns="${NS.s}" xmlns:r="${NS.r}"><dimension ref="A1:J${actualLastRow}"/><sheetViews><sheetView workbookViewId="0">${layout ? '<pane xSplit="1" ySplit="4" topLeftCell="B5" activePane="bottomRight" state="frozen"/>' : ''}</sheetView></sheetViews><sheetFormatPr defaultRowHeight="15"/>${layout ? '<cols><col min="1" max="1" width="34" customWidth="1"/><col min="2" max="10" width="18" customWidth="1"/></cols>' : ''}<sheetData>${rows}</sheetData>${suffix}</worksheet>`);
    const links = [];
    if (pictures) links.push({ id: 'rId70', type: 'drawing', target: '../drawings/drawing1.xml' });
    if (ranges >= 4) links.push({ id: 'rId71', type: 'table', target: '../tables/table1.xml' });
    if (links.length) put(parts, 'xl/worksheets/_rels/sheet1.xml.rels', rels(links));
  }
  if (pictures) {
    const drawingLinks = [];
    let drawings = '';
    for (let i = 0; i < pictures; i++) {
      const color = i % 2 ? 'blue' : 'red', id = `rId${i + 1}`;
      drawingLinks.push({ id, type: 'image', target: `../media/${rendered ? color : 'placeholder'}.png` });
      drawings += `<xdr:oneCellAnchor><xdr:from><xdr:col>${5 + i % 3}</xdr:col><xdr:colOff>0</xdr:colOff><xdr:row>${i * 3}</xdr:row><xdr:rowOff>0</xdr:rowOff></xdr:from><xdr:ext cx="${914400 + i * 45720}" cy="609600"/><xdr:pic><xdr:nvPicPr><xdr:cNvPr id="${i + 1}" name="Picture ${i + 1}" descr="${esc(rendered ? `Attachment ${color}` : `{{image:${color}}}`)}"/><xdr:cNvPicPr/></xdr:nvPicPr><xdr:blipFill><a:blip r:embed="${id}"/><a:stretch><a:fillRect/></a:stretch></xdr:blipFill><xdr:spPr><a:xfrm><a:off x="0" y="0"/><a:ext cx="${914400 + i * 45720}" cy="609600"/></a:xfrm><a:prstGeom prst="rect"><a:avLst/></a:prstGeom></xdr:spPr></xdr:pic><xdr:clientData/></xdr:oneCellAnchor>`;
    }
    put(parts, 'xl/drawings/drawing1.xml', `<xdr:wsDr xmlns:xdr="${NS.xdr}" xmlns:a="${NS.a}" xmlns:r="${NS.r}">${drawings}</xdr:wsDr>`);
    put(parts, 'xl/drawings/_rels/drawing1.xml.rels', rels(drawingLinks));
    types['xl/drawings/drawing1.xml'] = 'drawing';
    parts.set('xl/media/placeholder.png', png([150, 150, 150]));
    if (rendered) { parts.set('xl/media/red.png', png([220, 40, 40])); parts.set('xl/media/blue.png', png([40, 80, 220])); }
  }
  if (ranges >= 4) {
    put(parts, 'xl/tables/table1.xml', `<table xmlns="${NS.s}" id="1" name="Items" displayName="Items" ref="A39:D${endRow}" totalsRowShown="0"><autoFilter ref="A39:D${endRow}"/><tableColumns count="4">${['Name', 'Quantity', 'Price', 'Total'].map((v, i) => `<tableColumn id="${i + 1}" name="${v}"/>`).join('')}</tableColumns><tableStyleInfo name="TableStyleMedium2" showFirstColumn="0" showLastColumn="0" showRowStripes="1" showColumnStripes="0"/></table>`);
    types['xl/tables/table1.xml'] = 'spreadsheetml.table';
  }
  const names = ranges ? `<definedNames>${Array.from({ length: ranges }, (_, i) => `<definedName name="ItemRange${i + 1}">'Report'!$A$40:$D$${endRow}</definedName>`).join('')}</definedNames>` : '';
  put(parts, 'xl/workbook.xml', `<workbook xmlns="${NS.s}" xmlns:r="${NS.r}"><bookViews><workbookView/></bookViews><sheets>${sheetNames.map((name, i) => `<sheet name="${esc(name)}" sheetId="${i + 1}" r:id="rId${i + 1}"/>`).join('')}</sheets>${names}<calcPr calcId="191029" fullCalcOnLoad="1"/></workbook>`);
  put(parts, 'xl/styles.xml', `<styleSheet xmlns="${NS.s}"><fonts count="2"><font><sz val="11"/><name val="Calibri"/></font><font><b/><sz val="11"/><name val="Calibri"/></font></fonts><fills count="2"><fill><patternFill patternType="none"/></fill><fill><patternFill patternType="gray125"/></fill></fills><borders count="1"><border><left/><right/><top/><bottom/><diagonal/></border></borders><cellStyleXfs count="1"><xf numFmtId="0" fontId="0" fillId="0" borderId="0"/></cellStyleXfs><cellXfs count="3"><xf numFmtId="0" fontId="0" fillId="0" borderId="0" xfId="0"/><xf numFmtId="0" fontId="1" fillId="0" borderId="0" xfId="0" applyFont="1"/><xf numFmtId="4" fontId="0" fillId="0" borderId="0" xfId="0" applyNumberFormat="1"/></cellXfs><cellStyles count="1"><cellStyle name="Normal" xfId="0" builtinId="0"/></cellStyles></styleSheet>`);
  if (shared) {
    put(parts, 'xl/sharedStrings.xml', `<sst xmlns="${NS.s}" count="${sharedReferences}" uniqueCount="${sharedStrings.length}">${sharedStrings.join('')}</sst>`);
    types['xl/sharedStrings.xml'] = 'spreadsheetml.sharedStrings';
    bookRels.push({ id: 'rId51', type: 'sharedStrings', target: 'sharedStrings.xml' });
  }
  put(parts, 'xl/_rels/workbook.xml.rels', rels(bookRels));
  put(parts, '_rels/.rels', rels([{ id: 'rId1', type: 'officeDocument', target: 'xl/workbook.xml' }]));
  put(parts, '[Content_Types].xml', contentTypes(types));
  put(parts, 'customXml/fixture.xml', '<fixture xmlns="urn:imprint-doc-gen:fixture"><preserve>Unrelated package content must survive unchanged.</preserve></fixture>');
  return perturb(parts, disorder);
}
