import { amount, type Data } from './recipes.ts';
import { NS, esc, put, png, contentTypes, rels, perturb, type Parts } from './common.ts';

export function buildDocx(level: number, data: Data, rendered: boolean): Parts {
  const parts: Parts = new Map();
  const [fragment, pictures, tables, inline, stories, repeats, nested, sections, disorder, stress] = Array.from({ length: 10 }, (_, i) => amount(level, i + 1));
  const namespace = Object.entries(NS).filter(([key]) => ['w', 'r', 'wp', 'a', 'pic'].includes(key)).map(([k, v]) => `xmlns:${k}="${v}"`).join(' ');
  const text = (value: string, split = 1, bold = false) => {
    const chars = Array.from(value), count = Math.min(chars.length, split);
    const out = [];
    for (let i = 0; i < count; i++) out.push(`<w:r>${bold ? '<w:rPr><w:b/></w:rPr>' : ''}<w:t xml:space="preserve">${esc(chars.slice(Math.floor(i * chars.length / count), Math.floor((i + 1) * chars.length / count)).join(''))}</w:t></w:r>`);
    return out.join('');
  };
  const binding = (key: string, value: string, bold = false) => text(rendered ? value : `{{${key}}}`, rendered ? 1 : fragment, bold);
  const p = (runs: string, properties = '') => `<w:p>${properties ? `<w:pPr>${properties}</w:pPr>` : ''}${runs}</w:p>`;
  const cell = (content: string) => `<w:tc><w:tcPr><w:tcW w:w="2400" w:type="dxa"/></w:tcPr>${content}</w:tc>`;
  const table = (rows: string, columns = 3) => `<w:tbl><w:tblPr><w:tblW w:w="0" w:type="auto"/><w:tblBorders><w:top w:val="single" w:sz="4"/><w:bottom w:val="single" w:sz="4"/><w:insideH w:val="single" w:sz="4"/><w:insideV w:val="single" w:sz="4"/></w:tblBorders></w:tblPr><w:tblGrid>${'<w:gridCol w:w="2400"/>'.repeat(columns)}</w:tblGrid>${rows}</w:tbl>`;
  const imageRels = new Map<string, { id: string; type: string; target: string; external?: boolean }[]>();
  let drawingId = 0;
  const picture = (part: string, key: string, color: string, ordinal = 0) => {
    const links = imageRels.get(part) ?? [];
    const id = `rId${100 + links.length}`;
    links.push({ id, type: 'image', target: `media/${rendered ? color : 'placeholder'}.png` });
    imageRels.set(part, links);
    const cx = 914400 + ordinal * 45720, cy = 609600;
    const description = rendered ? `Attachment ${color}` : `{{image:${key}}}`;
    return `<w:r><w:drawing><wp:inline distT="0" distB="0" distL="0" distR="0"><wp:extent cx="${cx}" cy="${cy}"/><wp:docPr id="${++drawingId}" name="Picture ${drawingId}" descr="${esc(description)}"/><a:graphic><a:graphicData uri="${NS.pic}"><pic:pic><pic:nvPicPr><pic:cNvPr id="0" name="Attachment"/><pic:cNvPicPr/></pic:nvPicPr><pic:blipFill><a:blip r:embed="${id}"/><a:stretch><a:fillRect/></a:stretch></pic:blipFill><pic:spPr><a:xfrm><a:off x="0" y="0"/><a:ext cx="${cx}" cy="${cy}"/></a:xfrm><a:prstGeom prst="rect"><a:avLst/></a:prstGeom></pic:spPr></pic:pic></a:graphicData></a:graphic></wp:inline></w:drawing></w:r>`;
  };
  const main = 'word/document.xml';
  let body = p(text(`Imprint Doc Gen DOCX level ${level}`, 1, true)) + p(binding('title', data.title)) + p(binding('customer', data.customer));
  body += p(binding('spaced', data.spaced)) + p(binding('literal', data.literal)) + p(binding('count', String(data.count))) + p(binding('enabled', String(data.enabled)));
  for (let i = 0; i < pictures; i++) body += p(picture(main, i % 2 ? 'blue' : 'red', i % 2 ? 'blue' : 'red', i));
  if (tables) {
    const columns = 2 + Math.floor(tables / 4);
    body += table(Array.from({ length: tables }, (_, row) => `<w:tr>${Array.from({ length: columns }, (_, col) => cell(p(text(`R${row + 1}C${col + 1}: `) + binding('customer', data.customer)))).join('')}</w:tr>`).join(''), columns);
  }
  const mainRels = [{ id: 'rId1', type: 'styles', target: 'styles.xml' }];
  if (inline) {
    imageRels.set(main, [...(imageRels.get(main) ?? []), { id: 'rId5', type: 'hyperlink', target: 'https://example.invalid/imprint-doc-gen', external: true }]);
    for (let i = 0; i < inline; i++) body += p(`<w:bookmarkStart w:id="${i}" w:name="binding_${i}"/><w:hyperlink r:id="rId5">${binding('customer', data.customer, true)}</w:hyperlink><w:bookmarkEnd w:id="${i}"/><w:sdt><w:sdtPr><w:alias w:val="Customer"/><w:tag w:val="imprint-doc-gen-customer"/><w:id w:val="${5000 + i}"/></w:sdtPr><w:sdtContent>${binding('title', data.title)}</w:sdtContent></w:sdt>`);
  }
  const types: Record<string, string> = { [main]: 'wordprocessingml.document.main', 'word/styles.xml': 'wordprocessingml.styles' };
  if (stories) {
    mainRels.push({ id: 'rId2', type: 'header', target: 'header1.xml' }, { id: 'rId3', type: 'footer', target: 'footer1.xml' });
    for (const [kind, root] of [['header', 'hdr'], ['footer', 'ftr']] as const) {
      const part = `word/${kind}1.xml`;
      let content = Array.from({ length: stories }, () => p(binding('customer', data.customer))).join('');
      content += p(picture(part, kind === 'header' ? 'red' : 'blue', kind === 'header' ? 'red' : 'blue'));
      if (kind === 'footer') content += p('<w:fldSimple w:instr="PAGE"><w:r><w:t>1</w:t></w:r></w:fldSimple>');
      put(parts, part, `<w:${root} ${namespace}>${content}</w:${root}>`);
      types[part] = `wordprocessingml.${kind}`;
    }
  }
  if (repeats) {
    const row = (item: Data['items'][number], template: boolean) => `<w:tr>${cell(p((template ? text('{{#each items}}', fragment) : '') + binding('name', item.name)))}${cell(p(binding('quantity', String(item.quantity))))}${cell(p(picture(main, 'image', item.image)) + p(template ? text('{{/each}}', fragment) : ''))}</w:tr>`;
    body += table(rendered ? data.items.map(item => row(item, false)).join('') : row(data.items[0]!, true));
  }
  if (nested) {
    if (rendered) {
      for (const group of data.groups) {
        body += p(text(group.name, 1, true));
        for (const item of group.items) body += p(text(`${group.name} / ${item.name}`) + picture(main, 'image', item.image));
      }
    } else {
      body += p(text('{{#each groups}}', fragment)) + p(binding('name', '', true)) + p(text('{{#each items}}', fragment));
      body += p(binding('../name', '') + text(' / ') + binding('name', '') + picture(main, 'image', 'red'));
      body += p(text('{{/each}}', fragment)) + p(text('{{/each}}', fragment));
    }
  }
  const section = (landscape: boolean) => `${stories ? '<w:headerReference w:type="default" r:id="rId2"/><w:footerReference w:type="default" r:id="rId3"/>' : ''}<w:type w:val="nextPage"/><w:pgSz w:w="${landscape ? 16838 : 11906}" w:h="${landscape ? 11906 : 16838}"${landscape ? ' w:orient="landscape"' : ''}/><w:pgMar w:top="1440" w:right="1440" w:bottom="1440" w:left="1440" w:header="720" w:footer="720" w:gutter="0"/>`;
  for (let i = 0; i < sections; i++) body += p(binding('title', data.title), `<w:sectPr>${section(i % 2 === 0)}</w:sectPr>`) + p('<w:r><w:br w:type="page"/></w:r>');
  for (let i = 0; i < stress * 30; i++) body += p(text(`Stress paragraph ${i + 1}: `) + binding('title', data.title) + text(' — layout preservation '.repeat(4)));
  body += `<w:sectPr>${section(false)}</w:sectPr>`;
  put(parts, main, `<w:document ${namespace}><w:body>${body}</w:body></w:document>`);
  put(parts, 'word/styles.xml', `<w:styles xmlns:w="${NS.w}"><w:docDefaults><w:rPrDefault><w:rPr><w:rFonts w:ascii="Calibri" w:hAnsi="Calibri"/><w:sz w:val="22"/></w:rPr></w:rPrDefault><w:pPrDefault/></w:docDefaults><w:style w:type="paragraph" w:default="1" w:styleId="Normal"><w:name w:val="Normal"/></w:style></w:styles>`);
  put(parts, '_rels/.rels', rels([{ id: 'rId1', type: 'officeDocument', target: main }]));
  put(parts, 'word/_rels/document.xml.rels', rels([...mainRels, ...(imageRels.get(main) ?? [])]));
  for (const [part, links] of imageRels) if (part !== main) put(parts, `word/_rels/${part.slice(5)}.rels`, rels(links));
  if (pictures) {
    parts.set('word/media/placeholder.png', png([150, 150, 150]));
    if (rendered) { parts.set('word/media/red.png', png([220, 40, 40])); parts.set('word/media/blue.png', png([40, 80, 220])); }
  }
  put(parts, '[Content_Types].xml', contentTypes(types));
  put(parts, 'customXml/fixture.xml', '<fixture xmlns="urn:imprint-doc-gen:fixture"><preserve>Unrelated package content must survive unchanged.</preserve></fixture>');
  return perturb(parts, disorder);
}
