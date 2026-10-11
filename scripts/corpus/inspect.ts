import assert from 'node:assert/strict';
import { posix } from 'node:path';
import type { Element, Node } from '@xmldom/xmldom';
import { NS, parse, str, elements, targets, hash, type Parts, type Format } from './common.ts';
import { amount, type Data } from './recipes.ts';

export function canonical(node: Node): unknown {
  if (node.nodeType === 3 || node.nodeType === 4) return node.nodeValue;
  if (node.nodeType !== 1) return null;
  const el = node as Element;
  const attrs = Array.from(el.attributes).filter(a => a.namespaceURI !== 'http://www.w3.org/2000/xmlns/' && !a.name.startsWith('xmlns'))
    .map(a => [`${a.namespaceURI ?? ''}:${a.localName}`, a.value]).sort(([a], [b]) => a!.localeCompare(b!));
  const children = Array.from(el.childNodes).map(canonical).filter(v => v !== null && !(typeof v === 'string' && !v.trim()));
  return { name: `${el.namespaceURI}:${el.localName}`, attrs, children };
}
const textContent = (node: Element, namespace: string) => elements(node, namespace, 't').map(t => t.textContent).join('');
const selection = (node: Element, namespace: string, names: string[]) => Object.fromEntries(names.map(name => [name, elements(node, namespace, name).map(canonical)]));
export function validatePackage(parts: Parts) {
  assert(parts.has('[Content_Types].xml'));
  assert(parts.has('_rels/.rels'));
  const ct = parse(str(parts, '[Content_Types].xml'));
  const overrides = new Map(elements(ct, NS.ct, 'Override').map(e => [e.getAttribute('PartName')!.slice(1), e.getAttribute('ContentType')]));
  const defaults = new Map(elements(ct, NS.ct, 'Default').map(e => [e.getAttribute('Extension'), e.getAttribute('ContentType')]));
  for (const name of overrides.keys()) assert(parts.has(name), `Content type refers to missing ${name}`);
  for (const [name, bytes] of parts) {
    assert(overrides.has(name) || defaults.has(name.endsWith('.rels') ? 'rels' : posix.extname(name).slice(1)) || name === '[Content_Types].xml', `No content type for ${name}`);
    if (!name.endsWith('.xml') && !name.endsWith('.rels')) continue;
    const doc = parse(bytes.toString('utf8'));
    if (name.endsWith('.rels')) {
      const ownerDir = name === '_rels/.rels' ? '' : posix.dirname(posix.dirname(name));
      const ids = new Set();
      for (const r of elements(doc, NS.rel, 'Relationship')) {
        const id = r.getAttribute('Id'); assert(id && !ids.has(id), `Duplicate relationship in ${name}`); ids.add(id);
        if (r.getAttribute('TargetMode') === 'External') continue;
        const target = r.getAttribute('Target')!;
        const resolved = target.startsWith('/') ? target.slice(1) : posix.normalize(posix.join(ownerDir, target));
        assert(parts.has(resolved), `Broken relationship ${name} -> ${resolved}`);
      }
    } else {
      const links = targets(parts, name);
      for (const element of elements(doc, '*', '*')) for (const attr of Array.from(element.attributes)) {
        if (attr.namespaceURI === NS.r && ['embed', 'id', 'link'].includes(attr.localName!)) assert(links.has(attr.value), `Unresolved relationship ${attr.value} in ${name}`);
      }
    }
  }
  for (const [name] of parts) if (/xl\/worksheets\/.*\.xml$/.test(name)) {
    const doc = parse(str(parts, name));
    const seen = new Set(); let lastRow = 0;
    const sst = parts.has('xl/sharedStrings.xml') ? elements(parse(str(parts, 'xl/sharedStrings.xml')), NS.s, 'si') : [];
    for (const row of elements(doc, NS.s, 'row')) {
      const n = Number(row.getAttribute('r')); assert(n > lastRow, `Row order in ${name}`); lastRow = n;
      for (const c of elements(row, NS.s, 'c')) {
        const ref = c.getAttribute('r')!; assert(!seen.has(ref), `Duplicate cell ${ref}`); seen.add(ref);
        assert(ref.endsWith(String(n)) && Number(ref.replace(/^[A-Z]+/, '')) === n, `Cell row mismatch ${ref}`);
        if (c.getAttribute('t') === 's') { const index = Number(elements(c, NS.s, 'v')[0]?.textContent); assert(Number.isInteger(index) && sst[index], `Invalid shared string ${ref}`); }
      }
    }
  }
}
export function snapshot(parts: Parts, format: Format): Record<string, unknown> {
  const out: Record<string, unknown> = {};
  const sharedStringElements = format === 'xlsx' && parts.has('xl/sharedStrings.xml') ? elements(parse(str(parts, 'xl/sharedStrings.xml')), NS.s, 'si') : [];
  for (const [name] of [...parts].sort(([a], [b]) => a.localeCompare(b))) {
    if (!name.endsWith('.xml')) continue;
    const doc = parse(str(parts, name)), root = doc.documentElement!;
    const links = targets(parts, name);
    const canonicalResolved = (e: Element) => {
      const clone = e.cloneNode(true) as Element;
      for (const child of [clone, ...elements(clone, '*', '*')]) for (const attr of Array.from(child.attributes)) {
        if (attr.namespaceURI === NS.r && links.has(attr.value)) child.setAttributeNS(NS.r, attr.name, links.get(attr.value)!);
      }
      return canonical(clone);
    };
    if (format === 'docx' && /^word\/(document|header\d+|footer\d+)\.xml$/.test(name)) {
      out[name] = {
        paragraphs: elements(doc, NS.w, 'p').map(p => ({
          text: textContent(p, NS.w),
          properties: elements(p, NS.w, 'pPr').map(canonicalResolved),
          // Coalesce adjacent equally styled runs so legal text fragmentation is irrelevant.
          styledText: elements(p, NS.w, 'r').reduce<{ style: string; text: string }[]>((result, r) => {
            const text = textContent(r, NS.w); if (!text) return result;
            const style = JSON.stringify(elements(r, NS.w, 'rPr').map(canonical));
            if (result.at(-1)?.style === style) result.at(-1)!.text += text; else result.push({ style, text }); return result;
          }, []),
        })),
        rowCells: elements(doc, NS.w, 'tr').map(tr => elements(tr, NS.w, 'tc').length),
        tables: elements(doc, NS.w, 'tblPr').map(canonical),
        grids: elements(doc, NS.w, 'tblGrid').map(canonical),
        cells: elements(doc, NS.w, 'tcPr').map(canonical),
        sections: elements(doc, NS.w, 'sectPr').map(canonicalResolved),
        breaks: elements(doc, NS.w, 'br').map(canonical),
        fields: elements(doc, NS.w, 'fldSimple').map(canonical),
        bookmarks: elements(doc, NS.w, 'bookmarkStart').map(e => e.getAttributeNS(NS.w, 'name')),
        controls: elements(doc, NS.w, 'sdtPr').map(canonical),
        hyperlinks: elements(doc, NS.w, 'hyperlink').map(e => ({ text: textContent(e, NS.w), target: links.get(e.getAttributeNS(NS.r, 'id')!) })),
      };
    }
    if (format === 'xlsx' && /^xl\/worksheets\/.*\.xml$/.test(name)) {
      const strings = sharedStringElements;
      const textRuns = (node: Element) => {
        const runs = elements(node, NS.s, 'r');
        if (!runs.length) return [{ style: '[]', text: textContent(node, NS.s) }];
        return runs.reduce<{ style: string; text: string }[]>((result, r) => {
          const style = JSON.stringify(elements(r, NS.s, 'rPr').map(canonical)), text = textContent(r, NS.s);
          if (result.at(-1)?.style === style) result.at(-1)!.text += text; else result.push({ style, text });
          return result;
        }, []);
      };
      out[name] = {
        cells: elements(doc, NS.s, 'c').map(c => {
          const type = c.getAttribute('t'), raw = elements(c, NS.s, 'v')[0]?.textContent ?? '';
          return { ref: c.getAttribute('r'), style: c.getAttribute('s') || '0', type: type === 's' || type === 'inlineStr' ? 'text' : type || 'n',
            value: type === 's' ? textContent(strings[Number(raw)]!, NS.s) : type === 'inlineStr' ? textContent(c, NS.s) : raw,
            textRuns: type === 's' ? textRuns(strings[Number(raw)]!) : type === 'inlineStr' ? textRuns(c) : null,
            formula: elements(c, NS.s, 'f')[0]?.textContent ?? null };
        }),
        rowGeometry: elements(doc, NS.s, 'row').map(r => [r.getAttribute('r'), r.getAttribute('ht'), r.getAttribute('customHeight')]),
        ...selection(root, NS.s, ['dimension', 'pane', 'cols', 'mergeCells', 'autoFilter', 'conditionalFormatting', 'dataValidations', 'pageMargins']),
      };
    }
    if (format === 'xlsx' && name === 'xl/workbook.xml') out[name] = {
      sheets: elements(doc, NS.s, 'sheet').map(e => ({ name: e.getAttribute('name'), state: e.getAttribute('state') || 'visible', target: links.get(e.getAttributeNS(NS.r, 'id')!) })),
      names: elements(doc, NS.s, 'definedName').map(canonical),
    };
    if (format === 'xlsx' && name.startsWith('xl/tables/')) out[name] = canonical(root);
    const drawings = elements(doc, NS.a, 'blip');
    if (drawings.length) out[`${name}:images`] = drawings.map(blip => {
      const target = links.get(blip.getAttributeNS(NS.r, 'embed')!)!;
      let owner = blip.parentNode as Element;
      while (owner.parentNode && !['inline', 'anchor', 'oneCellAnchor', 'twoCellAnchor'].includes(owner.localName!)) owner = owner.parentNode as Element;
      return { hash: hash(parts.get(target)!),
        geometry: { ...selection(owner, NS.wp, ['extent', 'positionH', 'positionV']), ...selection(owner, NS.xdr, ['from', 'to', 'ext']), transform: elements(owner, NS.a, 'xfrm').map(canonical) } };
    });
  }
  return out;
}
export function assertNoMarkers(parts: Parts, format: Format) {
  for (const [name, bytes] of parts) if (name.endsWith('.xml')) {
    const doc = parse(bytes.toString('utf8'));
    const containers = format === 'docx' ? elements(doc, NS.w, 'p') : [...elements(doc, NS.s, 'si'), ...elements(doc, NS.s, 'is')];
    for (const e of containers) assert(!textContent(e, format === 'docx' ? NS.w : NS.s).includes('{{'), `Unresolved marker in ${name}`);
    for (const e of elements(doc, '*', '*')) for (const a of Array.from(e.attributes)) assert(!a.value.includes('{{'), `Unresolved attribute in ${name}`);
  }
}
// Independent semantic checks derived from input, not from the expected file's XML builder.
export function assertGolden(parts: Parts, format: Format, level: number, data: Data) {
  assertNoMarkers(parts, format);
  if (format === 'docx') {
    const doc = parse(str(parts, 'word/document.xml'));
    const paragraphs = elements(doc, NS.w, 'p').map(p => textContent(p, NS.w));
    for (const value of [data.title, data.customer, data.spaced, data.literal, '0', 'false']) assert(paragraphs.includes(value), `Missing scalar ${value}`);
    if (amount(level, 6)) for (const item of data.items) assert(paragraphs.includes(item.name), `Missing repeated item ${item.name}`);
    if (amount(level, 7)) for (const group of data.groups) for (const item of group.items) assert(paragraphs.includes(`${group.name} / ${item.name}`), 'Nested scope mismatch');
    const imageCount = amount(level, 2) + (amount(level, 6) ? data.items.length : 0) + (amount(level, 7) ? data.groups.length * 2 : 0);
    assert.equal(elements(doc, NS.a, 'blip').length, imageCount, 'Main-document image count');
  } else {
    const sheet = snapshot(parts, 'xlsx')['xl/worksheets/sheet1.xml'] as { cells: {ref: string; type: string; value: string; formula: string | null}[] };
    const cells = new Map(sheet.cells.map(c => [c.ref, c]));
    assert.equal(cells.get('A2')?.value, data.customer); assert.equal(cells.get('B1')?.value, data.title);
    assert.equal(cells.get('A3')?.value, data.spaced); assert.equal(cells.get('A4')?.value, '=1+1'); assert.equal(cells.get('A4')?.type, 'text');
    assert.equal(cells.get('B2')?.type, 'n'); assert.equal(cells.get('B2')?.value, '0'); assert.equal(cells.get('B3')?.type, 'b'); assert.equal(cells.get('B3')?.value, '0');
    if (amount(level, 6)) data.items.forEach((item, i) => {
      const row = 40 + i;
      assert.equal(cells.get(`A${row}`)?.value, item.name);
      assert.equal(cells.get(`D${row}`)?.formula, `B${row}*C${row}`);
      assert.equal(Number(cells.get(`D${row}`)?.value), item.quantity * item.price);
    });
  }
}
export function compare(actual: Parts, expected: Parts, format: Format, preserved: string[]) {
  validatePackage(actual); assertNoMarkers(actual, format);
  for (const name of preserved) assert.equal(hash(actual.get(name) ?? Buffer.alloc(0)), hash(expected.get(name)!), `Preservation failure: ${name}`);
  assert.deepEqual(snapshot(actual, format), snapshot(expected, format));
}
