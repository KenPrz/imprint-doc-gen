import sharp from 'sharp';
import { createHash } from 'node:crypto';
import type { Document, Element } from '@xmldom/xmldom';
import { readPackage, writePackage, resolvePart, limits, type PackageParts } from './package.ts';
import { ns, descendants, parseXml, serialize, indexXml } from './xml.ts';

export interface DocxComponent {
  id: string;
  version: string;
  document: Uint8Array;
  sha256: string;
}
export interface DocxComponentAssembly {
  document: Buffer;
  lock: { id: string; version: string; sha256: string; occurrences: number }[];
}
const marker = /^\{\{>\s*([A-Za-z0-9][A-Za-z0-9_.-]{0,127})@([A-Za-z0-9][A-Za-z0-9_.-]{0,127})\s*\}\}$/;
const identity = /^[A-Za-z0-9][A-Za-z0-9_.-]{0,127}$/;
const main = 'word/document.xml';
const relpath = 'word/_rels/document.xml.rels';
const stylespath = 'word/styles.xml';
function xml(parts: PackageParts, name: string): Document {
  const bytes = parts.get(name);
  if (!bytes) throw new Error(`Missing DOCX part: ${name}`);
  const source = bytes.toString('utf8');
  indexXml(source, name);
  return parseXml(source, name);
}
function elements(node: Element): Element[] {
  return [node, ...Array.from(node.getElementsByTagName('*'))];
}
function text(node: Element): string {
  return descendants(node, ns.word, 't')
    .map((e) => e.textContent ?? '')
    .join('');
}
function body(doc: Document): Element {
  const bodies = descendants(doc, ns.word, 'body');
  if (
    doc.documentElement?.namespaceURI !== ns.word ||
    doc.documentElement.localName !== 'document' ||
    bodies.length !== 1
  )
    throw new Error('Expected a transitional DOCX main document');
  return bodies[0]!;
}
function rels(parts: PackageParts): Document {
  return parts.has(relpath)
    ? xml(parts, relpath)
    : parseXml(`<Relationships xmlns="${ns.packageRelationship}"/>`, relpath);
}
function relationships(doc: Document): Map<string, Element> {
  const result = new Map<string, Element>();
  for (const rel of descendants(doc, ns.packageRelationship, 'Relationship')) {
    const id = rel.getAttribute('Id');
    if (!id || result.has(id)) throw new Error('Missing or duplicate relationship identity');
    result.set(id, rel);
  }
  return result;
}
function appendRelationship(doc: Document, id: string, type: string, target: string) {
  const rel = doc.createElementNS(ns.packageRelationship, 'Relationship');
  rel.setAttribute('Id', id);
  rel.setAttribute('Type', `${ns.relationship}/${type}`);
  rel.setAttribute('Target', target);
  doc.documentElement!.appendChild(rel);
}
function addType(doc: Document, path: string, type: string) {
  const el = doc.createElementNS(ns.contentTypes, 'Override');
  el.setAttribute('PartName', `/${path}`);
  el.setAttribute('ContentType', type);
  doc.documentElement!.appendChild(el);
}
/** Assemble checksum-pinned, non-nested DOCX fragments at whole body paragraph markers.
 * Supports directly formatted paragraphs/tables, explicit style graphs and embedded PNG/JPEG.
 * Requires identical document defaults and, when referenced, identical themes.
 * Imports ordinary list definitions with fresh list identities for each occurrence.
 * Rejects picture bullets, notes, fields, bookmarks,
 * tracked changes, sections within fragments and non-image relationship dependencies.
 * The host owns page geometry and document settings. No automatic release approval is implied.
 */
export async function assembleDocxComponents(
  template: Uint8Array,
  components: readonly DocxComponent[],
): Promise<DocxComponentAssembly> {
  if (!Array.isArray(components) || components.length > 64)
    throw new Error('DOCX component registry exceeds 64 entries');
  const hostParts = await readPackage(template);
  let expandedInput = [...hostParts.values()].reduce((sum, b) => sum + b.length, 0),
    totalPixels = 0;
  const hostStyles = hostParts.has(stylespath) ? xml(hostParts, stylespath) : undefined;
  const defaultsSignature = (document?: Document) =>
    document
      ? descendants(document, ns.word, 'docDefaults')
          .map((e) => e.toString())
          .join('')
      : '';
  const registry = new Map<
    string,
    { item: DocxComponent; parts: PackageParts; doc: Document; styles?: Document; rels: Map<string, Element> }
  >();
  let inputBytes = template.byteLength;
  for (const item of components) {
    if (
      !item ||
      !identity.test(item.id) ||
      !identity.test(item.version) ||
      !(item.document instanceof Uint8Array) ||
      !/^[a-f0-9]{64}$/.test(item.sha256)
    )
      throw new Error('Invalid pinned DOCX component');
    const key = `${item.id}@${item.version}`;
    if (registry.has(key)) throw new Error(`Duplicate DOCX component: ${key}`);
    inputBytes += item.document.byteLength;
    if (inputBytes > limits.expanded) throw new Error('DOCX component input budget exceeded');
    if (createHash('sha256').update(item.document).digest('hex') !== item.sha256)
      throw new Error(`DOCX component checksum mismatch: ${key}`);
    const parts = await readPackage(item.document),
      doc = xml(parts, main),
      componentBody = body(doc);
    expandedInput += [...parts.values()].reduce((sum, b) => sum + b.length, 0);
    if (expandedInput > limits.expanded) throw new Error('DOCX component expanded input budget exceeded');
    // Reject external links anywhere, even in currently unused ancillary parts.
    for (const name of parts.keys())
      if (name.endsWith('.rels'))
        for (const rel of relationships(xml(parts, name)).values()) {
          if (rel.getAttribute('TargetMode'))
            throw new Error('External DOCX component relationships are unsupported');
        }
    const packageRelationships = relationships(xml(parts, '_rels/.rels'));
    const officeDocuments = [...packageRelationships.values()].filter(
      (r) => r.getAttribute('Type') === `${ns.relationship}/officeDocument`,
    );
    if (officeDocuments.length !== 1 || resolvePart('', officeDocuments[0]!.getAttribute('Target')!) !== main)
      throw new Error('DOCX component requires canonical main document relationship');
    const mainRelationships = relationships(rels(parts));
    for (const [type, expected] of [
      ['styles', stylespath],
      ['numbering', 'word/numbering.xml'],
      ['theme', 'word/theme/theme1.xml'],
    ]) {
      const matches = [...mainRelationships.values()].filter(
        (r) => r.getAttribute('Type') === `${ns.relationship}/${type}`,
      );
      if (
        matches.length > 1 ||
        matches.some((r) => resolvePart(main, r.getAttribute('Target')!) !== expected)
      )
        throw new Error(`Noncanonical DOCX component ${type} relationship`);
    }
    for (const rel of mainRelationships.values())
      if (rel.getAttribute('Type') === `${ns.relationship}/image`) {
        const bytes = parts.get(resolvePart(main, rel.getAttribute('Target')!));
        if (!bytes) throw new Error('Missing DOCX component image');
        const decoder = sharp(bytes, { limitInputPixels: limits.pixels, failOn: 'warning' });
        const metadata = await decoder.metadata();
        if (
          !['png', 'jpeg'].includes(metadata.format ?? '') ||
          !metadata.width ||
          !metadata.height ||
          (metadata.pages ?? 1) !== 1
        )
          throw new Error('Only single-frame PNG/JPEG DOCX component images are supported');
        totalPixels += metadata.width * metadata.height;
        if (totalPixels > limits.pixels) throw new Error('DOCX component total image pixel budget exceeded');
        await decoder.stats();
      }
    const styles = parts.has(stylespath) ? xml(parts, stylespath) : undefined;
    if (defaultsSignature(styles) !== defaultsSignature(hostStyles))
      throw new Error('DOCX component document defaults differ from host defaults');
    for (const root of [
      componentBody,
      ...(styles ? [styles.documentElement!] : []),
      ...(parts.has('word/numbering.xml') ? [xml(parts, 'word/numbering.xml').documentElement!] : []),
    ])
      for (const el of elements(root)) {
        if (
          el.namespaceURI === ns.word &&
          [
            'footnoteReference',
            'endnoteReference',
            'commentReference',
            'commentRangeStart',
            'commentRangeEnd',
            'bookmarkStart',
            'bookmarkEnd',
            'fldChar',
            'instrText',
            'fldSimple',
            'altChunk',
            'ins',
            'del',
            'moveFrom',
            'moveTo',
            'sdt',
            'object',
            'pict',
            'subDoc',
            'customXml',
            'control',
          ].includes(el.localName!)
        )
          throw new Error(`Unsupported DOCX component construct: ${el.localName}`);
        if (el.localName === 'sectPr' && el.parentNode !== componentBody)
          throw new Error('DOCX component internal sections are unsupported');
        for (const attr of Array.from(el.attributes))
          if (
            /theme/i.test(attr.localName ?? '') &&
            (!parts.get('word/theme/theme1.xml') ||
              !parts
                .get('word/theme/theme1.xml')!
                .equals(hostParts.get('word/theme/theme1.xml') ?? Buffer.alloc(0)))
          )
            throw new Error('DOCX component theme dependencies require an identical host theme');
      }
    if (text(componentBody).includes('{{>'))
      throw new Error('Nested DOCX component references are unsupported (including cycles)');
    registry.set(key, { item, parts, doc, styles, rels: relationships(rels(parts)) });
  }
  const parts = hostParts,
    doc = xml(parts, main),
    host = body(doc),
    relDoc = rels(parts),
    hostRels = relationships(relDoc),
    types = xml(parts, '[Content_Types].xml');
  const stylesDoc = parts.has(stylespath)
    ? xml(parts, stylespath)
    : parseXml(`<w:styles xmlns:w="${ns.word}"/>`, stylespath);
  const numberingPath = 'word/numbering.xml';
  const numberingDoc = parts.has(numberingPath)
    ? xml(parts, numberingPath)
    : parseXml(`<w:numbering xmlns:w="${ns.word}"/>`, numberingPath);
  let changedNumbering = false;
  let nextNum =
    Math.max(
      0,
      ...descendants(numberingDoc, ns.word, 'num').map(
        (e) => Number(e.getAttributeNS(ns.word, 'numId')) || 0,
      ),
    ) + 1;
  let nextAbstract =
    Math.max(
      0,
      ...descendants(numberingDoc, ns.word, 'abstractNum').map(
        (e) => Number(e.getAttributeNS(ns.word, 'abstractNumId')) || 0,
      ),
    ) + 1;
  const styleIds = new Set(
    descendants(stylesDoc, ns.word, 'style').map((e) => e.getAttributeNS(ns.word, 'styleId')),
  );
  const styleNames = new Set(
    descendants(stylesDoc, ns.word, 'name').map((e) => e.getAttributeNS(ns.word, 'val')),
  );
  const locks = new Map<string, { id: string; version: string; sha256: string; occurrences: number }>();
  let serial = 0,
    picture = 0;
  for (const [name] of parts)
    if (name.startsWith('word/') && name.endsWith('.xml')) {
      for (const drawing of descendants(xml(parts, name), ns.wordDrawing, 'docPr')) {
        const id = Number(drawing.getAttribute('id'));
        if (!Number.isSafeInteger(id) || id < 0 || id > 2_000_000_000)
          throw new Error('Invalid host drawing identity');
        picture = Math.max(picture, id);
      }
    }
  const uniqueRel = () => {
    let id: string;
    do {
      id = `imprintComponent${++serial}`;
    } while (hostRels.has(id));
    hostRels.set(id, relDoc.documentElement!);
    return id;
  };
  let occurrence = 0,
    changedStyles = false;
  const markers = descendants(doc, ns.word, 'p').filter((p) => text(p).includes('{{>'));
  for (const paragraph of markers) {
    const match = marker.exec(text(paragraph));
    if (
      !match ||
      paragraph.parentNode !== host ||
      elements(paragraph).some(
        (e) =>
          e.namespaceURI === ns.word &&
          [
            'drawing',
            'sectPr',
            'br',
            'tab',
            'hyperlink',
            'bookmarkStart',
            'bookmarkEnd',
            'fldChar',
            'sdt',
          ].includes(e.localName!),
      )
    )
      throw new Error('DOCX component marker must occupy a whole top-level body paragraph');
    if (++occurrence > 128) throw new Error('DOCX component expansion exceeds 128 occurrences');
    const key = `${match[1]}@${match[2]}`,
      component = registry.get(key);
    if (!component) throw new Error(`Missing DOCX component: ${key}`);
    const styleMap = new Map<string, string>();
    const sourceStyles = component.styles ? descendants(component.styles, ns.word, 'style') : [];
    for (const style of sourceStyles) {
      const id = style.getAttributeNS(ns.word, 'styleId');
      if (!id || styleMap.has(id)) throw new Error('Missing or duplicate component style identity');
      let replacement = `imprint_${occurrence}_${styleMap.size}`;
      while (styleIds.has(replacement)) replacement += '_';
      styleMap.set(id, replacement);
      styleIds.add(replacement);
    }
    // Validate basedOn chains independently from link/next links, which may legally cycle.
    const styleById = new Map(sourceStyles.map((e) => [e.getAttributeNS(ns.word, 'styleId')!, e]));
    for (const id of styleMap.keys()) {
      const seen = new Set<string>();
      let current: string | undefined = id;
      while (current) {
        if (seen.has(current)) throw new Error('Cyclic DOCX component style inheritance');
        seen.add(current);
        const style = styleById.get(current);
        if (!style) throw new Error(`Missing DOCX component style: ${current}`);
        current = descendants(style, ns.word, 'basedOn')[0]?.getAttributeNS(ns.word, 'val') || undefined;
      }
    }
    const numberMap = new Map<string, string>();
    const importedNumbers: Element[] = [];
    const wantedNumbers = new Set(
      [
        ...descendants(body(component.doc), ns.word, 'numId'),
        ...sourceStyles.flatMap((s) => descendants(s, ns.word, 'numId')),
      ].map((e) => e.getAttributeNS(ns.word, 'val')!),
    );
    wantedNumbers.delete('0');
    if (wantedNumbers.size) {
      const sourceNumbering = xml(component.parts, numberingPath);
      const nums = new Map<string, Element>(),
        abstracts = new Map<string, Element>();
      for (const [local, attribute, map] of [
        ['num', 'numId', nums],
        ['abstractNum', 'abstractNumId', abstracts],
      ] as const) {
        for (const el of descendants(sourceNumbering, ns.word, local)) {
          const id = el.getAttributeNS(ns.word, attribute);
          if (!id || !/^\d+$/.test(id) || map.has(id))
            throw new Error('Invalid or duplicate DOCX numbering identity');
          map.set(id, el);
        }
      }
      const abstractMap = new Map<string, string>();
      for (const wanted of wantedNumbers) {
        const number = nums.get(wanted);
        if (!number) throw new Error(`Missing DOCX component numbering: ${wanted}`);
        const abstractId = descendants(number, ns.word, 'abstractNumId')[0]?.getAttributeNS(ns.word, 'val');
        const abstract = abstractId ? abstracts.get(abstractId) : undefined;
        if (!abstract) throw new Error('Missing DOCX component abstract numbering');
        if (
          descendants(abstract, ns.word, 'lvlPicBulletId').length ||
          descendants(abstract, ns.word, 'numStyleLink').length ||
          descendants(abstract, ns.word, 'styleLink').length
        )
          throw new Error('Picture bullets and linked numbering styles are unsupported');
        let abstractMapped = abstractMap.get(abstractId!);
        if (!abstractMapped) {
          abstractMapped = String(nextAbstract++);
          abstractMap.set(abstractId!, abstractMapped);
          const clone = numberingDoc.importNode(abstract, true) as Element;
          clone.setAttributeNS(ns.word, 'w:abstractNumId', abstractMapped);
          for (const nsid of descendants(clone, ns.word, 'nsid'))
            nsid.setAttributeNS(
              ns.word,
              'w:val',
              createHash('sha256')
                .update(`${key}:${occurrence}:${abstractId}`)
                .digest('hex')
                .slice(0, 8)
                .toUpperCase(),
            );
          importedNumbers.push(clone);
        }
        const mapped = String(nextNum++);
        numberMap.set(wanted, mapped);
        const clone = numberingDoc.importNode(number, true) as Element;
        clone.setAttributeNS(ns.word, 'w:numId', mapped);
        descendants(clone, ns.word, 'abstractNumId')[0]!.setAttributeNS(ns.word, 'w:val', abstractMapped);
        for (const level of descendants(abstract, ns.word, 'lvl')) {
          const ilvl = level.getAttributeNS(ns.word, 'ilvl')!;
          let override = descendants(clone, ns.word, 'lvlOverride').find(
            (e) => e.getAttributeNS(ns.word, 'ilvl') === ilvl,
          );
          if (!override) {
            override = numberingDoc.createElementNS(ns.word, 'w:lvlOverride');
            override.setAttributeNS(ns.word, 'w:ilvl', ilvl);
            clone.appendChild(override);
          }
          if (!descendants(override, ns.word, 'startOverride').length) {
            const start = numberingDoc.createElementNS(ns.word, 'w:startOverride');
            start.setAttributeNS(
              ns.word,
              'w:val',
              descendants(level, ns.word, 'start')[0]?.getAttributeNS(ns.word, 'val') ?? '1',
            );
            override.insertBefore(start, override.firstChild);
          }
        }
        importedNumbers.push(clone);
      }
    }
    const imageMap = new Map<string, string>();
    const remap = (root: Element, styleRoot = false) => {
      // Copy root namespace declarations to each inserted root, preserving arbitrary prefixes.
      for (const attr of Array.from(
        (styleRoot ? component.styles! : component.doc).documentElement!.attributes,
      ))
        if (attr.name === 'xmlns' || attr.prefix === 'xmlns')
          root.setAttributeNS('http://www.w3.org/2000/xmlns/', attr.name, attr.value);
      for (const el of elements(root)) {
        if (
          el.namespaceURI === ns.word &&
          ['pStyle', 'rStyle', 'tblStyle', 'basedOn', 'next', 'link'].includes(el.localName!)
        ) {
          const old = el.getAttributeNS(ns.word, 'val')!,
            mapped = styleMap.get(old);
          if (!mapped) throw new Error(`Missing DOCX component style: ${old}`);
          el.setAttributeNS(ns.word, 'w:val', mapped);
        }
        if (el.namespaceURI === ns.word && el.localName === 'numId') {
          const old = el.getAttributeNS(ns.word, 'val');
          if (old !== '0') {
            const mapped = numberMap.get(old ?? '');
            if (!mapped) throw new Error('Missing DOCX component numbering reference');
            el.setAttributeNS(ns.word, 'w:val', mapped);
          }
        }
        if (el.namespaceURI === ns.wordDrawing && el.localName === 'docPr')
          el.setAttribute('id', String(++picture));
        for (const attr of Array.from(el.attributes))
          if (['paraId', 'textId', 'anchorId', 'editId'].includes(attr.localName ?? ''))
            el.setAttributeNS(
              attr.namespaceURI,
              attr.name,
              createHash('sha256')
                .update(`${key}:${occurrence}:${++serial}`)
                .digest('hex')
                .slice(0, 8)
                .toUpperCase(),
            );
        for (const attr of Array.from(el.attributes))
          if (attr.namespaceURI === ns.relationship) {
            const old = attr.value;
            let mapped = imageMap.get(old);
            if (!mapped) {
              const rel = component.rels.get(old);
              if (
                !rel ||
                rel.getAttribute('Type') !== `${ns.relationship}/image` ||
                attr.localName !== 'embed'
              )
                throw new Error('Only embedded image relationships are supported in DOCX components');
              const sourcePath = resolvePart(main, rel.getAttribute('Target')!),
                bytes = component.parts.get(sourcePath);
              if (!bytes) throw new Error('Missing DOCX component image');
              const png = bytes.subarray(0, 8).equals(Buffer.from([137, 80, 78, 71, 13, 10, 26, 10]));
              const jpeg = bytes[0] === 255 && bytes[1] === 216 && bytes[2] === 255;
              if (!png && !jpeg) throw new Error('Only PNG/JPEG DOCX component images are supported');
              mapped = uniqueRel();
              let path = `word/media/${mapped}.${png ? 'png' : 'jpg'}`;
              while (parts.has(path)) path = path.replace(/\.(png|jpg)$/, '_.$1');
              parts.set(path, bytes);
              addType(types, path, png ? 'image/png' : 'image/jpeg');
              appendRelationship(relDoc, mapped, 'image', path.slice(5));
              imageMap.set(old, mapped);
            }
            el.setAttributeNS(ns.relationship, attr.name, mapped);
          }
      }
    };
    for (const imported of importedNumbers) {
      remap(imported);
      if (imported.localName === 'abstractNum') {
        const firstNum = descendants(numberingDoc, ns.word, 'num')[0];
        numberingDoc.documentElement!.insertBefore(imported, firstNum ?? null);
      } else numberingDoc.documentElement!.appendChild(imported);
      changedNumbering = true;
    }
    for (const style of sourceStyles) {
      const clone = stylesDoc.importNode(style, true) as Element;
      clone.setAttributeNS(ns.word, 'w:styleId', styleMap.get(style.getAttributeNS(ns.word, 'styleId')!)!);
      for (const name of descendants(clone, ns.word, 'name')) {
        let label = `${name.getAttributeNS(ns.word, 'val') ?? 'Style'} [${key}:${occurrence}]`;
        while (styleNames.has(label)) label += '_';
        styleNames.add(label);
        name.setAttributeNS(ns.word, 'w:val', label);
      }
      for (const alias of descendants(clone, ns.word, 'aliases')) alias.parentNode!.removeChild(alias);
      clone.removeAttributeNS(ns.word, 'default');
      remap(clone, true);
      stylesDoc.documentElement!.appendChild(clone);
      changedStyles = true;
    }
    for (const child of Array.from(body(component.doc).childNodes)) {
      if (child.nodeType !== 1) continue;
      const el = child as Element;
      if (el.namespaceURI === ns.word && el.localName === 'sectPr') {
        if (
          el.getElementsByTagNameNS(ns.word, 'headerReference').length ||
          el.getElementsByTagNameNS(ns.word, 'footerReference').length
        )
          throw new Error('DOCX component headers/footers are unsupported');
        continue;
      }
      if (el.namespaceURI !== ns.word || !['p', 'tbl'].includes(el.localName!))
        throw new Error('DOCX component body supports paragraphs and tables only');
      const clone = doc.importNode(el, true) as Element;
      for (const node of elements(clone)) {
        const spec =
          node.namespaceURI === ns.word
            ? (
                {
                  p: ['paragraph', 'pPr', 'pStyle'],
                  r: ['character', 'rPr', 'rStyle'],
                  tbl: ['table', 'tblPr', 'tblStyle'],
                } as Record<string, string[]>
              )[node.localName!]
            : undefined;
        if (!spec) continue;
        const defaultStyle = sourceStyles.find(
          (s) =>
            s.getAttributeNS(ns.word, 'type') === spec[0] &&
            ['1', 'true', 'on'].includes(s.getAttributeNS(ns.word, 'default') ?? ''),
        );
        if (!defaultStyle) continue;
        let props = Array.from(node.childNodes).find(
          (n) =>
            n.nodeType === 1 &&
            (n as Element).namespaceURI === ns.word &&
            (n as Element).localName === spec[1],
        ) as Element | undefined;
        if (!props) {
          props = doc.createElementNS(ns.word, `w:${spec[1]}`);
          node.insertBefore(props, node.firstChild);
        }
        if (!descendants(props, ns.word, spec[2]!).length) {
          const ref = doc.createElementNS(ns.word, `w:${spec[2]}`);
          ref.setAttributeNS(ns.word, 'w:val', defaultStyle.getAttributeNS(ns.word, 'styleId')!);
          props.insertBefore(ref, props.firstChild);
        }
      }
      remap(clone);
      host.insertBefore(clone, paragraph);
    }
    host.removeChild(paragraph);
    const lock = locks.get(key) ?? {
      id: component.item.id,
      version: component.item.version,
      sha256: component.item.sha256,
      occurrences: 0,
    };
    lock.occurrences++;
    locks.set(key, lock);
    parts.set(main, serialize(doc));
    if (parts.get(main)!.byteLength > limits.entry) throw new Error('Assembled DOCX exceeds XML budget');
  }
  if (changedNumbering) {
    if (!parts.has(numberingPath)) {
      appendRelationship(relDoc, uniqueRel(), 'numbering', 'numbering.xml');
      addType(
        types,
        numberingPath,
        'application/vnd.openxmlformats-officedocument.wordprocessingml.numbering+xml',
      );
    }
    parts.set(numberingPath, serialize(numberingDoc));
  }
  if (changedStyles) {
    if (!parts.has(stylespath)) {
      appendRelationship(relDoc, uniqueRel(), 'styles', 'styles.xml');
      addType(types, stylespath, 'application/vnd.openxmlformats-officedocument.wordprocessingml.styles+xml');
    }
    parts.set(stylespath, serialize(stylesDoc));
  }
  parts.set(main, serialize(doc));
  parts.set(relpath, serialize(relDoc));
  parts.set('[Content_Types].xml', serialize(types));
  if (
    parts.size > limits.entries ||
    [...parts.values()].some((b) => b.length > limits.entry) ||
    [...parts.values()].reduce((sum, b) => sum + b.length, 0) > limits.expanded
  )
    throw new Error('Assembled DOCX exceeds package budget');
  const output = await writePackage(parts);
  if (output.length > limits.archive) throw new Error('Assembled DOCX exceeds archive budget');
  return {
    document: output,
    lock: [...locks.values()].sort((a, b) => `${a.id}@${a.version}`.localeCompare(`${b.id}@${b.version}`)),
  };
}
