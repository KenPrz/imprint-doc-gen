import { XMLSerializer } from '@xmldom/xmldom';
import type { Element } from '@xmldom/xmldom';
import { ns, parseXml, descendants, serialize, escapeXml, type ElementSpan, type Edit } from './xml.ts';
import { RenderError } from './errors.ts';
import type { Scope } from './scope.ts';
import { limits, relsPath, resolvePart, type PackageParts } from './package.ts';
const xml = (node: Element) => new XMLSerializer().serializeToString(node);
export function inlinePictures(source: string, freshIdentity?: () => string): string {
  const doc = parseXml(source, 'inline-images');
  let changed = false;
  let id = Math.max(
    0,
    ...descendants(doc, ns.wordDrawing, 'docPr').map((n) => Number(n.getAttribute('id')) || 0),
  );
  for (const t of descendants(doc, ns.word, 't')) {
    const match = /^\{\{inline:([\w.\[\]+/]+):(\d+):(\d+)\}\}$/.exec(t.textContent ?? '');
    if (!match) continue;
    const [, path, w, h] = match,
      width = Number(w),
      height = Number(h);
    if (width < 1 || height < 1 || width > 4096 || height > 4096)
      throw new RenderError('INVALID_IMAGE_OPTIONS', 'Inline image dimensions must be 1–4096 pixels');
    const drawing = parseXml(
      `<w:drawing xmlns:w="${ns.word}" xmlns:wp="${ns.wordDrawing}" xmlns:a="${ns.drawing}" xmlns:pic="${ns.picture}" xmlns:r="${ns.relationship}"><wp:inline><wp:extent cx="${width * 9525}" cy="${height * 9525}"/><wp:docPr id="${freshIdentity ? freshIdentity() : ++id}" name="Inline image" descr="{d.${escapeXml(path!)}}"/><a:graphic><a:graphicData uri="${ns.picture}"><pic:pic><pic:nvPicPr><pic:cNvPr id="0" name="Inline image"/><pic:cNvPicPr/></pic:nvPicPr><pic:blipFill><a:blip r:embed="imprintInlinePlaceholder"/><a:stretch><a:fillRect/></a:stretch></pic:blipFill><pic:spPr><a:xfrm><a:off x="0" y="0"/><a:ext cx="${width * 9525}" cy="${height * 9525}"/></a:xfrm><a:prstGeom prst="rect"><a:avLst/></a:prstGeom></pic:spPr></pic:pic></a:graphicData></a:graphic></wp:inline></w:drawing>`,
      'inline-image',
    );
    t.parentNode!.replaceChild(doc.importNode(drawing.documentElement!, true), t);
    changed = true;
  }
  return changed ? serialize(doc).toString() : source;
}
export function advancedCloner(
  parts: PackageParts,
  mainPart: string,
  restart: boolean,
  renderNote: (source: string, part: string, scope: Scope) => Promise<string>,
) {
  const links = parts.get(relsPath(mainPart));
  const relationships = links
    ? descendants(parseXml(links.toString(), relsPath(mainPart)), ns.packageRelationship, 'Relationship')
    : [];
  const related = (kind: string) => {
    const link = relationships.find((n) => n.getAttribute('Type') === `${ns.relationship}/${kind}`);
    return link ? resolvePart(mainPart, link.getAttribute('Target')!) : undefined;
  };
  const definitions = new Map<
    string,
    { part: string; doc: ReturnType<typeof parseXml>; next: number; originals: Map<string, Element> }
  >();
  for (const kind of ['footnote', 'endnote', 'comment']) {
    const part = related(kind + 's');
    if (!part) continue;
    const doc = parseXml(parts.get(part)!.toString(), part),
      entries = descendants(doc, ns.word, kind);
    definitions.set(kind, {
      part,
      doc,
      next: Math.max(0, ...entries.map((n) => Number(n.getAttributeNS(ns.word, 'id')))) + 1,
      originals: new Map(
        entries.map((n) => [n.getAttributeNS(ns.word, 'id')!, n.cloneNode(true) as Element]),
      ),
    });
  }
  let nextNote = Math.max(1, ...[...definitions.values()].map((state) => state.next));
  const numberingPart = related('numbering');
  const numbering = numberingPart ? parseXml(parts.get(numberingPart)!.toString(), numberingPart) : undefined;
  let nextNum =
    Math.max(
      0,
      ...(numbering
        ? descendants(numbering, ns.word, 'num').map((n) => Number(n.getAttributeNS(ns.word, 'numId')))
        : []),
    ) + 1;
  const maps = new WeakMap<Scope, Map<string, string>>();
  const clonedIds = new Map<string, Set<string>>();
  const addedBytes = new Map<string, number>();
  let numberingBytes = numberingPart ? parts.get(numberingPart)!.length : 0;
  const additions = new Map<string, string[]>();
  let changedNumbering = false;
  return {
    isDefinition(part: string) {
      return [...definitions.values()].some((s) => s.part === part);
    },
    async fragment(nodes: ElementSpan[], scope: Scope): Promise<Edit[]> {
      if (scope.length < 2) return [];
      let map = maps.get(scope);
      if (!map) {
        map = new Map();
        maps.set(scope, map);
      }
      const edits: Edit[] = [];
      for (const node of nodes) {
        if (node.uri !== ns.word) continue;
        const kind =
          node.local.startsWith('comment') &&
          ['commentReference', 'commentRangeStart', 'commentRangeEnd'].includes(node.local)
            ? 'comment'
            : node.local === 'footnoteReference'
              ? 'footnote'
              : node.local === 'endnoteReference'
                ? 'endnote'
                : undefined;
        const attr = node.attributes.find((a) => a.uri === ns.word && a.local === (kind ? 'id' : 'val'));
        if (!attr) continue;
        if (kind) {
          if (
            kind === 'comment' &&
            [...parts.keys()].some((p) => /commentsExtended|commentsIds|commentsExtensible/.test(p))
          )
            throw new RenderError('UNSUPPORTED_TEMPLATE', 'Repeated threaded comments are unsupported');
          const key = kind + ':' + attr.value;
          if (!map.has(key)) {
            const state = definitions.get(kind),
              original = state?.originals.get(attr.value);
            if (!state || !original)
              throw new RenderError('INVALID_PACKAGE', 'Missing repeated note/comment definition');
            const id = String(nextNote++);
            map.set(key, id);
            const ids = clonedIds.get(kind) ?? new Set<string>();
            ids.add(attr.value);
            clonedIds.set(kind, ids);
            const clone = original.cloneNode(true) as Element;
            clone.setAttributeNS(ns.word, 'w:id', id);
            const rendered = await renderNote(xml(clone), state.part, scope);
            const size =
              (addedBytes.get(state.part) ?? parts.get(state.part)!.length) + Buffer.byteLength(rendered);
            if (size > limits.entry)
              throw new RenderError(
                'LIMIT_EXCEEDED',
                'Repeated note definitions exceed XML budget',
                state.part,
              );
            addedBytes.set(state.part, size);
            const list = additions.get(state.part) ?? [];
            list.push(rendered);
            additions.set(state.part, list);
          }
          edits.push({ start: attr.start, end: attr.end, value: map.get(key)! });
        } else if (restart && node.local === 'numId' && attr.value !== '0') {
          const key = 'num:' + attr.value;
          if (!map.has(key)) {
            const original =
              numbering &&
              descendants(numbering, ns.word, 'num').find(
                (n) => n.getAttributeNS(ns.word, 'numId') === attr.value,
              );
            if (!original) throw new RenderError('INVALID_PACKAGE', 'Missing numbering definition');
            const clone = original.cloneNode(true) as Element,
              id = String(nextNum++);
            clone.setAttributeNS(ns.word, 'w:numId', id);
            const abstractId = descendants(clone, ns.word, 'abstractNumId')[0]?.getAttributeNS(
              ns.word,
              'val',
            );
            const abstract = descendants(numbering!, ns.word, 'abstractNum').find(
              (n) => n.getAttributeNS(ns.word, 'abstractNumId') === abstractId,
            );
            if (!abstract) throw new RenderError('INVALID_PACKAGE', 'Missing abstract numbering definition');
            for (const level of descendants(abstract, ns.word, 'lvl')) {
              const index = level.getAttributeNS(ns.word, 'ilvl')!;
              let override = descendants(clone, ns.word, 'lvlOverride').find(
                (n) => n.getAttributeNS(ns.word, 'ilvl') === index,
              );
              if (!override) {
                override = numbering!.createElementNS(ns.word, 'w:lvlOverride');
                override.setAttributeNS(ns.word, 'w:ilvl', index);
                clone.appendChild(override);
              }
              let start = descendants(override, ns.word, 'startOverride')[0];
              if (!start) {
                start = numbering!.createElementNS(ns.word, 'w:startOverride');
                override.insertBefore(start, override.firstChild);
              }
              start.setAttributeNS(ns.word, 'w:val', '1');
            }
            numberingBytes += Buffer.byteLength(xml(clone));
            if (numberingBytes > limits.entry)
              throw new RenderError('LIMIT_EXCEEDED', 'Repeated numbering exceeds XML budget');
            numbering!.documentElement!.appendChild(clone);
            map.set(key, id);
            changedNumbering = true;
          }
          edits.push({ start: attr.start, end: attr.end, value: map.get(key)! });
        }
      }
      return edits;
    },
    async finish(rendered: Map<string, Buffer>, data: Record<string, unknown>) {
      for (const [kind, state] of definitions) {
        const refs = new Set<string>(),
          originalRefs = new Set<string>();
        for (const [part, original] of parts) {
          if (!part.endsWith('.xml') || [...definitions.values()].some((v) => v.part === part)) continue;
          const doc = parseXml((rendered.get(part) ?? original).toString(), part);
          const originalDoc = rendered.has(part) ? parseXml(original.toString(), part) : doc;
          for (const n of descendants(originalDoc, ns.word, kind + 'Reference'))
            originalRefs.add(n.getAttributeNS(ns.word, 'id')!);
          for (const n of descendants(doc, ns.word, kind + 'Reference'))
            refs.add(n.getAttributeNS(ns.word, 'id')!);
        }
        for (const original of descendants(state.doc, ns.word, kind)) {
          const id = original.getAttributeNS(ns.word, 'id')!;
          // Drop definitions whose references were cloned or removed by an empty repeat.
          const cloned = clonedIds.get(kind)?.has(id);
          if ((cloned || originalRefs.has(id)) && !refs.has(id) && (kind === 'comment' || Number(id) > 0)) {
            original.parentNode!.removeChild(original);
            continue;
          }
          const result = await renderNote(xml(original), state.part, [data]);
          original.parentNode!.replaceChild(
            state.doc.importNode(parseXml(result, state.part).documentElement!, true),
            original,
          );
        }
        for (const item of additions.get(state.part) ?? [])
          state.doc.documentElement!.appendChild(
            state.doc.importNode(parseXml(item, state.part).documentElement!, true),
          );
        rendered.set(state.part, serialize(state.doc));
      }
      if (changedNumbering) rendered.set(numberingPart!, serialize(numbering!));
    },
  };
}
