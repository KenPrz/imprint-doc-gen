import { posix } from 'node:path';
import type { Element } from '@xmldom/xmldom';
import { ns, parseXml, serialize, descendants } from './xml.ts';
import { formatNs as f } from './format-namespaces.ts';
import { readPackage, relsPath, resolvePart, limits } from './package.ts';
import { renderXmlTemplate, scopedPath } from './xml-template.ts';
import { standardBindings, standardLookup } from './standard-syntax.ts';
import { filterArray } from './expressions.ts';
import { standardImageResolver } from './image-binding.ts';
import { createRenderPackage } from './render-package.ts';
import { validateData } from './data.ts';
import { RenderError } from './errors.ts';
import type { Scope } from './scope.ts';
import type { OfficePackage } from './ooxml.ts';
import type { CompiledTemplate, RenderOptions } from './types.ts';
export async function compilePptx({ parts: original, mainPart }: OfficePackage): Promise<CompiledTemplate> {
  const main = parseXml(original.get(mainPart)!.toString(), mainPart);
  if (descendants(main, f.presentation, 'custShow').length || descendants(main, '*', 'section').length)
    throw new RenderError(
      'UNSUPPORTED_TEMPLATE',
      'Custom slide shows and sections require explicit ID rewriting',
      mainPart,
    );
  const list = descendants(parseXml(original.get(mainPart)!.toString(), mainPart), f.presentation, 'sldId');
  const links = descendants(
    parseXml(original.get(relsPath(mainPart))!.toString(), relsPath(mainPart)),
    ns.packageRelationship,
    'Relationship',
  );
  const slides = list.map((node) => {
    const id = node.getAttributeNS(ns.relationship, 'id'),
      link = links.find(
        (l) => l.getAttribute('Id') === id && l.getAttribute('Type') === `${ns.relationship}/slide`,
      );
    if (!link) throw new RenderError('INVALID_PACKAGE', 'Slide relationship is missing', mainPart);
    const part = resolvePart(mainPart, link.getAttribute('Target')!),
      source = original.get(part)!.toString(),
      document = parseXml(source, part);
    const tags = descendants(document, ns.drawing, 'p').flatMap((p) =>
      standardBindings(
        descendants(p, ns.drawing, 't')
          .map((t) => t.textContent)
          .join(''),
        part,
      ),
    );
    return { node, part, source, tags };
  });
  return Object.freeze({
    format: 'pptx' as const,
    async render(data: Record<string, unknown>, options: RenderOptions = {}) {
      validateData(data);
      const parts = new Map(original),
        attachments = Object.assign(Object.create(null), options.attachments ?? {}),
        imageKey = standardImageResolver(attachments);
      const presentation = parseXml(original.get(mainPart)!.toString(), mainPart),
        ids = descendants(presentation, f.presentation, 'sldIdLst')[0];
      if (!ids && slides.length) throw new RenderError('INVALID_PACKAGE', 'Missing slide list');
      const rels = parseXml(original.get(relsPath(mainPart))!.toString(), relsPath(mainPart)),
        types = parseXml(original.get('[Content_Types].xml')!.toString(), '[Content_Types].xml');
      const newLinks = new Set<string>();
      const replacements = new Map<string, string>(),
        retired = new Set<string>();
      let clonedBytes = 0;
      const usedParts = new Set(parts.keys()),
        usedLinks = new Set(links.map((l) => l.getAttribute('Id')));
      let next = 1,
        count = 0;
      const newPart = (folder: string) => {
        let part: string;
        do {
          part = `${folder}/imprint-${next++}.xml`;
        } while (usedParts.has(part));
        usedParts.add(part);
        return part;
      };
      const work: { part: string; source: string; scope: Scope; contexts: string[] }[] = [];
      const addType = (part: string, type: string) => {
        const n = types.createElementNS(ns.contentTypes, 'Override');
        n.setAttribute('PartName', '/' + part);
        n.setAttribute('ContentType', type);
        types.documentElement!.appendChild(n);
      };
      const addSlide = (slide: (typeof slides)[number], scope: Scope, contexts: string[]) => {
        if (++count > 1000) throw new RenderError('LIMIT_EXCEEDED', 'Presentation exceeds 1000 slides');
        clonedBytes += Buffer.byteLength(slide.source);
        if (clonedBytes > limits.expanded)
          throw new RenderError('LIMIT_EXCEEDED', 'Expanded slides exceed package budget');
        const part = newPart(posix.join(posix.dirname(mainPart), 'slides'));
        parts.set(part, Buffer.from(slide.source));
        addType(part, 'application/vnd.openxmlformats-officedocument.presentationml.slide+xml');
        if (!replacements.has(slide.part)) replacements.set(slide.part, part);
        let rid: string;
        do {
          rid = `imprintSlide${next++}`;
        } while (usedLinks.has(rid));
        usedLinks.add(rid);
        newLinks.add(rid);
        const link = rels.createElementNS(ns.packageRelationship, 'Relationship');
        link.setAttribute('Id', rid);
        link.setAttribute('Type', `${ns.relationship}/slide`);
        link.setAttribute('Target', posix.relative(posix.dirname(mainPart), part));
        rels.documentElement!.appendChild(link);
        const node = presentation.createElementNS(f.presentation, 'p:sldId');
        node.setAttribute('id', String(255 + count));
        node.setAttributeNS(ns.relationship, 'r:id', rid);
        ids!.appendChild(node);
        const oldRels = original.get(relsPath(slide.part));
        if (oldRels) {
          const relationships = parseXml(oldRels.toString(), relsPath(slide.part));
          for (const relation of descendants(relationships, ns.packageRelationship, 'Relationship')) {
            if (relation.getAttribute('TargetMode') === 'External') continue;
            const target = resolvePart(slide.part, relation.getAttribute('Target')!);
            if (relation.getAttribute('Type') === `${ns.relationship}/notesSlide`) {
              const note = newPart(posix.join(posix.dirname(mainPart), 'notesSlides')),
                source = original.get(target)!.toString();
              clonedBytes += Buffer.byteLength(source);
              if (clonedBytes > limits.expanded)
                throw new RenderError('LIMIT_EXCEEDED', 'Expanded notes exceed package budget');
              if (!replacements.has(target)) replacements.set(target, note);
              parts.set(note, Buffer.from(source));
              addType(note, 'application/vnd.openxmlformats-officedocument.presentationml.notesSlide+xml');
              relation.setAttribute('Target', posix.relative(posix.dirname(part), note));
              const old = original.get(relsPath(target));
              if (old) {
                const noteRels = parseXml(old.toString(), relsPath(target));
                for (const r of descendants(noteRels, ns.packageRelationship, 'Relationship'))
                  if (r.getAttribute('TargetMode') !== 'External')
                    r.setAttribute(
                      'Target',
                      posix.relative(
                        posix.dirname(note),
                        r.getAttribute('Type') === `${ns.relationship}/slide`
                          ? part
                          : resolvePart(target, r.getAttribute('Target')!),
                      ),
                    );
                parts.set(relsPath(note), serialize(noteRels));
              }
              work.push({ part: note, source, scope, contexts });
            } else relation.setAttribute('Target', posix.relative(posix.dirname(part), target));
          }
          parts.set(relsPath(part), serialize(relationships));
        }
        work.push({ part, source: slide.source, scope, contexts });
      };
      if (ids) while (ids.firstChild) ids.removeChild(ids.firstChild);
      for (let n = 0; n < slides.length; n++) {
        const slide = slides[n]!,
          nextSlide = slides[n + 1],
          candidate = slide.tags.find((t) => t.path.includes('[i]')),
          arrayPath = candidate?.path.split('[i]')[0];
        const repeated =
          arrayPath &&
          nextSlide &&
          nextSlide.tags.length &&
          nextSlide.tags.every(
            (t) => t.path === `${arrayPath}[i+1]` || t.path.startsWith(`${arrayPath}[i+1].`),
          ) &&
          !slide.tags.some((t) => t.path.includes(`${arrayPath}[i+1]`));
        if (repeated) {
          const items = filterArray(
            standardLookup([data], arrayPath!, slide.part) ?? [],
            candidate!.filters?.filter((f) => f.arrayPath === arrayPath),
            slide.part,
          );
          if (!Array.isArray(items))
            throw new RenderError('INVALID_DATA', 'Expected slide array', slide.part);
          for (const item of items) {
            if (!item || typeof item !== 'object' || Array.isArray(item))
              throw new RenderError('INVALID_DATA', 'Expected object slide items');
            addSlide(slide, [data, item], [arrayPath!]);
          }
          n++;
        } else addSlide(slide, [data], []);
      }
      for (const slide of slides) {
        retired.add(slide.part);
        retired.add(relsPath(slide.part));
        const bytes = original.get(relsPath(slide.part));
        if (bytes)
          for (const relation of descendants(
            parseXml(bytes.toString(), relsPath(slide.part)),
            ns.packageRelationship,
            'Relationship',
          ))
            if (relation.getAttribute('Type') === `${ns.relationship}/notesSlide`) {
              const note = resolvePart(slide.part, relation.getAttribute('Target')!);
              retired.add(note);
              retired.add(relsPath(note));
            }
      }
      for (const part of retired) parts.delete(part);
      for (const entry of descendants(types, ns.contentTypes, 'Override'))
        if (retired.has(entry.getAttribute('PartName')!.slice(1))) entry.parentNode!.removeChild(entry);
      // Retarget preserved internal links to the first generated copy.
      for (const [path, bytes] of parts)
        if (path.endsWith('.rels') && path !== relsPath(mainPart)) {
          const owner =
              path === '_rels/.rels'
                ? ''
                : posix.join(posix.dirname(posix.dirname(path)), posix.basename(path).slice(0, -5)),
            document = parseXml(bytes.toString(), path);
          let changed = false;
          for (const relation of descendants(document, ns.packageRelationship, 'Relationship'))
            if (relation.getAttribute('TargetMode') !== 'External') {
              const target = resolvePart(owner, relation.getAttribute('Target')!);
              if (retired.has(target)) {
                const replacement = replacements.get(target);
                if (!replacement)
                  throw new RenderError(
                    'UNSUPPORTED_TEMPLATE',
                    'Internal link targets a removed marker slide',
                    path,
                  );
                relation.setAttribute('Target', posix.relative(posix.dirname(owner), replacement));
                changed = true;
              }
            }
          if (changed) parts.set(path, serialize(document));
        }
      if (parts.size > limits.entries)
        throw new RenderError('LIMIT_EXCEEDED', 'Expanded presentation has too many parts');
      // Replace original slide relationships with generated slide relationships.
      for (const link of descendants(rels, ns.packageRelationship, 'Relationship'))
        if (
          link.getAttribute('Type') === `${ns.relationship}/slide` &&
          !newLinks.has(link.getAttribute('Id')!)
        )
          link.parentNode!.removeChild(link);
      parts.set(mainPart, serialize(presentation));
      parts.set(relsPath(mainPart), serialize(rels));
      parts.set('[Content_Types].xml', serialize(types));
      for (const [part, bytes] of parts)
        if (/^ppt\/(slideMasters|slideLayouts|notesMasters)\/[^/]+\.xml$/.test(part))
          work.push({ part, source: bytes.toString(), scope: [data], contexts: [] });
      const output = createRenderPackage(parts, mainPart, attachments, options.images);
      for (const plan of work)
        output.setPart(
          plan.part,
          await renderXmlTemplate(plan.source, plan.part, data, {
            scope: plan.scope,
            contexts: plan.contexts,
            picture: async (node, scope, contexts) => {
              if (node.namespaceURI !== f.presentation || node.localName !== 'pic') return;
              const descriptor = descendants(node, f.presentation, 'cNvPr')[0],
                blip = descendants(node, ns.drawing, 'blip')[0];
              if (!descriptor || !blip) return;
              const text = descriptor.getAttribute('descr') ?? '',
                tags = standardBindings(text, plan.part);
              if (!tags.length) return;
              if (tags.length !== 1 || tags[0]!.start !== 0 || tags[0]!.end !== text.length)
                throw new RenderError(
                  'TEMPLATE_SYNTAX',
                  'Picture description must contain one binding',
                  plan.part,
                );
              const key = imageKey(scopedPath(tags[0]!.path, contexts, plan.part), scope, plan.part);
              descriptor.setAttribute('descr', key ? 'Image attachment' : 'Image placeholder');
              if (!key) return;
              const extent = descendants(node, ns.drawing, 'ext')[0],
                size = extent
                  ? {
                      width: Math.max(1, Math.round(Number(extent.getAttribute('cx')) / 9525)),
                      height: Math.max(1, Math.round(Number(extent.getAttribute('cy')) / 9525)),
                    }
                  : undefined;
              blip.setAttributeNS(ns.relationship, 'r:embed', await output.addImage(plan.part, key, size));
            },
          }),
        );
      return output.finish();
    },
  });
}
