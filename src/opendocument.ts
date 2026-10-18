import { createHash } from 'node:crypto';
import type { Element } from '@xmldom/xmldom';
import { formatNs as f } from './format-namespaces.ts';
import { ns, parseXml, serialize, descendants } from './xml.ts';
import { renderXmlTemplate, scopedPath } from './xml-template.ts';
import { standardBindings } from './standard-syntax.ts';
import { standardImageResolver } from './image-binding.ts';
import { normalizeImage } from './render-package.ts';
import { limits, writePackage } from './package.ts';
import { validateData } from './data.ts';
import { RenderError } from './errors.ts';
import type { OfficePackage } from './ooxml.ts';
import type { CompiledTemplate, RenderOptions } from './types.ts';
const size = (value: string | null) => {
  const m = /^([\d.]+)(cm|mm|in|pt|px)$/.exec(value ?? '');
  return m
    ? Math.max(
        1,
        Math.round(Number(m[1]) * { cm: 96 / 2.54, mm: 96 / 25.4, in: 96, pt: 96 / 72, px: 1 }[m[2]!]!),
      )
    : 0;
};
export async function compileOpenDocument({
  parts: original,
  format,
}: OfficePackage): Promise<CompiledTemplate> {
  for (const part of ['content.xml', 'styles.xml'])
    if (original.has(part)) {
      const doc = parseXml(original.get(part)!.toString(), part);
      for (const node of descendants(doc, '*', '*')) {
        if (
          (node.namespaceURI === f.office && ['scripts', 'event-listeners'].includes(node.localName!)) ||
          node.namespaceURI === 'urn:oasis:names:tc:opendocument:xmlns:script:1.0'
        )
          throw new RenderError('UNSUPPORTED_TEMPLATE', 'Document scripts are not supported', part);
        if (node.hasAttributeNS(f.xlink, 'href') && node.localName !== 'a') {
          const path = node.getAttributeNS(f.xlink, 'href')!.replace(/^\.\//, '');
          if (path && !path.startsWith('#') && !original.has(path))
            throw new RenderError(
              'UNSUPPORTED_TEMPLATE',
              'External or missing document assets are unsupported',
              part,
            );
        }
        if (node.namespaceURI === f.text && ['p', 'h'].includes(node.localName!))
          standardBindings(node.textContent ?? '', part);
      }
    }
  return Object.freeze({
    format,
    async render(data: Record<string, unknown>, options: RenderOptions = {}) {
      validateData(data);
      const parts = new Map(original),
        attachments = Object.assign(Object.create(null), options.attachments ?? {}),
        imageKey = standardImageResolver(attachments);
      const manifest = parseXml(parts.get('META-INF/manifest.xml')!.toString(), 'META-INF/manifest.xml');
      let changed = false;
      for (const part of ['content.xml', 'styles.xml'])
        if (parts.has(part))
          parts.set(
            part,
            await renderXmlTemplate(parts.get(part)!.toString(), part, data, {
              odsFormulas: format === 'ods',
              picture: async (node, scope, contexts) => {
                if (node.namespaceURI !== f.draw || node.localName !== 'frame') return;
                const description = descendants(node, f.svg, 'desc')[0],
                  image = descendants(node, f.draw, 'image')[0];
                if (!description || !image) return;
                const tags = standardBindings(description.textContent ?? '', part);
                if (!tags.length) return;
                if (
                  tags.length !== 1 ||
                  tags[0]!.start !== 0 ||
                  tags[0]!.end !== (description.textContent ?? '').length
                )
                  throw new RenderError(
                    'TEMPLATE_SYNTAX',
                    'Picture description must contain one binding',
                    part,
                  );
                const key = imageKey(scopedPath(tags[0]!.path, contexts, part), scope, part);
                description.textContent = key ? 'Image attachment' : 'Image placeholder';
                if (!key) return;
                if (!Object.hasOwn(attachments, key))
                  throw new RenderError('MISSING_ATTACHMENT', `Missing attachment ${key}`, part);
                const normalized = await normalizeImage(attachments[key], options.images?.[key], {
                  width: size(node.getAttributeNS(f.svg, 'width')),
                  height: size(node.getAttributeNS(f.svg, 'height')),
                });
                const path = `Pictures/imprint-${createHash('sha256').update(normalized.bytes).digest('hex')}.${normalized.extension}`;
                parts.set(path, normalized.bytes);
                if (
                  parts.size > limits.entries ||
                  [...parts.values()].reduce((n, b) => n + b.length, 0) > limits.expanded
                )
                  throw new RenderError('LIMIT_EXCEEDED', 'Images exceed package budget');
                image.setAttributeNS(f.xlink, 'xlink:href', path);
                if (
                  !descendants(manifest, f.manifest, 'file-entry').some(
                    (n) => n.getAttributeNS(f.manifest, 'full-path') === path,
                  )
                ) {
                  const entry = manifest.createElementNS(f.manifest, 'manifest:file-entry');
                  entry.setAttributeNS(f.manifest, 'manifest:full-path', path);
                  entry.setAttributeNS(f.manifest, 'manifest:media-type', normalized.contentType);
                  manifest.documentElement!.appendChild(entry);
                  changed = true;
                }
              },
            }),
          );
      if (changed) parts.set('META-INF/manifest.xml', serialize(manifest));
      if (
        parts.size > limits.entries ||
        [...parts.values()].reduce((n, b) => n + b.length, 0) > limits.expanded
      )
        throw new RenderError('LIMIT_EXCEEDED', 'Expanded document exceeds package budget');
      return writePackage(parts);
    },
  });
}
