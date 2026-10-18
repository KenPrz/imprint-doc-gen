import { odsFormulaRewriter } from './ods-formulas.ts';
import { XMLSerializer } from '@xmldom/xmldom';
import type { Element, Node, Text } from '@xmldom/xmldom';
import { ns, parseXml, serialize, descendants } from './xml.ts';
import { formatNs as f } from './format-namespaces.ts';
import { standardBindings, standardLookup } from './standard-syntax.ts';
import { filterArray } from './expressions.ts';
import { scalar, checkPath, type Scope } from './scope.ts';
import { RenderError } from './errors.ts';
import { limits } from './package.ts';

const elements = (node: Node): Element[] =>
  Array.from(node.childNodes).filter((n) => n.nodeType === 1) as Element[];
const paragraph = (node: Element) =>
  (node.namespaceURI === ns.drawing && node.localName === 'p') ||
  (node.namespaceURI === f.text && ['p', 'h'].includes(node.localName!));
const unit = (node: Element) =>
  paragraph(node) ||
  (node.namespaceURI === ns.drawing && node.localName === 'tr') ||
  (node.namespaceURI === f.table && node.localName === 'table-row') ||
  (node.namespaceURI === f.draw && node.localName === 'page');
function textNodes(node: Node): Text[] {
  const out: Text[] = [];
  for (const child of Array.from(node.childNodes)) {
    if (child.nodeType === 3 || child.nodeType === 4) out.push(child as Text);
    else if (!(
      child.nodeType === 1 &&
      (child as Element).namespaceURI === f.svg &&
      (child as Element).localName === 'desc'
    ))
      out.push(...textNodes(child));
  }
  return out;
}
function paragraphs(node: Element): Element[] {
  return paragraph(node) ? [node] : descendants(node, '*', '*').filter(paragraph);
}
function bindings(node: Element, part: string) {
  return paragraphs(node).flatMap((p) =>
    standardBindings(
      textNodes(p)
        .map((n) => n.data)
        .join(''),
      part,
    ),
  );
}
export function scopedPath(path: string, contexts: string[], part: string) {
  for (let n = contexts.length - 1; n >= 0; n--) {
    const prefix = `${contexts[n]}[i].`;
    if (path.startsWith(prefix)) {
      const value = '../'.repeat(contexts.length - n - 1) + path.slice(prefix.length);
      checkPath(value, part);
      return value;
    }
  }
  const value = '../'.repeat(contexts.length) + path;
  checkPath(value, part);
  return value;
}
export interface XmlTemplateOptions {
  odsFormulas?: boolean;
  scope?: Scope;
  contexts?: string[];
  picture?: (node: Element, scope: Scope, contexts: string[]) => Promise<void>;
}
export async function renderXmlTemplate(
  source: string,
  part: string,
  data: Record<string, unknown>,
  options: XmlTemplateOptions = {},
): Promise<Buffer> {
  const document = parseXml(source, part);
  const formulas = options.odsFormulas ? odsFormulaRewriter(document) : undefined;
  let iterations = 0,
    inserted = 0,
    expanded = 0,
    nextIdentity = 1;
  const existing = new Set(
    descendants(document, '*', '*')
      .flatMap((n) => [n.getAttribute('draw:name'), n.getAttribute('xml:id')])
      .filter(Boolean),
  );
  const fresh = () => {
    let id: string;
    do {
      id = `imprint-${nextIdentity++}`;
    } while (existing.has(id));
    existing.add(id);
    return id;
  };
  const visit = async (parent: Element, scope: Scope, contexts: string[]): Promise<void> => {
    if (contexts.length > 16) throw new RenderError('LIMIT_EXCEEDED', 'Loop nesting exceeds 16', part);
    const children = elements(parent);
    for (let i = 0; i < children.length; i++) {
      const child = children[i]!;
      const tags = unit(child) ? bindings(child, part) : [];
      const unresolved = tags.find((b) => {
        let path = b.path;
        for (const context of [...contexts].reverse())
          if (path.startsWith(`${context}[i].`)) {
            path = path.slice(context.length + 4);
            break;
          }
        return path.includes('[i');
      });
      if (unresolved) {
        const increment = unresolved.path.indexOf('[i+1]');
        const activeContext = [...contexts]
          .reverse()
          .find((context) => unresolved.path.startsWith(`${context}[i].`));
        const index = unresolved.path.indexOf('[i]', activeContext ? activeContext.length + 4 : 0);
        if (increment >= 0 || index < 0)
          throw new RenderError('TEMPLATE_STRUCTURE', 'Unmatched increment marker', part);
        const arrayPath = unresolved.path.slice(0, index);
        let end = i + 1;
        for (; end < children.length; end++)
          if (
            unit(children[end]!) &&
            bindings(children[end]!, part).some(
              (b) => b.path === `${arrayPath}[i+1]` || b.path.startsWith(`${arrayPath}[i+1].`),
            )
          )
            break;
        if (end === children.length)
          throw new RenderError('TEMPLATE_STRUCTURE', `Missing increment marker for ${arrayPath}`, part);
        const marker = children[end]!,
          region = children.slice(i, end);
        if (bindings(marker, part).some((b) => !b.path.startsWith(`${arrayPath}[i+1]`)))
          throw new RenderError('TEMPLATE_STRUCTURE', 'Mixed marker fields', part);
        if (
          !formulas && region.some((n) =>
            descendants(n, f.table, 'table-cell').some((c) => c.hasAttributeNS(f.table, 'formula')),
          )
        )
          throw new RenderError(
            'UNSUPPORTED_FORMULA',
            'Formulas inside repeated OpenDocument rows need explicit reference rewriting',
            part,
          );
        const filters = region
          .flatMap((n) => bindings(n, part).flatMap((b) => b.filters ?? []))
          .filter((filter) => filter.arrayPath === arrayPath);
        const items = filterArray(
          standardLookup(scope, scopedPath(arrayPath, contexts, part), part) ?? [],
          filters,
          part,
        );
        if (!Array.isArray(items)) throw new RenderError('INVALID_DATA', `Expected array ${arrayPath}`, part);
        expanded +=
          region.reduce(
            (n, element) => n + Buffer.byteLength(new XMLSerializer().serializeToString(element)),
            0,
          ) * items.length;
        if (expanded > limits.entry)
          throw new RenderError('LIMIT_EXCEEDED', 'Repeated XML exceeds its byte budget', part);
        for (const item of items) {
          if (++iterations > 10000)
            throw new RenderError('LIMIT_EXCEEDED', 'Expansion exceeds 10000 iterations', part);
          if (!item || typeof item !== 'object' || Array.isArray(item))
            throw new RenderError('INVALID_DATA', 'Expected object array items', part);
          for (const original of region) {
            const clone = original.cloneNode(true) as Element;
            formulas?.clone(original,clone,child);
            for (const n of [clone, ...descendants(clone, '*', '*')]) {
              if (n.hasAttributeNS(f.draw, 'name')) n.setAttributeNS(f.draw, 'draw:name', fresh());
              if (n.hasAttributeNS(ns.xml, 'id'))
                throw new RenderError(
                  'UNSUPPORTED_TEMPLATE',
                  'Repeated XML ID references need an explicit adapter',
                  part,
                );
            }
            parent.insertBefore(clone, child);
            await visit(clone, [...scope, item], [...contexts, arrayPath]);
          }
        }
        for (const old of [...region, marker]) parent.removeChild(old);
        i = end;
        continue;
      }
      await visit(child, scope, contexts);
    }
    if (paragraph(parent)) {
      const nodes = textNodes(parent),
        text = nodes.map((n) => n.data).join(''),
        tags = standardBindings(text, part);
      if (tags.length) {
        const replacements = tags.map((tag) => ({
          ...tag,
          value: standardLookup(scope, scopedPath(tag.path, contexts, part), part, tag.formatters),
        }));
        if (parent.parentNode?.nodeType === 1) {
          const cell = parent.parentNode as Element;
          if (
            cell.namespaceURI === f.table &&
            cell.localName === 'table-cell' &&
            tags.length === 1 &&
            tags[0]!.start === 0 &&
            tags[0]!.end === text.length
          ) {
            const value = replacements[0]!.value;
            for (const name of ['value', 'boolean-value', 'date-value', 'string-value'])
              cell.removeAttributeNS(f.office, name);
            cell.setAttributeNS(
              f.office,
              'office:value-type',
              typeof value === 'number' ? 'float' : typeof value === 'boolean' ? 'boolean' : 'string',
            );
            if (typeof value === 'number') cell.setAttributeNS(f.office, 'office:value', String(value));
            if (typeof value === 'boolean')
              cell.setAttributeNS(f.office, 'office:boolean-value', String(value));
          }
        }
        let offset = 0;
        for (const node of nodes) {
          const start = offset;
          offset += node.data.length;
          let result = '',
            cursor = 0;
          for (const tag of replacements) {
            const from = Math.max(tag.start, start) - start,
              to = Math.min(tag.end, offset) - start;
            if (from >= to) continue;
            const value = tag.start >= start ? scalar(tag.value, tag.path, part, true) : '';
            inserted += Buffer.byteLength(value);
            if (inserted > limits.entry)
              throw new RenderError('LIMIT_EXCEEDED', 'Rendered text exceeds byte budget', part);
            result += node.data.slice(cursor, from) + value;
            cursor = to;
          }
          node.data = result + node.data.slice(cursor);
          if (parent.namespaceURI === f.text && /[\n\t]| {2,}/.test(node.data)) {
            const owner = node.parentNode!;
            for (const segment of node.data.split(/(\n|\t| {2,})/)) {
              if (!segment) continue;
              if (segment === '\n' || segment === '\t' || /^ {2,}$/.test(segment)) {
                const element = document.createElementNS(
                  f.text,
                  segment === '\n' ? 'text:line-break' : segment === '\t' ? 'text:tab' : 'text:s',
                );
                if (segment.startsWith(' ')) element.setAttributeNS(f.text, 'text:c', String(segment.length));
                owner.insertBefore(element, node);
              } else owner.insertBefore(document.createTextNode(segment), node);
            }
            owner.removeChild(node);
          } else if (parent.namespaceURI === ns.drawing && node.data.includes('\n')) {
            const textElement = node.parentNode as Element,
              run = textElement.parentNode as Element;
            if (run.namespaceURI !== ns.drawing || run.localName !== 'r')
              throw new RenderError(
                'UNSUPPORTED_TEMPLATE',
                'Multiline drawing text must use a text run',
                part,
              );
            const owner = run.parentNode!,
              lines = node.data.split('\n');
            textElement.textContent = lines[0]!;
            textElement.setAttributeNS(ns.xml, 'xml:space', 'preserve');
            let previous: Node = run;
            for (const line of lines.slice(1)) {
              const br = document.createElementNS(ns.drawing, 'a:br'),
                clone = run.cloneNode(true) as Element;
              descendants(clone, ns.drawing, 't')[0]!.textContent = line;
              owner.insertBefore(br, previous.nextSibling);
              owner.insertBefore(clone, br.nextSibling);
              previous = clone;
            }
          } else if (node.parentNode?.nodeType === 1)
            (node.parentNode as Element).setAttributeNS(ns.xml, 'xml:space', 'preserve');
        }
      }
    }
    if (options.picture) await options.picture(parent, scope, contexts);
  };
  await visit(document.documentElement!, options.scope ?? [data], options.contexts ?? []);
  formulas?.finish();
  const bytes = serialize(document);
  if (bytes.length > limits.entry)
    throw new RenderError('LIMIT_EXCEEDED', 'Rendered XML exceeds 8 MiB', part);
  return bytes;
}
