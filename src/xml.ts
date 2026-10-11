import { DOMParser, XMLSerializer } from '@xmldom/xmldom';
import type { Document, Element } from '@xmldom/xmldom';
import { RenderError } from './errors.ts';

export const ns = {
  word: 'http://schemas.openxmlformats.org/wordprocessingml/2006/main',
  relationship: 'http://schemas.openxmlformats.org/officeDocument/2006/relationships',
  packageRelationship: 'http://schemas.openxmlformats.org/package/2006/relationships',
  contentTypes: 'http://schemas.openxmlformats.org/package/2006/content-types',
  drawing: 'http://schemas.openxmlformats.org/drawingml/2006/main',
  wordDrawing: 'http://schemas.openxmlformats.org/drawingml/2006/wordprocessingDrawing',
  picture: 'http://schemas.openxmlformats.org/drawingml/2006/picture',
  xml: 'http://www.w3.org/XML/1998/namespace',
};
export const escapeXml = (text: string) => text.replaceAll('&', '&amp;').replaceAll('<', '&lt;').replaceAll('>', '&gt;').replaceAll('"', '&quot;').replaceAll("'", '&apos;');
export function parseXml(source: string, part: string): Document {
  if (/<!DOCTYPE|<!ENTITY/i.test(source)) throw new RenderError('INVALID_XML', 'DTD and entity declarations are not supported', part);
  try { return new DOMParser({ onError: (_level, message) => { throw new Error(message); } }).parseFromString(source, 'application/xml'); }
  catch { throw new RenderError('INVALID_XML', 'Malformed XML', part); }
}
export const serialize = (document: Document) => Buffer.from(new XMLSerializer().serializeToString(document));
export const descendants = (node: Document | Element, uri: string, local: string): Element[] => Array.from(node.getElementsByTagNameNS(uri, local));
export function decodeXml(value: string): string {
  return value.replace(/&(#x[\da-fA-F]+|#\d+|amp|lt|gt|quot|apos);/g, (_all, name: string) => {
    if (name.startsWith('#')) return String.fromCodePoint(name[1] === 'x' ? parseInt(name.slice(2), 16) : Number(name.slice(1)));
    return ({ amp: '&', lt: '<', gt: '>', quot: '"', apos: "'" } as Record<string, string>)[name]!;
  });
}
export interface AttributeSpan { name: string; local: string; uri: string; value: string; start: number; end: number }
export interface ElementSpan {
  name: string; local: string; uri: string; start: number; openEnd: number; closeStart: number; end: number;
  attributes: AttributeSpan[]; parent?: ElementSpan; namespaces: Map<string, string>;
}
export interface Edit { start: number; end: number; value: string }
export function applyEdits(source: string, edits: Edit[]): string {
  let boundary = source.length;
  for (const edit of [...edits].sort((a, b) => b.start - a.start || b.end - a.end)) {
    if (edit.end > boundary || edit.start > edit.end) throw new RenderError('INTERNAL', 'Overlapping XML edits');
    source = source.slice(0, edit.start) + edit.value + source.slice(edit.end); boundary = edit.start;
  }
  return source;
}
// Validate with an XML parser, then index source spans. Untouched bytes are never reserialized.
export function indexXml(source: string, part: string): ElementSpan[] {
  parseXml(source, part);
  const nodes: ElementSpan[] = [], stack: ElementSpan[] = [];
  const tokens = /<!--[\s\S]*?-->|<!\[CDATA\[[\s\S]*?\]\]>|<\?[\s\S]*?\?>|<(?:"[^"]*"|'[^']*'|[^'">])*>/g;
  for (const match of source.matchAll(tokens)) {
    const raw = match[0], start = match.index;
    if (raw.startsWith('<?') || raw.startsWith('<!')) continue;
    if (raw.startsWith('</')) { const node = stack.pop()!; node.closeStart = start; node.end = start + raw.length; continue; }
    const name = /^<([^\s/>]+)/.exec(raw)![1]!;
    const parent = stack.at(-1), namespaces = new Map(parent?.namespaces ?? [['xml', ns.xml]]);
    const attributes: AttributeSpan[] = [];
    for (const attr of raw.matchAll(/([^\s=<>]+)\s*=\s*("([^"]*)"|'([^']*)')/g)) {
      const name = attr[1]!, quote = attr[2]![0]!, value = attr[3] ?? attr[4]!;
      const valueStart = start + attr.index + attr[0].indexOf(quote) + 1;
      attributes.push({ name, local: name.split(':').at(-1)!, uri: '', value: decodeXml(value), start: valueStart, end: valueStart + value.length });
      if (name === 'xmlns') namespaces.set('', decodeXml(value));
      else if (name.startsWith('xmlns:')) namespaces.set(name.slice(6), decodeXml(value));
    }
    for (const attr of attributes) attr.uri = attr.name.includes(':') ? namespaces.get(attr.name.split(':')[0]!) ?? '' : '';
    const prefix = name.includes(':') ? name.split(':')[0]! : '';
    const node: ElementSpan = { name, local: name.split(':').at(-1)!, uri: namespaces.get(prefix) ?? '', start, openEnd: start + raw.length, closeStart: start + raw.length, end: start + raw.length, attributes, parent, namespaces };
    nodes.push(node); if (!raw.endsWith('/>')) stack.push(node);
    if (stack.length > 128 || nodes.length > 200_000) throw new RenderError('LIMIT_EXCEEDED', 'XML structure exceeds its complexity budget', part);
  }
  return nodes;
}
export function ancestor(node: ElementSpan, uri: string, local: string): ElementSpan | undefined {
  let parent = node.parent;
  while (parent) { if (parent.uri === uri && parent.local === local) return parent; parent = parent.parent; }
}
