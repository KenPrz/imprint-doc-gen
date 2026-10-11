import { carboneBindings } from './carbone.ts';
import { RenderError } from './errors.ts';
import { checkPath, lookup, scalar, type Scope } from './scope.ts';
import { limits } from './package.ts';
import { ns, applyEdits, decodeXml, escapeXml, type ElementSpan, type Edit } from './xml.ts';

interface Binding { start: number; end: number; path: string; kind: 'value' | 'open' | 'close' }
interface TextSpan { node: ElementSpan; value: string; offset: number }
export interface CellText { text: string; spans: TextSpan[]; bindings: Binding[]; whole?: string }
export const decodeXstring = (text: string) => text.replace(/_x([\da-fA-F]{4})_/g, (_, hex: string) => String.fromCharCode(parseInt(hex, 16)));
export const encodeXstring = (text: string) => escapeXml(text.replace(/_x[\da-fA-F]{4}_/g, value => '_x005F_' + value.slice(1)).replace(/[\u0000-\u0008\u000b-\u000f\u0010-\u001f\ufffe\uffff]/g, value => `_x${value.charCodeAt(0).toString(16).toUpperCase().padStart(4, '0')}_`));
export function readLeaf(source: string, node: ElementSpan): string {
  if (node.openEnd === node.end) return '';
  let result = '';
  for (const token of source.slice(node.openEnd, node.closeStart).matchAll(/<!\[CDATA\[([\s\S]*?)\]\]>|<!--[\s\S]*?-->|[^<]+|</g)) {
    if (token[0] === '<') throw new RenderError('INVALID_XML', 'Nested markup inside a spreadsheet text node');
    if (token[1] !== undefined) result += token[1];
    else if (!token[0].startsWith('<!--')) result += decodeXml(token[0]);
  }
  return result;
}
export function attributeEdit(node: ElementSpan, name: string, value: string): Edit {
  const attribute = node.attributes.find(a => a.name === name);
  if (attribute) return { start: attribute.start, end: attribute.end, value: escapeXml(value) };
  const end = node.openEnd - (node.openEnd === node.end ? 2 : 1);
  return { start: end, end, value: ` ${name}="${escapeXml(value)}"` };
}
export function compileCellText(source: string, container: ElementSpan, children: Map<ElementSpan, ElementSpan[]>, part: string): CellText {
  const spans: TextSpan[] = []; let text = '';
  const contents = children.get(container) ?? [];
  const textNodes = contents.flatMap(node => node.uri === ns.spreadsheet && node.local === 'r' ? children.get(node) ?? [] : [node]);
  for (const node of textNodes) if (node.uri === ns.spreadsheet && node.local === 't') {
    const value = decodeXstring(readLeaf(source, node));
    spans.push({ node, value, offset: text.length }); text += value;
  }
  if (carboneBindings(text, part).length) throw new RenderError('UNSUPPORTED_CARBONE_TEMPLATE', 'Carbone template compatibility currently supports DOCX; XLSX uses Imprint syntax', part);
  const bindings: Binding[] = []; let cursor = 0;
  while (true) {
    const start = text.indexOf('{{', cursor); if (start < 0) break;
    const close = text.indexOf('}}', start + 2);
    if (close < 0 || text.slice(start + 2, close).includes('{{')) throw new RenderError('TEMPLATE_SYNTAX', 'Unclosed or nested cell placeholder', part);
    const tag = text.slice(start + 2, close).trim();
    const kind = tag === '/each' ? 'close' : tag.startsWith('#each ') ? 'open' : 'value';
    const path = kind === 'open' ? tag.slice(6).trim() : kind === 'close' ? '' : tag;
    if (kind !== 'close') checkPath(path, part);
    bindings.push({ start, end: close + 2, path, kind }); cursor = close + 2;
  }
  if (bindings.length && contents.some(n => n.uri === ns.spreadsheet && ['rPh', 'phoneticPr'].includes(n.local))) throw new RenderError('UNSUPPORTED_TEMPLATE', 'Bound phonetic strings require a later adapter', part);
  const values = bindings.filter(b => b.kind === 'value');
  const residue = applyEdits(text, bindings.map(b => ({ ...b, value: '' })));
  const whole = values.length === 1 && residue === '' ? values[0]!.path : undefined;
  return { text, spans, bindings, whole };
}
export function renderCellText(plan: CellText, data: Record<string, unknown>, part: string, scope: Scope = [data]): { edits: Edit[]; value: unknown } {
  const replacements = plan.bindings.map(binding => ({ ...binding, value: binding.kind === 'value' ? scalar(lookup(scope, binding.path, part), binding.path, part, true) : '' }));
  const result = applyEdits(plan.text, replacements, limits.entry);
  if (result.length > 32767) throw new RenderError('LIMIT_EXCEEDED', 'Spreadsheet cell text exceeds 32767 UTF-16 code units', part);
  const edits: Edit[] = [];
  for (const span of plan.spans) {
    const local = replacements.filter(b => b.start < span.offset + span.value.length && b.end > span.offset).map(b => ({ start: Math.max(b.start - span.offset, 0), end: Math.min(b.end - span.offset, span.value.length), value: b.start >= span.offset ? b.value : '' }));
    if (!local.length) continue;
    edits.push({ start: span.node.openEnd, end: span.node.closeStart, value: encodeXstring(applyEdits(span.value, local, limits.entry)) });
    // Every edited text node is nonempty in the input, so insertion before '>' is safe.
    edits.push(attributeEdit(span.node, 'xml:space', 'preserve'));
  }
  return { edits, value: plan.whole ? lookup(scope, plan.whole, part) : result };
}
