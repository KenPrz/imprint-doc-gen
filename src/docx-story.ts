import { filterArray, type Formatter, type ArrayFilter } from './expressions.ts';
import { standardBindings, parseStandardPath, inferStandardLoops, scopedStandardPath, standardLookup, type RepeatRegion } from './standard-syntax.ts';
import { RenderError } from './errors.ts';
import { limits } from './package.ts';
import { checkPath, lookup, scalar, type Scope } from './scope.ts';
import { ns, indexXml, ancestor, applyEdits, decodeXml, escapeXml, type ElementSpan, type AttributeSpan, type Edit } from './xml.ts';

interface TextSpan { node: ElementSpan; value: string; offset: number }
interface Binding { block?: 'showBegin' | 'showEnd'; formatters?: Formatter[]; filters?: ArrayFilter[]; start: number; end: number; path: string; kind: 'value' | 'open' | 'close'; standard?: boolean }
interface ParagraphPlan { node: ElementSpan; text: string; spans: TextSpan[]; bindings: Binding[] }
export interface ImagePlan { size?: { width: number; height: number }; description: AttributeSpan; reference: AttributeSpan; path: string; start: number; end: number; standard?: boolean }
type LoopPlan = RepeatRegion;
interface Fragment { kind: 'fragment'; start: number; end: number; paragraphs: ParagraphPlan[]; images: ImagePlan[]; identities: Identity[] }
interface Loop { condition?: 'ifNEM'; filters?: ArrayFilter[]; kind: 'loop'; path: string; children: Segment[]; standard?: boolean }
type Segment = Fragment | Loop;
interface Identity { attribute: AttributeSpan; hex: boolean }
export interface StoryPlan { part: string; source: string; nodes: ElementSpan[]; images: ImagePlan[]; segments: Segment[] }
export interface RenderBudget { iterations: number }
export const loopLimits = Object.freeze({ depth: 16, iterations: 10_000 });
const word2010 = 'http://schemas.microsoft.com/office/word/2010/wordml';
const drawing2010 = 'http://schemas.microsoft.com/office/word/2010/wordprocessingDrawing';

export function identityAttributes(node: ElementSpan): Identity[] {
  const result: Identity[] = [];
  for (const attribute of node.attributes) {
    if (node.uri === ns.wordDrawing && node.local === 'docPr' && attribute.local === 'id' && !attribute.uri) result.push({ attribute, hex: false });
    if (node.uri === ns.word && node.local === 'id' && node.parent?.local === 'sdtPr' && attribute.uri === ns.word && attribute.local === 'val') result.push({ attribute, hex: false });
    if ((attribute.uri === word2010 && ['paraId', 'textId'].includes(attribute.local)) || (attribute.uri === drawing2010 && ['anchorId', 'editId'].includes(attribute.local))) result.push({ attribute, hex: true });
  }
  return result;
}
function readText(source: string, node: ElementSpan): string {
  const raw = source.slice(node.openEnd, node.closeStart);
  let result = '';
  for (const token of raw.matchAll(/<!\[CDATA\[([\s\S]*?)\]\]>|<!--[\s\S]*?-->|[^<]+|</g)) {
    if (token[0] === '<') throw new RenderError('INVALID_XML', 'Nested markup inside a Word text node');
    if (token[1] !== undefined) result += token[1];
    else if (!token[0].startsWith('<!--')) result += decodeXml(token[0]);
  }
  return result;
}
function classify(path: string, part: string): Pick<Binding, 'path' | 'kind'> {
  if (path === '/each') return { kind: 'close', path: '' };
  if (path.startsWith('#each ')) { const value = path.slice(6).trim(); checkPath(value, part); return { kind: 'open', path: value }; }
  checkPath(path, part); return { kind: 'value', path };
}
function markerOnly(paragraph: ParagraphPlan, binding: Binding, nodes: ElementSpan[]): boolean {
  if (paragraph.bindings.length !== 1 || paragraph.text.slice(0, binding.start).trim() || paragraph.text.slice(binding.end).trim()) return false;
  return nodes.filter(n => n.start > paragraph.node.start && n.end < paragraph.node.end).every(n =>
    n.uri === ns.word && n.local !== 'sectPr' && (['r', 't', 'pPr', 'rPr'].includes(n.local) || ancestor(n, ns.word, 'rPr') || ancestor(n, ns.word, 'pPr')));
}
function loopsFor(paragraphs: ParagraphPlan[], nodes: ElementSpan[], part: string): LoopPlan[] {
  const stack: { paragraph: ParagraphPlan; binding: Binding }[] = [], loops: LoopPlan[] = [];
  for (const paragraph of paragraphs) for (const binding of paragraph.bindings) {
    if (binding.kind === 'value') continue;
    if (binding.kind === 'open') {
      stack.push({ paragraph, binding });
      if (stack.length > loopLimits.depth) throw new RenderError('LIMIT_EXCEEDED', 'Loop nesting exceeds 16 levels', part);
      continue;
    }
    const open = stack.pop();
    if (!open) throw new RenderError('TEMPLATE_SYNTAX', 'Unexpected /each without an opening loop', part);
    const first = open.paragraph.node, last = paragraph.node;
    const firstRow = ancestor(first, ns.word, 'tr'), lastRow = ancestor(last, ns.word, 'tr');
    let loop: LoopPlan;
    if (first !== last && first.parent === last.parent && first.parent?.uri === ns.word && ['body', 'hdr', 'ftr', 'tc', 'sdtContent', 'footnote', 'endnote', 'comment'].includes(first.parent.local) && markerOnly(open.paragraph, open.binding, nodes) && markerOnly(paragraph, binding, nodes)) {
      loop = { start: first.start, end: last.end, bodyStart: first.end, bodyEnd: last.start, path: open.binding.path, children: [] };
    } else if (firstRow && firstRow === lastRow && firstRow.parent?.uri === ns.word && firstRow.parent.local === 'tbl') {
      loop = { start: firstRow.start, end: firstRow.end, bodyStart: firstRow.start, bodyEnd: firstRow.end, path: open.binding.path, children: [] };
    } else throw new RenderError('TEMPLATE_STRUCTURE', 'Loop markers must select one table row or standalone sibling paragraphs', part);
    if (ancestor(first, ns.word, 'p')) throw new RenderError('UNSUPPORTED_TEMPLATE', 'Loops inside text boxes require a later adapter', part);
    loops.push(loop);
  }
  if (stack.length) throw new RenderError('TEMPLATE_SYNTAX', 'Unclosed #each loop', part);
  return loops;
}
/** Narrow Carbone-compatible sibling-paragraph conditional. No implicit scope
 * changes, nested blocks, repeat mixing, inline blocks or cross-container cuts. */
function conditionsFor(paragraphs: ParagraphPlan[], nodes: ElementSpan[], part: string): LoopPlan[] {
  const result: LoopPlan[] = [];
  let open: { paragraph: ParagraphPlan; binding: Binding } | undefined;
  for (const paragraph of paragraphs) for (const binding of paragraph.bindings) {
    if (!binding.block) continue;
    if (binding.block === 'showBegin') {
      if (open) throw new RenderError('UNSUPPORTED_TEMPLATE', 'Nested conditional blocks are unsupported', part);
      if (paragraph.text.slice(0, binding.start).trim()) throw new RenderError('TEMPLATE_STRUCTURE', 'showBegin must start its paragraph', part);
      open = { paragraph, binding }; continue;
    }
    if (!open || open.binding.path !== binding.path) throw new RenderError('TEMPLATE_STRUCTURE', 'showEnd must match the opening field', part);
    const first = open.paragraph.node, last = paragraph.node;
    if (first === last || first.parent !== last.parent || first.parent?.uri !== ns.word ||
        !['body', 'hdr', 'ftr', 'tc', 'sdtContent', 'footnote', 'endnote', 'comment'].includes(first.parent.local) ||
        !markerOnly(paragraph, binding, nodes) || ancestor(first, ns.word, 'p'))
      throw new RenderError('TEMPLATE_STRUCTURE', 'Conditional blocks require sibling paragraphs and a standalone showEnd', part);
    const unsafe = new Set(['sectPr', 'bookmarkStart', 'bookmarkEnd', 'permStart', 'permEnd', 'commentRangeStart', 'commentRangeEnd', 'commentReference', 'footnoteReference', 'endnoteReference', 'ins', 'del', 'moveFrom', 'moveTo']);
    if (nodes.some(n => n.start >= first.start && n.end <= last.end && n.uri === ns.word && unsafe.has(n.local)))
      throw new RenderError('UNSUPPORTED_TEMPLATE', 'Conditional blocks cannot contain section, range, revision or note anchors', part);
    // Carbone keeps the closing marker paragraph, empty, in both branches.
    result.push({ start: first.start, end: last.start, bodyStart: first.start, bodyEnd: last.start, path: binding.path, children: [], condition: 'ifNEM', standard: true });
    open = undefined;
  }
  if (open) throw new RenderError('TEMPLATE_STRUCTURE', 'Unclosed showBegin conditional block', part);
  return result;
}
function isNonempty(value: unknown): boolean {
  return value !== null && value !== undefined && value !== '' &&
    !(Array.isArray(value) && value.length === 0) &&
    !(typeof value === 'object' && !Array.isArray(value) && Object.keys(value).length === 0);
}
function nestLoops(loops: LoopPlan[], nodes: ElementSpan[], part: string): LoopPlan[] {
  const unsupported = new Set(['bookmarkStart', 'bookmarkEnd', 'permStart', 'permEnd', 'moveFrom', 'moveTo', 'ins', 'del', 'vMerge', 'dataBinding', 'pict', 'object', 'altChunk']);
  for (const loop of loops) for (const node of nodes) if (node.start >= loop.start && node.end <= loop.end && node.uri === ns.word && unsupported.has(node.local)) throw new RenderError('UNSUPPORTED_TEMPLATE', `Repeated ${node.local} requires a later cloning adapter`, part);
  const repeatedNodes=nodes.filter(n=>loops.some(loop=>n.start>=loop.bodyStart&&n.end<=loop.bodyEnd));
  if(repeatedNodes.some(n=>n.uri===ns.word&&n.local==='commentReference') && repeatedNodes.some(n=>n.uri===ns.word&&n.local==='endnoteReference'))throw new RenderError('UNSUPPORTED_TEMPLATE','Repeated comments and endnotes cannot share a story because LibreOffice misrenders their text',part);
  for (const loop of loops) {
    const comments=nodes.filter(n=>n.uri===ns.word && n.start>=loop.bodyStart && n.end<=loop.bodyEnd && ['commentRangeStart','commentRangeEnd'].includes(n.local));
    for(const note of nodes.filter(n=>n.uri===ns.word&&['footnoteReference','endnoteReference'].includes(n.local)&&n.start>=loop.bodyStart&&n.end<=loop.bodyEnd)) {
      const paragraph=ancestor(note,ns.word,'p');
      if(paragraph && nodes.some(n=>n.uri===ns.word&&['commentRangeStart','commentRangeEnd','commentReference'].includes(n.local)&&n.start>paragraph.start&&n.end<paragraph.end))throw new RenderError('UNSUPPORTED_TEMPLATE','Repeated comments and note anchors must use separate paragraphs for converter compatibility',part);
    }
    const counts=new Map<string,number>();
    for(const node of comments){const id=node.attributes.find(a=>a.uri===ns.word&&a.local==='id')?.value ?? '';counts.set(id,(counts.get(id)??0)+(node.local==='commentRangeStart'?1:-1));}
    for(const start of comments.filter(n=>n.local==='commentRangeStart')) {
      const id=start.attributes.find(a=>a.uri===ns.word&&a.local==='id')?.value;
      const end=comments.find(n=>n.local==='commentRangeEnd'&&n.attributes.some(a=>a.uri===ns.word&&a.local==='id'&&a.value===id));
      if(end && nodes.some(n=>n.start>start.start&&n.end<end.end&&n.uri===ns.word&&['footnoteReference','endnoteReference'].includes(n.local)))throw new RenderError('UNSUPPORTED_TEMPLATE','Repeated comment ranges must not enclose note references; LibreOffice can misrender their text',part);
    }
    if([...counts.values()].some(n=>n!==0))throw new RenderError('UNSUPPORTED_TEMPLATE','Repeated comment ranges must be contained within one repeat region',part);
  }
  const roots: LoopPlan[] = [], parents: LoopPlan[] = [];
  for (const loop of loops.sort((a, b) => a.start - b.start || b.end - a.end)) {
    while (parents.length && loop.start >= parents.at(-1)!.end) parents.pop();
    const parent = parents.at(-1);
    if (parent) {
      if (loop.start < parent.bodyStart || loop.end > parent.bodyEnd || (loop.start === parent.start && loop.end === parent.end)) throw new RenderError('TEMPLATE_STRUCTURE', 'Loop regions overlap or select the same row', part);
      parent.children.push(loop);
    } else roots.push(loop);
    parents.push(loop);
  }
  return roots;
}
export function compileStory(part: string, source: string): StoryPlan {
  const nodes = indexXml(source, part), grouped = new Map<ElementSpan, TextSpan[]>();
  for (const node of nodes) if (node.uri === ns.word && node.local === 't') {
    const paragraph = ancestor(node, ns.word, 'p');
    if (!paragraph) continue;
    const spans = grouped.get(paragraph) ?? [];
    spans.push({ node, value: readText(source, node), offset: 0 }); grouped.set(paragraph, spans);
  }
  const paragraphs: ParagraphPlan[] = [];
  for (const [node, spans] of grouped) {
    let logical = '';
    for (const span of spans) { span.offset = logical.length; logical += span.value; }
    const bindings: Binding[] = []; let cursor = 0;
    while (true) {
      const start = logical.indexOf('{{', cursor); if (start < 0) break;
      const close = logical.indexOf('}}', start + 2);
      if (close < 0 || logical.slice(start + 2, close).includes('{{')) throw new RenderError('TEMPLATE_SYNTAX', 'Unclosed or nested placeholder', part);
      const binding = classify(logical.slice(start + 2, close).trim(), part), end = close + 2;
      const first = spans.find(s => s.offset <= start && s.offset + s.value.length > start)!;
      const last = spans.find(s => s.offset < end && s.offset + s.value.length >= end)!;
      if (nodes.some(n => n.start >= first.node.end && n.end <= last.node.start && n.uri === ns.word && ['br', 'tab', 'fldChar', 'instrText'].includes(n.local))) throw new RenderError('UNSUPPORTED_TEMPLATE', 'Placeholder crosses a field, line break or tab', part);
      bindings.push({ start, end, ...binding }); cursor = end;
    }
    bindings.push(...standardBindings(logical, part)); bindings.sort((a, b) => a.start - b.start);
    for (const binding of bindings.filter(b => b.standard)) {
      const first = spans.find(s => s.offset <= binding.start && s.offset + s.value.length > binding.start)!;
      const last = spans.find(s => s.offset < binding.end && s.offset + s.value.length >= binding.end)!;
      if (nodes.some(n => n.start >= first.node.end && n.end <= last.node.start && n.uri === ns.word && ['br', 'tab', 'fldChar', 'instrText'].includes(n.local))) throw new RenderError('UNSUPPORTED_TEMPLATE', 'Placeholder crosses a field, line break or tab', part);
    }
    if (bindings.length) paragraphs.push({ node, text: logical, spans, bindings });
  }
  const conditions = conditionsFor(paragraphs, nodes, part);
  const standardLoops = inferStandardLoops(paragraphs, part);
  const nativeLoops = loopsFor(paragraphs, nodes, part);
  if (conditions.length && (standardLoops.length || nativeLoops.length)) throw new RenderError('UNSUPPORTED_TEMPLATE', 'Conditional blocks cannot share a story with repetition markers', part);
  if (standardLoops.length && nativeLoops.length) throw new RenderError('UNSUPPORTED_TEMPLATE', 'Do not mix standard and legacy repetition markers in one story', part);
  if (nativeLoops.length) for (const paragraph of paragraphs) for (const binding of paragraph.bindings) if (binding.standard) binding.path = scopedStandardPath(binding.path, paragraph.node.start, nativeLoops, part);
  const images: ImagePlan[] = [], owners = new Set<number>();
  for (const node of nodes) {
    if (!((node.uri === ns.wordDrawing && node.local === 'docPr') || (node.uri === ns.picture && node.local === 'cNvPr'))) continue;
    for (const description of node.attributes.filter(a => ['descr', 'title', 'name'].includes(a.local) && (a.value.includes('{{image:') || /^\s*\{d\./.test(a.value)))) {
      const binding = /^\s*\{\{image:([^{}]+)\}\}\s*$/.exec(description.value);
      const standard = !binding;
      const match = /^\s*\{([^{}]+)\}\s*$/.exec(description.value);
      if (!binding && !match) throw new RenderError('TEMPLATE_SYNTAX', 'Image binding must occupy the whole picture description', part);
      const path = binding ? binding[1]!.trim() : scopedStandardPath(parseStandardPath(match![1]!, part), node.start, [...standardLoops, ...nativeLoops], part); checkPath(path, part);
      const owner = ancestor(node, ns.wordDrawing, 'inline') ?? ancestor(node, ns.wordDrawing, 'anchor');
      if (!owner || owners.has(owner.start)) throw new RenderError('UNSUPPORTED_TEMPLATE', 'Ambiguous image binding', part);
      const blips = nodes.filter(n => n.uri === ns.drawing && n.local === 'blip' && n.start > owner.start && n.end < owner.end);
      const reference = blips[0]?.attributes.find(a => a.uri === ns.relationship && a.local === 'embed');
      if (blips.length !== 1 || !reference) throw new RenderError('UNSUPPORTED_TEMPLATE', 'Expected one embedded picture per image binding', part);
      const extent = nodes.find(n => n.uri === ns.wordDrawing && n.local === 'extent' && n.start > owner.start && n.end < owner.end);
      const size = extent ? { width: Math.max(1, Math.round(Number(extent.attributes.find(a => a.local === 'cx')?.value) / 9525)), height: Math.max(1, Math.round(Number(extent.attributes.find(a => a.local === 'cy')?.value) / 9525)) } : undefined;
      owners.add(owner.start); images.push({ size, description, reference, path, start: owner.start, end: owner.end, standard });
    }
  }
  const loops = conditions.length ? conditions : nestLoops([...nativeLoops, ...standardLoops], nodes, part), identities = nodes.flatMap(identityAttributes);
  const region = (start: number, end: number, loops: LoopPlan[]): Segment[] => {
    const result: Segment[] = [];
    const fragment = (a: number, b: number) => {
      if (a === b) return;
      result.push({ kind: 'fragment', start: a, end: b, paragraphs: paragraphs.filter(p => p.node.start >= a && p.node.end <= b), images: images.filter(i => i.start >= a && i.end <= b), identities: identities.filter(i => i.attribute.start >= a && i.attribute.end <= b) });
    };
    let cursor = start;
    for (const loop of loops) { fragment(cursor, loop.start); result.push({ kind: 'loop', condition: loop.condition, path: loop.path, filters: loop.filters, standard: loop.standard, children: region(loop.bodyStart, loop.bodyEnd, loop.children) }); cursor = loop.end; }
    fragment(cursor, end); return result;
  };
  return { part, source, nodes, images, segments: region(0, source.length, loops) };
}
export async function renderStory(plan: StoryPlan, data: Record<string, unknown>, budget: RenderBudget, imageEdits: (image: ImagePlan, scope: Scope) => Promise<Edit[]>, freshIdentity: (hex: boolean) => string, fragment?: (nodes: ElementSpan[], scope: Scope) => Promise<Edit[]>, initialScope?: Scope): Promise<string> {
  const chunks: string[] = []; let size = 0;
  const append = (chunk: string) => {
    size += Buffer.byteLength(chunk);
    if (size > limits.entry) throw new RenderError('LIMIT_EXCEEDED', 'Rendered part exceeds 8 MiB', plan.part);
    chunks.push(chunk);
  };
  const visit = async (segments: Segment[], scope: Scope): Promise<void> => {
    for (const segment of segments) {
      if (segment.kind === 'loop') {
        if (segment.condition === 'ifNEM') {
          if (isNonempty(standardLookup(scope, segment.path, plan.part))) await visit(segment.children, scope);
          continue;
        }
        const items = segment.standard ? filterArray(standardLookup(scope, segment.path, plan.part) ?? [], segment.filters, plan.part) : lookup(scope, segment.path, plan.part);
        if (!Array.isArray(items)) throw new RenderError('INVALID_DATA', `Expected an array at ${segment.path}`, plan.part);
        for (const item of items) {
          if (++budget.iterations > loopLimits.iterations) throw new RenderError('LIMIT_EXCEEDED', 'Loop expansion exceeds 10000 iterations', plan.part);
          if (!item || typeof item !== 'object' || Array.isArray(item)) throw new RenderError('INVALID_DATA', `Expected object items at ${segment.path}`, plan.part);
          await visit(segment.children, [...scope, item]);
        }
        continue;
      }
      const edits: Edit[] = []; let insertedBytes = 0;
      const addEdit = (...values: Edit[]) => {
        for (const edit of values) {
          insertedBytes += Buffer.byteLength(edit.value);
          if (insertedBytes > limits.entry - size) throw new RenderError('LIMIT_EXCEEDED', 'Replacement text exceeds the part byte budget', plan.part);
          edits.push(edit);
        }
      };
      for (const paragraph of segment.paragraphs) {
        const localEdits = new Map<TextSpan, Edit[]>();
        for (const binding of paragraph.bindings) {
          const value = binding.kind === 'value' && !binding.block ? scalar(binding.standard ? standardLookup(scope, binding.path, plan.part, binding.formatters) : lookup(scope, binding.path, plan.part), binding.path, plan.part, Boolean(binding.standard)) : '';
          for (const span of paragraph.spans) {
            const start = Math.max(binding.start, span.offset), end = Math.min(binding.end, span.offset + span.value.length);
            if (start >= end) continue;
            const local = localEdits.get(span) ?? [];
            local.push({ start: start - span.offset, end: end - span.offset, value: start === binding.start ? value : '' }); localEdits.set(span, local);
          }
        }
        for (const [span, replacements] of localEdits) {
          const prefix = span.node.name.includes(':') ? span.node.name.split(':')[0] + ':' : '';
          const escaped = escapeXml(applyEdits(span.value, replacements, limits.entry));
          const text = paragraph.bindings.some(binding => binding.standard) ? escaped.replace(/\r\n|\r|\n|\t/g, value => `</${prefix}t><${prefix}${value === '\t' ? 'tab' : 'br'}/><${prefix}t xml:space="preserve">`) : escaped;
          addEdit({ start: span.node.openEnd, end: span.node.closeStart, value: text });
          const space = span.node.attributes.find(a => a.uri === ns.xml && a.local === 'space');
          if (space) addEdit({ start: space.start, end: space.end, value: 'preserve' });
          else addEdit({ start: span.node.openEnd - 1, end: span.node.openEnd - 1, value: ' xml:space="preserve"' });
        }
      }
      for (const image of segment.images) addEdit(...await imageEdits(image, scope));
      if (scope.length > 1) for (const { attribute, hex } of segment.identities) addEdit({ start: attribute.start, end: attribute.end, value: freshIdentity(hex) });
      const source = plan.source.slice(segment.start, segment.end);
      if (fragment) addEdit(...await fragment(plan.nodes.filter(n => n.start >= segment.start && n.end <= segment.end), scope));
      const rendered = applyEdits(source, edits.map(e => ({ ...e, start: e.start - segment.start, end: e.end - segment.start })), limits.entry - size);
      append(rendered);
    }
  };
  await visit(plan.segments, initialScope ?? [data]);
  let result = chunks.join('');
  // Removing all block-loop paragraphs must not leave an empty table cell.
  // A cell ending in a nested table also needs a trailing paragraph.
  if (plan.segments.some(s => s.kind === 'loop')) {
    const nodes = indexXml(result, plan.part), lastChildren = new Map<ElementSpan, ElementSpan>();
    for (const node of nodes) if (node.parent?.uri === ns.word && node.parent.local === 'tc' && node.local !== 'tcPr') lastChildren.set(node.parent, node);
    const edits: Edit[] = [];
    for (const node of nodes) if (node.uri === ns.word && node.local === 'tc' && (!lastChildren.has(node) || lastChildren.get(node)!.local === 'tbl')) {
      const prefix = node.name.includes(':') ? node.name.split(':')[0] + ':' : '';
      if (node.openEnd === node.end) edits.push({ start: node.start, end: node.end, value: result.slice(node.start, node.end - 2) + `><${prefix}p/></${node.name}>` });
      else edits.push({ start: node.closeStart, end: node.closeStart, value: `<${prefix}p/>` });
    }
    if (edits.length) result = applyEdits(result, edits, limits.entry);
  }
  return result;
}
