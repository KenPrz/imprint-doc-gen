import { RenderError } from './errors.ts';
import { checkPath, lookup, type Scope } from './scope.ts';
import { ns, ancestor, type ElementSpan } from './xml.ts';

export interface CarboneBinding { start: number; end: number; path: string; kind: 'value' | 'open' | 'close'; carbone?: boolean }
export interface CarboneParagraph { node: ElementSpan; text: string; bindings: CarboneBinding[] }
export interface RepeatRegion { start: number; end: number; bodyStart: number; bodyEnd: number; path: string; children: RepeatRegion[]; carbone?: boolean; arrayPath?: string }

export function parseCarbonePath(tag: string, part: string): string {
  if (!/^d\.[A-Za-z_$][\w$]*(?:\[i(?:\+1)?\])?(?:\.[A-Za-z_$][\w$]*(?:\[i(?:\+1)?\])?)*$/.test(tag)) throw new RenderError('UNSUPPORTED_CARBONE_TEMPLATE', `Unsupported Carbone expression: ${tag}`, part);
  const path = tag.slice(2);
  checkPath(path.replace(/\[i(?:\+1)?\]/g, ''), part);
  return path;
}
export function carboneBindings(text: string, part: string): CarboneBinding[] {
  const bindings: CarboneBinding[] = [];
  for (const start of text.matchAll(/(?<!\{)\{(?=[dc](?:\.|\[|\}|:)|#|t\(|o\.)/g)) {
    const close = text.indexOf('}', start.index + 1);
    if (close < 0 || text.slice(start.index + 1, close).includes('{')) throw new RenderError('TEMPLATE_SYNTAX', 'Unclosed Carbone tag', part);
    const path = parseCarbonePath(text.slice(start.index + 1, close), part);
    bindings.push({ start: start.index, end: close + 1, path, kind: 'value', carbone: true });
  }
  return bindings;
}
export function carboneLookup(scope: Scope, path: string, part: string): unknown {
  try { return lookup(scope, path, part); }
  catch (error) { if (error instanceof RenderError && error.code === 'MISSING_DATA') return null; throw error; }
}
function relativePath(path: string, context: RepeatRegion[], part: string): string {
  for (let index = context.length - 1; index >= 0; index--) {
    const prefix = context[index]!.arrayPath! + '[i].';
    if (path.startsWith(prefix)) {
      const result = '../'.repeat(context.length - index - 1) + path.slice(prefix.length);
      checkPath(result, part); return result;
    }
  }
  const result = '../'.repeat(context.length) + path;
  checkPath(result, part); return result;
}
export function scopedCarbonePath(path: string, position: number, loops: RepeatRegion[], part: string): string {
  return relativePath(path, loops.filter(loop => position >= loop.bodyStart && position < loop.bodyEnd).sort((a, b) => a.start - b.start || b.end - a.end), part);
}
export function inferCarboneLoops(paragraphs: CarboneParagraph[], part: string): RepeatRegion[] {
  const occurrences = paragraphs.flatMap(paragraph => paragraph.bindings.filter(binding => binding.carbone).map(binding => ({ paragraph, binding })));
  const loops: RepeatRegion[] = [], previousEnds = new Map<string, number>(), seen = new Set<string>();
  for (const { paragraph, binding } of occurrences) {
    const increment = binding.path.indexOf('[i+1]');
    if (increment < 0) continue;
    if (binding.path.indexOf('[i', increment + 5) >= 0) throw new RenderError('UNSUPPORTED_CARBONE_TEMPLATE', 'Increment markers must end the innermost array', part);
    const arrayPath = binding.path.slice(0, increment);
    const lastRow = ancestor(paragraph.node, ns.word, 'tr');
    const last = lastRow ?? paragraph.node;
    const markerKey = `${arrayPath}:${last.start}`;
    if (seen.has(markerKey)) continue;
    seen.add(markerKey);
    const first = occurrences.find(item => item.paragraph.node.start >= (previousEnds.get(arrayPath) ?? 0) && item.paragraph.node.start < last.start && (item.binding.path === `${arrayPath}[i]` || item.binding.path.startsWith(`${arrayPath}[i].`)));
    if (!first) throw new RenderError('TEMPLATE_STRUCTURE', `Missing [i] pattern before ${arrayPath}[i+1]`, part);
    const firstRow = ancestor(first.paragraph.node, ns.word, 'tr');
    let start: ElementSpan;
    if (firstRow && lastRow && firstRow !== lastRow && firstRow.parent === lastRow.parent) start = firstRow;
    else if (!lastRow && first.paragraph.node.parent === paragraph.node.parent && paragraph.node.parent?.uri === ns.word && ['body', 'hdr', 'ftr', 'tc', 'sdtContent'].includes(paragraph.node.parent.local)) start = first.paragraph.node;
    else throw new RenderError('UNSUPPORTED_CARBONE_TEMPLATE', 'Carbone repeats must select sibling rows or paragraph blocks', part);
    if (ancestor(start, ns.word, 'p')) throw new RenderError('UNSUPPORTED_CARBONE_TEMPLATE', 'Carbone repeats in text boxes are unsupported', part);
    // The increment row/paragraph is an exemplar to discard, not another data row.
    const markerParagraphs = paragraphs.filter(p => p.node.start >= last.start && p.node.end <= last.end);
    for (const marker of markerParagraphs) if (marker.bindings.some(tag => !tag.carbone || !tag.path.startsWith(`${arrayPath}[i+1]`))) throw new RenderError('TEMPLATE_STRUCTURE', 'Mixed data bindings in the Carbone increment marker', part);
    loops.push({ start: start.start, end: last.end, bodyStart: start.start, bodyEnd: last.start, path: arrayPath, arrayPath, children: [], carbone: true });
    previousEnds.set(arrayPath, last.end);
  }
  loops.sort((a, b) => a.start - b.start || b.end - a.end);
  for (const loop of loops) {
    const parents = loops.filter(other => other !== loop && loop.start >= other.bodyStart && loop.end <= other.bodyEnd);
    if (parents.length >= 16) throw new RenderError('LIMIT_EXCEEDED', 'Loop nesting exceeds 16 levels', part);
    loop.path = relativePath(loop.arrayPath!, parents, part);
  }
  for (const { paragraph, binding } of occurrences) {
    if (loops.some(loop => paragraph.node.start >= loop.bodyEnd && paragraph.node.end <= loop.end)) continue;
    binding.path = scopedCarbonePath(binding.path, paragraph.node.start, loops, part);
  }
  return loops;
}
