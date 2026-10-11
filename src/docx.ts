import { createHash } from 'node:crypto';
import { posix } from 'node:path';
import sharp from 'sharp';
import { RenderError } from './errors.ts';
import { readPackage, writePackage, resolvePart, relsPath, limits, type PackageParts } from './package.ts';
import { ns, indexXml, ancestor, applyEdits, decodeXml, escapeXml, parseXml, serialize, descendants, type ElementSpan, type AttributeSpan, type Edit } from './xml.ts';

export interface RenderOptions { attachments?: Record<string, Uint8Array> }
export interface CompiledTemplate { readonly format: 'docx'; render(data: Record<string, unknown>, options?: RenderOptions): Promise<Buffer> }
interface TextSpan { node: ElementSpan; value: string; offset: number }
interface Binding { start: number; end: number; path: string }
interface ParagraphPlan { spans: TextSpan[]; bindings: Binding[] }
interface ImagePlan { description: AttributeSpan; reference: AttributeSpan; path: string }
interface StoryPlan { part: string; source: string; paragraphs: ParagraphPlan[]; images: ImagePlan[] }
const forbidden = new Set(['__proto__', 'constructor', 'prototype']);
function checkPath(path: string, part: string) {
  if (/^[#\/^]/.test(path)) throw new RenderError('UNSUPPORTED_TEMPLATE', 'Structural loops and conditions are not implemented in milestone 1', part);
  if (!/^[A-Za-z_$][\w$]*(?:\.[A-Za-z_$][\w$]*)*$/.test(path) || path.split('.').some(p => forbidden.has(p))) throw new RenderError('TEMPLATE_EXPRESSION', `Unsupported data path: ${path}`, part);
}
function lookup(data: Record<string, unknown>, path: string, part: string): unknown {
  let value: unknown = data;
  for (const key of path.split('.')) {
    if (value === null || typeof value !== 'object' || !Object.hasOwn(value, key)) throw new RenderError('MISSING_DATA', `Missing value: ${path}`, part);
    value = (value as Record<string, unknown>)[key];
  }
  return value;
}
function scalar(value: unknown, path: string, part: string): string {
  if (value === null) return '';
  if (!['string', 'number', 'boolean'].includes(typeof value) || (typeof value === 'number' && !Number.isFinite(value))) throw new RenderError('INVALID_DATA', `Expected a scalar at ${path}`, part);
  const text = String(value);
  if (/[\u0000-\u0008\u000b\u000c\u000e-\u001f\ud800-\udfff\ufffe\uffff]/u.test(text)) throw new RenderError('INVALID_DATA', `Value contains invalid XML characters: ${path}`, part);
  if (text.length > 1_000_000) throw new RenderError('LIMIT_EXCEEDED', 'A replacement value exceeds its character budget', part);
  if (/[\r\n\t]/.test(text)) throw new RenderError('UNSUPPORTED_DATA', 'Multiline and tabbed replacements require a later rendering milestone', part);
  return text;
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
function compileStory(part: string, source: string): StoryPlan {
  const nodes = indexXml(source, part), grouped = new Map<ElementSpan, TextSpan[]>();
  for (const node of nodes) if (node.uri === ns.word && node.local === 't') {
    const paragraph = ancestor(node, ns.word, 'p');
    if (!paragraph) continue;
    const spans = grouped.get(paragraph) ?? [];
    spans.push({ node, value: readText(source, node), offset: 0 }); grouped.set(paragraph, spans);
  }
  const paragraphs: ParagraphPlan[] = [];
  for (const spans of grouped.values()) {
    let logical = '';
    for (const span of spans) { span.offset = logical.length; logical += span.value; }
    const bindings: Binding[] = [];
    let cursor = 0;
    while (true) {
      const start = logical.indexOf('{{', cursor); if (start < 0) break;
      const close = logical.indexOf('}}', start + 2);
      if (close < 0 || logical.slice(start + 2, close).includes('{{')) throw new RenderError('TEMPLATE_SYNTAX', 'Unclosed or nested placeholder', part);
      const path = logical.slice(start + 2, close).trim(); checkPath(path, part);
      const end = close + 2;
      const first = spans.find(s => s.offset <= start && s.offset + s.value.length > start)!;
      const last = spans.find(s => s.offset < end && s.offset + s.value.length >= end)!;
      if (nodes.some(n => n.start > first.node.end && n.end < last.node.start && n.uri === ns.word && ['br', 'tab', 'fldChar', 'instrText'].includes(n.local))) throw new RenderError('UNSUPPORTED_TEMPLATE', 'Placeholder crosses a field, line break or tab', part);
      bindings.push({ start, end, path }); cursor = end;
    }
    if (bindings.length) paragraphs.push({ spans, bindings });
  }
  const images: ImagePlan[] = [], owners = new Set<number>();
  for (const node of nodes) {
    if (!((node.uri === ns.wordDrawing && node.local === 'docPr') || (node.uri === ns.picture && node.local === 'cNvPr'))) continue;
    for (const description of node.attributes.filter(a => ['descr', 'title', 'name'].includes(a.local) && a.value.includes('{{image:'))) {
      const binding = /^\s*\{\{image:([^{}]+)\}\}\s*$/.exec(description.value);
      if (!binding) throw new RenderError('TEMPLATE_SYNTAX', 'Image binding must occupy the whole picture description', part);
      const path = binding[1]!.trim(); checkPath(path, part);
      const owner = ancestor(node, ns.wordDrawing, 'inline') ?? ancestor(node, ns.wordDrawing, 'anchor');
      if (!owner || owners.has(owner.start)) throw new RenderError('UNSUPPORTED_TEMPLATE', 'Ambiguous image binding', part);
      const blips = nodes.filter(n => n.uri === ns.drawing && n.local === 'blip' && n.start > owner.start && n.end < owner.end);
      const reference = blips[0]?.attributes.find(a => a.uri === ns.relationship && a.local === 'embed');
      if (blips.length !== 1 || !reference) throw new RenderError('UNSUPPORTED_TEMPLATE', 'Expected one embedded picture per image binding', part);
      owners.add(owner.start); images.push({ description, reference, path });
    }
  }
  return { part, source, paragraphs, images };
}
function validateData(data: unknown): asserts data is Record<string, unknown> {
  if (!data || typeof data !== 'object' || Array.isArray(data)) throw new RenderError('INVALID_DATA', 'Data must be a JSON object');
  let count = 0; const visiting = new WeakSet<object>();
  const visit = (value: unknown, depth: number) => {
    if (++count > 100_000 || depth > 64) throw new RenderError('LIMIT_EXCEEDED', 'Data exceeds its complexity budget');
    if (value === null || typeof value === 'string' || typeof value === 'boolean' || (typeof value === 'number' && Number.isFinite(value))) return;
    if (typeof value !== 'object' || visiting.has(value)) throw new RenderError('INVALID_DATA', 'Data must contain only acyclic JSON values');
    if (!Array.isArray(value) && ![Object.prototype, null].includes(Object.getPrototypeOf(value))) throw new RenderError('INVALID_DATA', 'Data must contain plain objects');
    visiting.add(value);
    for (const [key, descriptor] of Object.entries(Object.getOwnPropertyDescriptors(value))) {
      if (Array.isArray(value) && key === 'length') continue;
      if (forbidden.has(key) || !Object.hasOwn(descriptor, 'value')) throw new RenderError('INVALID_DATA', 'Accessors and prototype keys are not supported');
      visit(descriptor.value, depth + 1);
    }
    visiting.delete(value);
  };
  visit(data, 0);
}
async function normalizeImage(bytes: Uint8Array): Promise<{ bytes: Buffer; extension: string; contentType: string }> {
  if (!bytes.byteLength || bytes.byteLength > limits.image) throw new RenderError('LIMIT_EXCEEDED', 'Attachment must be between 1 byte and 8 MiB');
  try {
    const buffer = Buffer.from(bytes), image = sharp(buffer, { limitInputPixels: limits.pixels, failOn: 'warning' });
    const metadata = await image.metadata();
    if (!['png', 'jpeg'].includes(metadata.format ?? '') || !metadata.width || !metadata.height) throw new RenderError('UNSUPPORTED_IMAGE', 'Milestone 1 accepts PNG and JPEG attachments');
    await image.stats(); // Fully decode rather than trusting only the header.
    if (metadata.orientation && metadata.orientation !== 1) return { bytes: await image.rotate().png().toBuffer(), extension: 'png', contentType: 'image/png' };
    return { bytes: buffer, extension: metadata.format === 'jpeg' ? 'jpg' : 'png', contentType: metadata.format === 'jpeg' ? 'image/jpeg' : 'image/png' };
  } catch (error) {
    if (error instanceof RenderError) throw error;
    throw new RenderError('INVALID_IMAGE', 'Cannot decode image attachment');
  }
}
function validateRelationships(parts: PackageParts) {
  for (const [path, bytes] of parts) if (path.endsWith('.rels')) {
    const document = parseXml(bytes.toString('utf8'), path), ids = new Set<string>();
    const owner = path === '_rels/.rels' ? '' : posix.join(posix.dirname(posix.dirname(path)), posix.basename(path).slice(0, -5));
    for (const link of descendants(document, ns.packageRelationship, 'Relationship')) {
      const id = link.getAttribute('Id')!;
      if (!id || ids.has(id)) throw new RenderError('INVALID_PACKAGE', 'Duplicate or missing relationship ID', path); ids.add(id);
      if (link.getAttribute('TargetMode') === 'External') {
        if (!link.getAttribute('Type')?.endsWith('/hyperlink')) throw new RenderError('UNSUPPORTED_TEMPLATE', 'External assets are not supported', path);
      } else if (!parts.has(resolvePart(owner, link.getAttribute('Target')!))) throw new RenderError('INVALID_PACKAGE', 'Relationship references a missing package part', path);
    }
  }
}
export async function compileDocx(input: Uint8Array): Promise<CompiledTemplate> {
  const original = await readPackage(Buffer.from(input));
  validateRelationships(original);
  if ([...original.keys()].some(name => name.startsWith('_xmlsignatures/') || /vbaProject/i.test(name))) throw new RenderError('UNSUPPORTED_FORMAT', 'Signed or macro-enabled templates are not supported');
  const packageLinks = parseXml(original.get('_rels/.rels')!.toString('utf8'), '_rels/.rels');
  const mainLink = descendants(packageLinks, ns.packageRelationship, 'Relationship').find(e => e.getAttribute('Type') === `${ns.relationship}/officeDocument`);
  if (!mainLink) throw new RenderError('UNSUPPORTED_FORMAT', 'Missing document relationship');
  const mainPart = resolvePart('', mainLink.getAttribute('Target')!);
  const mainRoot = parseXml(original.get(mainPart)!.toString('utf8'), mainPart).documentElement!;
  if (mainRoot.namespaceURI !== ns.word || mainRoot.localName !== 'document') throw new RenderError('UNSUPPORTED_FORMAT', 'Only transitional WordprocessingML DOCX documents are supported', mainPart);
  const contentTypes = parseXml(original.get('[Content_Types].xml')!.toString('utf8'), '[Content_Types].xml');
  const mainType = descendants(contentTypes, ns.contentTypes, 'Override').find(e => e.getAttribute('PartName') === `/${mainPart}`)?.getAttribute('ContentType');
  if (mainType !== 'application/vnd.openxmlformats-officedocument.wordprocessingml.document.main+xml') throw new RenderError('UNSUPPORTED_FORMAT', 'Milestone 1 renders DOCX templates only');
  const storyParts = [mainPart];
  const mainRelationships = original.get(relsPath(mainPart));
  if (mainRelationships) for (const link of descendants(parseXml(mainRelationships.toString('utf8'), relsPath(mainPart)), ns.packageRelationship, 'Relationship')) {
    if (['header', 'footer'].some(kind => link.getAttribute('Type') === `${ns.relationship}/${kind}`)) storyParts.push(resolvePart(mainPart, link.getAttribute('Target')!));
  }
  const plans = [...new Set(storyParts)].map(part => compileStory(part, original.get(part)!.toString('utf8')));
  for (const plan of plans) if (plan.images.length) {
    const path = relsPath(plan.part), bytes = original.get(path);
    if (!bytes) throw new RenderError('INVALID_PACKAGE', 'Picture is missing its relationship part', plan.part);
    const links = descendants(parseXml(bytes.toString('utf8'), path), ns.packageRelationship, 'Relationship');
    for (const image of plan.images) if (!links.some(link => link.getAttribute('Id') === image.reference.value && link.getAttribute('Type') === `${ns.relationship}/image`)) throw new RenderError('INVALID_PACKAGE', 'Picture references an undefined image relationship', plan.part);
  }
  return Object.freeze({ format: 'docx' as const, async render(data: Record<string, unknown>, options: RenderOptions = {}) {
    validateData(data);
    const parts = new Map(original), attachmentMap = options.attachments ?? {};
    const images = new Map<string, Awaited<ReturnType<typeof normalizeImage>>>();
    // Resolve data and attachment bytes before modifying any part; no partially rendered result escapes.
    for (const plan of plans) for (const image of plan.images) {
      const key = Object.hasOwn(attachmentMap, image.path) ? image.path : scalar(lookup(data, image.path, plan.part), image.path, plan.part);
      if (!Object.hasOwn(attachmentMap, key) || !(attachmentMap[key] instanceof Uint8Array)) throw new RenderError('MISSING_ATTACHMENT', `Missing attachment: ${key}`, plan.part);
      if (!images.has(key)) images.set(key, await normalizeImage(attachmentMap[key]!));
    }
    const types = parseXml(original.get('[Content_Types].xml')!.toString('utf8'), '[Content_Types].xml');
    let typesChanged = false;
    for (const plan of plans) {
      const edits: Edit[] = [];
      for (const paragraph of plan.paragraphs) {
        const localEdits = new Map<TextSpan, Edit[]>();
        for (const binding of paragraph.bindings) {
          const value = scalar(lookup(data, binding.path, plan.part), binding.path, plan.part);
          for (const span of paragraph.spans) {
            const start = Math.max(binding.start, span.offset), end = Math.min(binding.end, span.offset + span.value.length);
            if (start >= end) continue;
            const local = localEdits.get(span) ?? [];
            local.push({ start: start - span.offset, end: end - span.offset, value: start === binding.start ? value : '' }); localEdits.set(span, local);
          }
        }
        for (const [span, replacements] of localEdits) {
          edits.push({ start: span.node.openEnd, end: span.node.closeStart, value: escapeXml(applyEdits(span.value, replacements)) });
          const space = span.node.attributes.find(a => a.uri === ns.xml && a.local === 'space');
          if (space) edits.push({ start: space.start, end: space.end, value: 'preserve' });
          else edits.push({ start: span.node.openEnd - 1, end: span.node.openEnd - 1, value: ' xml:space="preserve"' });
        }
      }
      if (plan.images.length) {
        const path = relsPath(plan.part), relationships = parseXml(original.get(path)!.toString('utf8'), path), root = relationships.documentElement!;
        const used = new Set(descendants(root, ns.packageRelationship, 'Relationship').map(e => e.getAttribute('Id')));
        let next = 1;
        for (const binding of plan.images) {
          const key = Object.hasOwn(attachmentMap, binding.path) ? binding.path : scalar(lookup(data, binding.path, plan.part), binding.path, plan.part);
          const image = images.get(key)!;
          const digest = createHash('sha256').update(image.bytes).digest('hex');
          const imagePart = posix.join(posix.dirname(mainPart), 'media', `imprint-${digest}.${image.extension}`);
          parts.set(imagePart, image.bytes);
          while (used.has(`imprintImage${next}`)) next++;
          const id = `imprintImage${next++}`; used.add(id);
          const link = relationships.createElementNS(ns.packageRelationship, root.prefix ? `${root.prefix}:Relationship` : 'Relationship');
          link.setAttribute('Id', id); link.setAttribute('Type', `${ns.relationship}/image`); link.setAttribute('Target', posix.relative(posix.dirname(plan.part), imagePart)); root.appendChild(link);
          edits.push({ start: binding.reference.start, end: binding.reference.end, value: id }, { start: binding.description.start, end: binding.description.end, value: escapeXml(`Attachment ${key}`) });
          const defaults = descendants(types, ns.contentTypes, 'Default');
          const existing = defaults.find(e => e.getAttribute('Extension') === image.extension);
          if (existing && existing.getAttribute('ContentType') !== image.contentType) throw new RenderError('INVALID_PACKAGE', 'Conflicting image content type');
          if (!existing) {
            const root = types.documentElement!, entry = types.createElementNS(ns.contentTypes, root.prefix ? `${root.prefix}:Default` : 'Default');
            entry.setAttribute('Extension', image.extension); entry.setAttribute('ContentType', image.contentType);
            root.insertBefore(entry, descendants(root, ns.contentTypes, 'Override')[0] ?? null); typesChanged = true;
          }
        }
        parts.set(path, serialize(relationships));
      }
      if (edits.length) {
        const result = applyEdits(plan.source, edits);
        if (Buffer.byteLength(result) > limits.entry) throw new RenderError('LIMIT_EXCEEDED', 'Rendered part exceeds 8 MiB', plan.part);
        parseXml(result, plan.part); parts.set(plan.part, Buffer.from(result));
      }
    }
    if (typesChanged) parts.set('[Content_Types].xml', serialize(types));
    if ([...parts.values()].reduce((sum, value) => sum + value.byteLength, 0) > limits.expanded) throw new RenderError('LIMIT_EXCEEDED', 'Rendered package exceeds 32 MiB');
    return writePackage(parts);
  } });
}
