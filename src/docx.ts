import { createHash } from 'node:crypto';
import { posix } from 'node:path';
import sharp from 'sharp';
import { RenderError } from './errors.ts';
import { compileStory, renderStory, identityAttributes, type RenderBudget } from './docx-story.ts';
import { forbidden, lookup, scalar } from './scope.ts';
import { readPackage, writePackage, resolvePart, relsPath, limits, type PackageParts } from './package.ts';
import { ns, indexXml, escapeXml, parseXml, serialize, descendants, type Edit } from './xml.ts';

export interface RenderOptions { attachments?: Record<string, Uint8Array> }
export interface CompiledTemplate { readonly format: 'docx'; render(data: Record<string, unknown>, options?: RenderOptions): Promise<Buffer> }
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
    if (!['png', 'jpeg'].includes(metadata.format ?? '') || !metadata.width || !metadata.height) throw new RenderError('UNSUPPORTED_IMAGE', 'Image attachments must be PNG or JPEG');
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
  if (mainType !== 'application/vnd.openxmlformats-officedocument.wordprocessingml.document.main+xml') throw new RenderError('UNSUPPORTED_FORMAT', 'Only DOCX templates are supported');
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
  // Reserve identities in every XML part, including untouched notes or drawings.
  const reservedIdentities = new Set<number>();
  for (const [part, bytes] of original) if (part.endsWith('.xml')) {
    for (const { attribute, hex } of indexXml(bytes.toString('utf8'), part).flatMap(identityAttributes)) {
      const value = Number.parseInt(attribute.value, hex ? 16 : 10);
      if (Number.isFinite(value)) reservedIdentities.add(value);
    }
  }
  return Object.freeze({ format: 'docx' as const, async render(data: Record<string, unknown>, options: RenderOptions = {}) {
    validateData(data);
    const parts = new Map(original), attachmentMap = options.attachments ?? {};
    const images = new Map<string, Awaited<ReturnType<typeof normalizeImage>>>();
    const usedIdentities = new Set(reservedIdentities); let nextIdentity = 1;
    const freshIdentity = (hex: boolean) => {
      while (usedIdentities.has(nextIdentity)) nextIdentity++;
      const id = nextIdentity++; usedIdentities.add(id);
      return hex ? id.toString(16).toUpperCase().padStart(8, '0') : String(id);
    };
    const budget: RenderBudget = { iterations: 0 };
    const types = parseXml(original.get('[Content_Types].xml')!.toString('utf8'), '[Content_Types].xml');
    let typesChanged = false;
    let packageBytes = [...parts.values()].reduce((sum, value) => sum + value.byteLength, 0);
    const setPart = (path: string, bytes: Buffer) => {
      if (!parts.has(path) && parts.size >= limits.entries) throw new RenderError('LIMIT_EXCEEDED', 'Rendered package has too many entries', path);
      packageBytes += bytes.byteLength - (parts.get(path)?.byteLength ?? 0);
      if (bytes.byteLength > limits.entry || packageBytes > limits.expanded) throw new RenderError('LIMIT_EXCEEDED', 'Rendered package exceeds its byte budget', path);
      parts.set(path, bytes);
    };
    for (const plan of plans) {
      const path = relsPath(plan.part);
      const relationships = plan.images.length ? parseXml(original.get(path)!.toString('utf8'), path) : undefined;
      const used = new Set(relationships ? descendants(relationships, ns.packageRelationship, 'Relationship').map(e => e.getAttribute('Id')) : []);
      let next = 1, relationshipsChanged = false;
      const result = await renderStory(plan, data, budget, async (binding, scope): Promise<Edit[]> => {
        // Literal attachment keys retain milestone-1 behavior at root scope.
        // Inside loops, image paths resolve from the current item or explicit ../ parent.
        const key = scope.length === 1 && Object.hasOwn(attachmentMap, binding.path) ? binding.path : scalar(lookup(scope, binding.path, plan.part), binding.path, plan.part);
        if (!Object.hasOwn(attachmentMap, key) || !(attachmentMap[key] instanceof Uint8Array)) throw new RenderError('MISSING_ATTACHMENT', `Missing attachment: ${key}`, plan.part);
        if (!images.has(key)) images.set(key, await normalizeImage(attachmentMap[key]!));
        const image = images.get(key)!;
        const digest = createHash('sha256').update(image.bytes).digest('hex');
        const imagePart = posix.join(posix.dirname(mainPart), 'media', `imprint-${digest}.${image.extension}`);
        const existingImage = parts.get(imagePart);
        if (existingImage && !existingImage.equals(image.bytes)) throw new RenderError('INVALID_PACKAGE', 'Image part name conflicts with existing content', imagePart);
        if (!existingImage) setPart(imagePart, image.bytes);
        while (used.has(`imprintImage${next}`)) next++;
        const id = `imprintImage${next++}`; used.add(id);
        const root = relationships!.documentElement!;
        const link = relationships!.createElementNS(ns.packageRelationship, root.prefix ? `${root.prefix}:Relationship` : 'Relationship');
        link.setAttribute('Id', id); link.setAttribute('Type', `${ns.relationship}/image`); link.setAttribute('Target', posix.relative(posix.dirname(plan.part), imagePart)); root.appendChild(link);
        relationshipsChanged = true;
        const existing = descendants(types, ns.contentTypes, 'Default').find(e => e.getAttribute('Extension') === image.extension);
        if (existing && existing.getAttribute('ContentType') !== image.contentType) throw new RenderError('INVALID_PACKAGE', 'Conflicting image content type');
        if (!existing) {
          const root = types.documentElement!, entry = types.createElementNS(ns.contentTypes, root.prefix ? `${root.prefix}:Default` : 'Default');
          entry.setAttribute('Extension', image.extension); entry.setAttribute('ContentType', image.contentType);
          root.insertBefore(entry, descendants(root, ns.contentTypes, 'Override')[0] ?? null); typesChanged = true;
        }
        return [{ start: binding.reference.start, end: binding.reference.end, value: id }, { start: binding.description.start, end: binding.description.end, value: escapeXml(`Attachment ${key}`) }];
      }, freshIdentity);
      if (relationshipsChanged) setPart(path, serialize(relationships!));
      if (result !== plan.source) { parseXml(result, plan.part); setPart(plan.part, Buffer.from(result)); }
    }
    if (typesChanged) setPart('[Content_Types].xml', serialize(types));
    return writePackage(parts);
  } });
}
