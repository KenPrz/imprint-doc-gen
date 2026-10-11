import { createHash } from 'node:crypto';
import { posix } from 'node:path';
import sharp from 'sharp';
import { RenderError } from './errors.ts';
import { writePackage, relsPath, limits, type PackageParts } from './package.ts';
import { ns, parseXml, serialize, descendants } from './xml.ts';

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

export function createRenderPackage(original: PackageParts, mainPart: string, attachments: Record<string, Uint8Array>) {
  const parts = new Map(original), images = new Map<string, Awaited<ReturnType<typeof normalizeImage>>>();
  const types = parseXml(original.get('[Content_Types].xml')!.toString('utf8'), '[Content_Types].xml');
  const relationships = new Map<string, { document: ReturnType<typeof parseXml>; used: Set<string | null>; next: number }>();
  let typesChanged = false, packageBytes = [...parts.values()].reduce((sum, value) => sum + value.byteLength, 0);
  const setPart = (path: string, bytes: Buffer) => {
    if (!parts.has(path) && parts.size >= limits.entries) throw new RenderError('LIMIT_EXCEEDED', 'Rendered package has too many entries', path);
    packageBytes += bytes.byteLength - (parts.get(path)?.byteLength ?? 0);
    if (bytes.byteLength > limits.entry || packageBytes > limits.expanded) throw new RenderError('LIMIT_EXCEEDED', 'Rendered package exceeds its byte budget', path);
    parts.set(path, bytes);
  };
  return {
    setPart,
    async addImage(owner: string, key: string): Promise<string> {
      if (!Object.hasOwn(attachments, key) || !(attachments[key] instanceof Uint8Array)) throw new RenderError('MISSING_ATTACHMENT', `Missing attachment: ${key}`, owner);
      if (!images.has(key)) images.set(key, await normalizeImage(attachments[key]!));
      const image = images.get(key)!;
      const digest = createHash('sha256').update(image.bytes).digest('hex');
      const imagePart = posix.join(posix.dirname(mainPart), 'media', `imprint-${digest}.${image.extension}`);
      const existingImage = parts.get(imagePart);
      if (existingImage && !existingImage.equals(image.bytes)) throw new RenderError('INVALID_PACKAGE', 'Image part name conflicts with existing content', imagePart);
      if (!existingImage) setPart(imagePart, image.bytes);
      const path = relsPath(owner);
      if (!relationships.has(path)) {
        const bytes = original.get(path);
        if (!bytes) throw new RenderError('INVALID_PACKAGE', 'Picture is missing its relationship part', owner);
        const document = parseXml(bytes.toString('utf8'), path);
        relationships.set(path, { document, used: new Set(descendants(document, ns.packageRelationship, 'Relationship').map(e => e.getAttribute('Id'))), next: 1 });
      }
      const state = relationships.get(path)!;
      while (state.used.has(`imprintImage${state.next}`)) state.next++;
      const id = `imprintImage${state.next++}`; state.used.add(id);
      const root = state.document.documentElement!;
      const link = state.document.createElementNS(ns.packageRelationship, root.prefix ? `${root.prefix}:Relationship` : 'Relationship');
      link.setAttribute('Id', id); link.setAttribute('Type', `${ns.relationship}/image`); link.setAttribute('Target', posix.relative(posix.dirname(owner), imagePart)); root.appendChild(link);
      const existing = descendants(types, ns.contentTypes, 'Default').find(e => e.getAttribute('Extension') === image.extension);
      if (existing && existing.getAttribute('ContentType') !== image.contentType) throw new RenderError('INVALID_PACKAGE', 'Conflicting image content type');
      if (!existing) {
        const root = types.documentElement!, entry = types.createElementNS(ns.contentTypes, root.prefix ? `${root.prefix}:Default` : 'Default');
        entry.setAttribute('Extension', image.extension); entry.setAttribute('ContentType', image.contentType);
        root.insertBefore(entry, descendants(root, ns.contentTypes, 'Override')[0] ?? null); typesChanged = true;
      }
      return id;
    },
    async finish() {
      for (const [path, state] of relationships) setPart(path, serialize(state.document));
      if (typesChanged) setPart('[Content_Types].xml', serialize(types));
      return writePackage(parts);
    },
  };
}
