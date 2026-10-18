import { validateSvg } from './svg.ts';
import { createHash } from 'node:crypto';
import { posix } from 'node:path';
import sharp from 'sharp';
import type { ImageOptions } from './types.ts';
import { RenderError } from './errors.ts';
import { writePackage, relsPath, limits, type PackageParts } from './package.ts';
import { ns, parseXml, serialize, descendants } from './xml.ts';

export async function normalizeImage(bytes: Uint8Array, options: ImageOptions = {}, size?: { width: number; height: number }): Promise<{ bytes: Buffer; extension: string; contentType: string }> {
  if (!bytes.byteLength || bytes.byteLength > limits.image) throw new RenderError('LIMIT_EXCEEDED', 'Attachment must be between 1 byte and 8 MiB');
  try {
    const buffer = Buffer.from(bytes);
    if (/^\s*(?:<\?xml[^>]*>\s*)?<svg[\s>]/i.test(buffer.toString('utf8'))) validateSvg(buffer);
    const image = sharp(buffer, { limitInputPixels: limits.pixels, failOn: 'warning' });
    const metadata = await image.metadata();
    if (!['png', 'jpeg', 'svg'].includes(metadata.format ?? '') || !metadata.width || !metadata.height) throw new RenderError('UNSUPPORTED_IMAGE', 'Image attachments must be PNG, JPEG or passive SVG');
    if (Object.keys(options).some(key => !['fit','width','height','rotate','background','crop'].includes(key))) throw new RenderError('INVALID_IMAGE_OPTIONS', 'Unknown image option');
    if (options.fit !== undefined && !['contain','cover','fill'].includes(options.fit)) throw new RenderError('INVALID_IMAGE_OPTIONS', 'Expected contain, cover or fill');
    if (options.rotate !== undefined && (!Number.isFinite(options.rotate) || Math.abs(options.rotate) > 360)) throw new RenderError('INVALID_IMAGE_OPTIONS', 'Rotation must be between -360 and 360');
    if (options.background !== undefined && !/^#[0-9a-f]{6}(?:[0-9a-f]{2})?$/i.test(options.background)) throw new RenderError('INVALID_IMAGE_OPTIONS', 'Background must be a hex color');
    for (const value of [options.width,options.height]) if(value!==undefined && (!Number.isSafeInteger(value)||value<1||value>4096))throw new RenderError('INVALID_IMAGE_OPTIONS','Image dimensions must be between 1 and 4096 pixels');
    const width = options.width ?? size?.width, height = options.height ?? size?.height;
    if (Object.keys(options).length && (options.fit || options.width || options.height)) for (const value of [width, height]) if (!Number.isSafeInteger(value) || value! < 1 || value! > 4096) throw new RenderError('INVALID_IMAGE_OPTIONS', 'A fitting canvas must have dimensions between 1 and 4096 pixels');
    if (options.crop && (Object.keys(options.crop).sort().join(',') !== 'height,left,top,width' || Object.values(options.crop).some(n => !Number.isSafeInteger(n) || n < 0) || !options.crop.width || !options.crop.height)) throw new RenderError('INVALID_IMAGE_OPTIONS', 'Invalid crop rectangle');
    await image.stats(); // Fully decode rather than trusting only the header.
    if (metadata.format === 'svg') validateSvg(buffer);
    if (Object.keys(options).length || metadata.format === 'svg') {
      let pipeline = image.autoOrient();
      if (options.crop) pipeline = pipeline.extract(options.crop);
      if (options.rotate !== undefined) pipeline = pipeline.rotate(options.rotate, { background: options.background ?? '#00000000' });
      if (options.fit || options.width || options.height) pipeline = pipeline.resize(width, height, { fit: options.fit ?? 'contain', background: options.background ?? '#00000000' });
      const result = await pipeline.png().toBuffer();
      if (result.length > limits.image) throw new RenderError('LIMIT_EXCEEDED', 'Transformed image exceeds 8 MiB');
      return { bytes: result, extension: 'png', contentType: 'image/png' };
    }
    if (metadata.orientation && metadata.orientation !== 1) { const rotated=await image.rotate().png().toBuffer();if(rotated.length>limits.image)throw new RenderError('LIMIT_EXCEEDED','Oriented image exceeds 8 MiB');return { bytes: rotated, extension: 'png', contentType: 'image/png' }; }
    return { bytes: buffer, extension: metadata.format === 'jpeg' ? 'jpg' : 'png', contentType: metadata.format === 'jpeg' ? 'image/jpeg' : 'image/png' };
  } catch (error) {
    if (error instanceof RenderError) throw error;
    throw new RenderError('INVALID_IMAGE', 'Cannot decode image attachment');
  }
}

export function createRenderPackage(original: PackageParts, mainPart: string, attachments: Record<string, Uint8Array>, imageOptions: Record<string, ImageOptions> = {}) {
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
    async addImage(owner: string, key: string, size?: { width: number; height: number }): Promise<string> {
      if (!Object.hasOwn(attachments, key) || !(attachments[key] instanceof Uint8Array)) throw new RenderError('MISSING_ATTACHMENT', `Missing attachment: ${key}`, owner);
      const options = imageOptions[key] ?? {}, cacheKey = JSON.stringify([key, options, size]);
      if (!images.has(cacheKey)) images.set(cacheKey, await normalizeImage(attachments[key]!, options, size));
      const image = images.get(cacheKey)!;
      const digest = createHash('sha256').update(image.bytes).digest('hex');
      const imagePart = posix.join(posix.dirname(mainPart), 'media', `imprint-${digest}.${image.extension}`);
      const existingImage = parts.get(imagePart);
      if (existingImage && !existingImage.equals(image.bytes)) throw new RenderError('INVALID_PACKAGE', 'Image part name conflicts with existing content', imagePart);
      if (!existingImage) setPart(imagePart, image.bytes);
      const path = relsPath(owner);
      if (!relationships.has(path)) {
        const bytes = original.get(path);
        const document = parseXml(bytes?.toString('utf8') ?? `<Relationships xmlns="${ns.packageRelationship}"/>`, path);
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
