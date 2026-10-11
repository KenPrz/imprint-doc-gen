import { posix } from 'node:path';
import { ZipReader, ZipWriter, Uint8ArrayReader, Uint8ArrayWriter } from '@zip.js/zip.js';
import { RenderError } from './errors.ts';

export type PackageParts = Map<string, Buffer>;
export const limits = Object.freeze({ archive: 8 * 1024 * 1024, entry: 8 * 1024 * 1024, expanded: 32 * 1024 * 1024, entries: 1024, image: 8 * 1024 * 1024, pixels: 16_000_000 });
export function resolvePart(owner: string, target: string): string {
  if (!target || target.includes('\\') || target.includes('\0') || /[?#]/.test(target)) throw new RenderError('PACKAGE_PATH', 'Unsupported relationship target', owner);
  const decoded = decodeURIComponent(target);
  const resolved = posix.normalize(decoded.startsWith('/') ? decoded.slice(1) : posix.join(posix.dirname(owner), decoded));
  if (resolved.startsWith('../') || resolved === '..' || resolved.startsWith('/')) throw new RenderError('PACKAGE_PATH', 'Relationship escapes package', owner);
  return resolved;
}
export const relsPath = (part: string) => posix.join(posix.dirname(part), '_rels', posix.basename(part) + '.rels');
export async function readPackage(input: Uint8Array): Promise<PackageParts> {
  if (input.byteLength > limits.archive) throw new RenderError('LIMIT_EXCEEDED', 'Template exceeds 8 MiB');
  const zip = new ZipReader(new Uint8ArrayReader(input), { useWebWorkers: false });
  const parts: PackageParts = new Map();
  let total = 0;
  try {
    const entries = await zip.getEntries();
    if (entries.length > limits.entries) throw new RenderError('LIMIT_EXCEEDED', 'Too many package entries');
    for (const entry of entries) {
      const name = entry.filename;
      if (name.startsWith('/') || name.includes('\\') || name.includes('\0') || name.split('/').includes('..') || posix.normalize(name) !== name || parts.has(name)) throw new RenderError('PACKAGE_PATH', 'Unsafe or duplicate package path');
      if (entry.directory) continue;
      if (entry.encrypted) throw new RenderError('UNSUPPORTED_FORMAT', 'Encrypted templates are not supported');
      if (entry.uncompressedSize > limits.entry) throw new RenderError('LIMIT_EXCEEDED', 'Package entry exceeds 8 MiB', name);
      const chunks: Buffer[] = []; let size = 0;
      const sink = new WritableStream<Uint8Array>({ write(chunk) {
        size += chunk.byteLength; total += chunk.byteLength;
        if (size > limits.entry || total > limits.expanded) throw new RenderError('LIMIT_EXCEEDED', 'Expanded package exceeds its byte budget', name);
        chunks.push(Buffer.from(chunk));
      } });
      await entry.getData!(sink, { checkSignature: true, signal: AbortSignal.timeout(15_000) });
      parts.set(name, Buffer.concat(chunks));
    }
    if ((!parts.has('[Content_Types].xml') || !parts.has('_rels/.rels')) && (!parts.has('mimetype') || !parts.has('META-INF/manifest.xml'))) throw new RenderError('UNSUPPORTED_FORMAT', 'Expected an Office Open XML package');
    return parts;
  } catch (error) {
    if (error instanceof RenderError) throw error;
    throw new RenderError('INVALID_PACKAGE', 'Cannot read template ZIP package');
  } finally { await zip.close(); }
}
export async function writePackage(parts: PackageParts): Promise<Buffer> {
  const zip = new ZipWriter(new Uint8ArrayWriter(), { useWebWorkers: false });
  for (const [name, bytes] of [...parts].sort(([a], [b]) => a === 'mimetype' ? -1 : b === 'mimetype' ? 1 : a.localeCompare(b))) await zip.add(name, new Uint8ArrayReader(bytes), { level: name === 'mimetype' ? 0 : 6, extendedTimestamp: false, dataDescriptor: false, lastModDate: new Date(2020, 0, 1) });
  return Buffer.from(await zip.close());
}
