import { createHash } from 'node:crypto';
import { deflateSync } from 'node:zlib';
import { posix } from 'node:path';
import { ZipReader, ZipWriter, Uint8ArrayReader, Uint8ArrayWriter } from '@zip.js/zip.js';
import { DOMParser } from '@xmldom/xmldom';
import type { Document, Element } from '@xmldom/xmldom';

export const NS = {
  w: 'http://schemas.openxmlformats.org/wordprocessingml/2006/main',
  s: 'http://schemas.openxmlformats.org/spreadsheetml/2006/main',
  r: 'http://schemas.openxmlformats.org/officeDocument/2006/relationships',
  rel: 'http://schemas.openxmlformats.org/package/2006/relationships',
  ct: 'http://schemas.openxmlformats.org/package/2006/content-types',
  a: 'http://schemas.openxmlformats.org/drawingml/2006/main',
  wp: 'http://schemas.openxmlformats.org/drawingml/2006/wordprocessingDrawing',
  pic: 'http://schemas.openxmlformats.org/drawingml/2006/picture',
  xdr: 'http://schemas.openxmlformats.org/drawingml/2006/spreadsheetDrawing',
};
export type Format = 'docx' | 'xlsx';
export type Parts = Map<string, Buffer>;
export const xml = (s: string) => `<?xml version="1.0" encoding="UTF-8" standalone="yes"?>${s}`;
export const esc = (s: unknown) => String(s).replace(/&/g, '&amp;').replace(/</g, '&lt;').replace(/>/g, '&gt;').replace(/"/g, '&quot;').replace(/'/g, '&apos;');
export const hash = (b: Uint8Array | string) => createHash('sha256').update(b).digest('hex');
export const put = (p: Parts, name: string, text: string | Buffer) => p.set(name, Buffer.isBuffer(text) ? text : Buffer.from(xml(text)));
export const rels = (items: { id: string; type: string; target: string; external?: boolean }[]) =>
  `<Relationships xmlns="${NS.rel}">${items.map(r => `<Relationship Id="${r.id}" Type="${NS.r}/${r.type}" Target="${esc(r.target)}"${r.external ? ' TargetMode="External"' : ''}/>`).join('')}</Relationships>`;
export function contentTypes(overrides: Record<string, string>) {
  return `<Types xmlns="${NS.ct}"><Default Extension="rels" ContentType="application/vnd.openxmlformats-package.relationships+xml"/><Default Extension="xml" ContentType="application/xml"/><Default Extension="png" ContentType="image/png"/>${Object.entries(overrides).map(([part, type]) => `<Override PartName="/${part}" ContentType="application/vnd.openxmlformats-officedocument.${type}+xml"/>`).join('')}</Types>`;
}
export function parse(text: string): Document {
  if (/<!DOCTYPE|<!ENTITY/i.test(text)) throw new Error('DTD/entity declarations are outside this corpus');
  return new DOMParser({ onError: (severity, message) => { throw new Error(`${severity}: ${message}`); } }).parseFromString(text, 'application/xml');
}
export const elements = (node: Document | Element, ns: string, local: string): Element[] => Array.from(node.getElementsByTagNameNS(ns, local));
export const str = (p: Parts, name: string) => { const b = p.get(name); if (!b) throw new Error(`Missing part ${name}`); return b.toString('utf8'); };
export function relationshipPath(part: string) { return posix.join(posix.dirname(part), '_rels', posix.basename(part) + '.rels'); }
export function targets(p: Parts, part: string): Map<string, string> {
  const path = relationshipPath(part);
  if (!p.has(path)) return new Map();
  return new Map(elements(parse(str(p, path)), NS.rel, 'Relationship').map(r => [r.getAttribute('Id')!, r.getAttribute('TargetMode') === 'External' ? r.getAttribute('Target')! : posix.normalize(posix.join(posix.dirname(part), r.getAttribute('Target')!))]));
}
export async function pack(parts: Parts, disorder = 0): Promise<Buffer> {
  const writer = new ZipWriter(new Uint8ArrayWriter(), { useWebWorkers: false });
  const entries = [...parts].sort(([a], [b]) => disorder ? hash(`${disorder}:${a}`).localeCompare(hash(`${disorder}:${b}`)) : a.localeCompare(b));
  for (const [name, data] of entries) await writer.add(name, new Uint8ArrayReader(data), {
    level: disorder % 2 ? 0 : 6, lastModDate: new Date(2020, 0, 1, 0, 0, 0), extendedTimestamp: false, dataDescriptor: false,
  });
  return Buffer.from(await writer.close());
}
export async function unpack(bytes: Uint8Array): Promise<Parts> {
  const reader = new ZipReader(new Uint8ArrayReader(bytes), { useWebWorkers: false });
  const out: Parts = new Map();
  try {
    const entries = await reader.getEntries();
    if (entries.length > 2048) throw new Error('Entry limit exceeded');
    let total = 0;
    for (const entry of entries) {
      if (entry.directory) continue;
      if (entry.filename.startsWith('/') || entry.filename.split('/').includes('..') || out.has(entry.filename)) throw new Error('Unsafe or duplicate ZIP path');
      total += entry.uncompressedSize;
      if (total > 64 * 1024 * 1024) throw new Error('Expanded byte limit exceeded');
      out.set(entry.filename, Buffer.from(await entry.getData!(new Uint8ArrayWriter(), { checkSignature: true })));
    }
    return out;
  } finally { await reader.close(); }
}
// Only creates small deterministic test attachments; not an application image encoder.
function crc32(b: Buffer) {
  let crc = 0xffffffff;
  for (const byte of b) { crc ^= byte; for (let i = 0; i < 8; i++) crc = (crc >>> 1) ^ ((crc & 1) ? 0xedb88320 : 0); }
  return (crc ^ 0xffffffff) >>> 0;
}
export function png(color: [number, number, number], width = 24, height = 16): Buffer {
  const chunk = (name: string, data: Buffer) => { const b = Buffer.alloc(data.length + 12); b.writeUInt32BE(data.length); b.write(name, 4); data.copy(b, 8); b.writeUInt32BE(crc32(b.subarray(4, -4)), b.length - 4); return b; };
  const header = Buffer.alloc(13); header.writeUInt32BE(width); header.writeUInt32BE(height, 4); header[8] = 8; header[9] = 2;
  const pixels = Buffer.alloc(height * (1 + width * 3));
  for (let y = 0; y < height; y++) for (let x = 0; x < width; x++) for (let c = 0; c < 3; c++) pixels[y * (1 + width * 3) + 1 + x * 3 + c] = color[c]!;
  return Buffer.concat([Buffer.from('89504e470d0a1a0a', 'hex'), chunk('IHDR', header), chunk('IDAT', deflateSync(pixels)), chunk('IEND', Buffer.alloc(0))]);
}
export function perturb(parts: Parts, severity: number): Parts {
  if (!severity) return parts;
  const out: Parts = new Map();
  for (const [name, bytes] of parts) {
    if (!name.endsWith('.xml') && !name.endsWith('.rels')) { out.set(name, bytes); continue; }
    let text = bytes.toString('utf8');
    // Namespace aliases and rIds are semantically irrelevant; replace consistently.
    text = text.replace(/\bw:/g, 'word:').replace(/xmlns:w=/g, 'xmlns:word=');
    text = text.replace(/\brId(\d+)/g, (_, n) => `rel_${Number(n) * 37 + severity}`);
    if (severity >= 3) text = text.replace(/></g, `><!-- fixture padding ${severity} --><`);
    // LibreOffice's DOCX detector requires conventional package metadata prefixes.
    // Keep alias variation in document/worksheet content, not content-types or rels.
    if (severity >= 5 && name !== '[Content_Types].xml' && !name.endsWith('.rels')) text = text.replace(/xmlns="([^"]+)"/, 'xmlns:pkg="$1"');
    // Prefix all formerly default-namespace elements, avoiding prefixed elements.
    if (severity >= 5 && text.includes('xmlns:pkg=')) text = text.replace(/<(\/?)([A-Za-z_][\w.-]*)(?=[\s/>])/g, '<$1pkg:$2');
    out.set(name, Buffer.from(text));
  }
  return out;
}
