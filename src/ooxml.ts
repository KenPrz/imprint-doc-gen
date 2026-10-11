import { formatNs } from './format-namespaces.ts';
import { mediaTypes } from './types.ts';
import { posix } from 'node:path';
import { RenderError } from './errors.ts';
import { readPackage, resolvePart, type PackageParts } from './package.ts';
import { ns, parseXml, descendants } from './xml.ts';
import type { DocumentFormat } from './types.ts';

export function validateRelationships(parts: PackageParts) {
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

export interface OfficePackage { parts: PackageParts; mainPart: string; format: DocumentFormat }
export async function openOfficePackage(input: Uint8Array): Promise<OfficePackage> {
  const parts = await readPackage(Buffer.from(input));
  if (parts.has('mimetype')) {
    const mime = parts.get('mimetype')!.toString(), format = (['odt','ods','odp'] as const).find(f=>mediaTypes[f]===mime);
    if(!format || !parts.has('content.xml')) throw new RenderError('UNSUPPORTED_FORMAT','Unsupported OpenDocument package');
    if(!parts.has('META-INF/manifest.xml'))throw new RenderError('INVALID_PACKAGE','OpenDocument manifest is missing');
    const content=parseXml(parts.get('content.xml')!.toString(),'content.xml');
    if(content.documentElement?.namespaceURI!==formatNs.office || content.documentElement.localName!=='document-content' || !descendants(content,formatNs.office,{odt:'text',ods:'spreadsheet',odp:'presentation'}[format]).length)throw new RenderError('INVALID_PACKAGE','OpenDocument content does not match its media type');
    const manifest=parseXml(parts.get('META-INF/manifest.xml')!.toString(),'META-INF/manifest.xml');
    if(manifest.documentElement?.namespaceURI!==formatNs.manifest || manifest.documentElement.localName!=='manifest' || !descendants(manifest,formatNs.manifest,'file-entry').some(entry=>entry.getAttributeNS(formatNs.manifest,'full-path')==='/'&&entry.getAttributeNS(formatNs.manifest,'media-type')===mime))throw new RenderError('INVALID_PACKAGE','OpenDocument manifest media type is missing or inconsistent');
    if(descendants(manifest,formatNs.manifest,'encryption-data').length || [...parts.keys()].some(p=>/signatures|^Basic\/|^Scripts\//i.test(p))) throw new RenderError('UNSUPPORTED_FORMAT','Encrypted, signed or scripted OpenDocument packages are unsupported');
    for(const entry of descendants(manifest,formatNs.manifest,'file-entry')) {
      const path=entry.getAttributeNS(formatNs.manifest,'full-path');
      if(path && path!=='/' && !path.endsWith('/') && !parts.has(path)) throw new RenderError('INVALID_PACKAGE','OpenDocument manifest references a missing part');
    }
    return { parts,mainPart:'content.xml',format };
  }
  validateRelationships(parts);
  if ([...parts.keys()].some(name => name.startsWith('_xmlsignatures/') || /vbaProject/i.test(name))) throw new RenderError('UNSUPPORTED_FORMAT', 'Signed or macro-enabled templates are not supported');
  const links = descendants(parseXml(parts.get('_rels/.rels')!.toString('utf8'), '_rels/.rels'), ns.packageRelationship, 'Relationship').filter(e => e.getAttribute('Type') === `${ns.relationship}/officeDocument`);
  if (links.length !== 1) throw new RenderError('UNSUPPORTED_FORMAT', 'Expected one Office document relationship');
  const mainPart = resolvePart('', links[0]!.getAttribute('Target')!);
  const root = parseXml(parts.get(mainPart)!.toString('utf8'), mainPart).documentElement!;
  const types = parseXml(parts.get('[Content_Types].xml')!.toString('utf8'), '[Content_Types].xml');
  const type = descendants(types, ns.contentTypes, 'Override').find(e => e.getAttribute('PartName') === `/${mainPart}`)?.getAttribute('ContentType');
  if (root.namespaceURI === ns.word && root.localName === 'document' && type === 'application/vnd.openxmlformats-officedocument.wordprocessingml.document.main+xml') return { parts, mainPart, format: 'docx' };
  if (root.namespaceURI === ns.spreadsheet && root.localName === 'workbook' && type === 'application/vnd.openxmlformats-officedocument.spreadsheetml.sheet.main+xml') return { parts, mainPart, format: 'xlsx' };
  if (root.namespaceURI === formatNs.presentation && root.localName === 'presentation' && type === 'application/vnd.openxmlformats-officedocument.presentationml.presentation.main+xml') return { parts,mainPart,format:'pptx' };
  throw new RenderError('UNSUPPORTED_FORMAT', 'Expected a supported transitional DOCX, XLSX or PPTX package', mainPart);
}
