import { parseXml, descendants } from './xml.ts';
import { RenderError } from './errors.ts';
// Deliberately passive geometry only: no CSS, scripts, fonts, references or URL loaders.
export function validateSvg(bytes: Buffer): void {
  const source = new TextDecoder('utf-8', { fatal: true }).decode(bytes);
  if (source.replace(/^\s*<\?xml\s[^?]*\?>/, '').includes('<?'))
    throw new RenderError('UNSUPPORTED_IMAGE', 'SVG processing instructions are unsupported');
  const doc = parseXml(source, 'attachment.svg');
  const allowed = new Set([
    'svg',
    'g',
    'path',
    'rect',
    'circle',
    'ellipse',
    'line',
    'polyline',
    'polygon',
    'title',
    'desc',
  ]);
  const attrs = new Set([
    'xmlns',
    'width',
    'height',
    'viewBox',
    'x',
    'y',
    'x1',
    'y1',
    'x2',
    'y2',
    'cx',
    'cy',
    'r',
    'rx',
    'ry',
    'd',
    'points',
    'fill',
    'fill-rule',
    'fill-opacity',
    'stroke',
    'stroke-width',
    'stroke-opacity',
    'stroke-linecap',
    'stroke-linejoin',
    'stroke-miterlimit',
    'stroke-dasharray',
    'stroke-dashoffset',
    'opacity',
    'transform',
    'preserveAspectRatio',
  ]);
  if (doc.documentElement?.localName !== 'svg')
    throw new RenderError('UNSUPPORTED_IMAGE', 'Expected SVG root');
  for (const node of [doc.documentElement!, ...descendants(doc.documentElement!, '*', '*')]) {
    if (node.namespaceURI !== 'http://www.w3.org/2000/svg' || !allowed.has(node.localName!))
      throw new RenderError('UNSUPPORTED_IMAGE', 'SVG supports passive geometry only');
    for (const a of Array.from(node.attributes))
      if (!attrs.has(a.name) || /url|[&<>]|\\/i.test(a.value))
        throw new RenderError('UNSUPPORTED_IMAGE', 'Unsupported SVG attribute or reference');
  }
}
