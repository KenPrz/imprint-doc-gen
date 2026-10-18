import type { Element } from '@xmldom/xmldom';
import { readPackage, writePackage, limits } from './package.ts';
import { parseXml, descendants, serialize } from './xml.ts';
import { RenderError } from './errors.ts';
export interface ChartData {
  categories: string[];
  series: { name: string; values: number[] }[];
}
const c = 'http://schemas.openxmlformats.org/drawingml/2006/chart';
export async function updateCharts(bytes: Buffer, updates: Record<string, ChartData>): Promise<Buffer> {
  const parts = await readPackage(bytes);
  if (Object.keys(updates).length > 32) throw new RenderError('LIMIT_EXCEEDED', 'At most 32 chart updates');
  for (const [part, data] of Object.entries(updates)) {
    if (!/^(word|xl|ppt)\/charts\/chart[\w-]*\.xml$/.test(part) || !parts.has(part))
      throw new RenderError('INVALID_DATA', 'Chart part does not exist');
    if (
      !data ||
      !Array.isArray(data.categories) ||
      !Array.isArray(data.series) ||
      !data.series.length ||
      data.series.length > 100 ||
      data.categories.length > 10000 ||
      data.categories.some((v) => typeof v !== 'string' || v.length > 32767) ||
      data.series.some(
        (s) =>
          !s ||
          typeof s.name !== 'string' ||
          s.name.length > 32767 ||
          !Array.isArray(s.values) ||
          s.values.length !== data.categories.length ||
          s.values.some((v) => typeof v !== 'number' || !Number.isFinite(v)),
      )
    )
      throw new RenderError('INVALID_DATA', 'Invalid category chart data');
    const estimate = data.series.reduce(
      (n, s) =>
        n +
        Buffer.byteLength(s.name) +
        data.categories.reduce((m, v) => m + Buffer.byteLength(v) + 100, 0) +
        s.values.length * 80,
      0,
    );
    if (estimate > limits.entry)
      throw new RenderError('LIMIT_EXCEEDED', 'Chart point data exceeds XML budget');
    const doc = parseXml(parts.get(part)!.toString(), part);
    if (
      descendants(doc, c, 'externalData').length ||
      descendants(doc, c, 'pivotSource').length ||
      descendants(doc, c, 'f').length
    )
      throw new RenderError(
        'UNSUPPORTED_TEMPLATE',
        'Chart updates require literal data, without workbook or pivot links',
      );
    const plots = descendants(doc, c, 'plotArea');
    const chart =
      plots[0] &&
      Array.from(plots[0].childNodes).filter(
        (n) =>
          n.nodeType === 1 && ['barChart', 'lineChart', 'pieChart', 'areaChart'].includes(n.localName ?? ''),
      );
    if (!chart || chart.length !== 1 || descendants(doc, c, 'ser').length === 0)
      throw new RenderError(
        'UNSUPPORTED_TEMPLATE',
        'Expected a single category chart with a prototype series',
      );
    const owner = chart[0]! as Element,
      existing = descendants(owner, c, 'ser'),
      prototype = existing[0]!;
    if (owner.localName === 'pieChart' && data.series.length !== 1)
      throw new RenderError('INVALID_DATA', 'Pie charts require one series');
    const element = (name: string, text?: string) => {
      const e = doc.createElementNS(c, 'c:' + name);
      if (text !== undefined) e.textContent = text;
      return e;
    };
    for (let i = 0; i < data.series.length; i++) {
      const input = data.series[i]!,
        ser = prototype.cloneNode(true) as typeof prototype;
      for (const name of ['idx', 'order']) {
        const n = descendants(ser, c, name)[0];
        if (!n) throw new RenderError('INVALID_PACKAGE', 'Chart series identity missing');
        n.setAttribute('val', String(i));
      }
      for (const name of ['tx', 'cat', 'val']) {
        const old = descendants(ser, c, name)[0];
        if (!old) throw new RenderError('UNSUPPORTED_TEMPLATE', 'Chart series requires tx, cat and val');
        const node = element(name);
        if (name === 'tx') node.appendChild(element('v', input.name));
        else {
          const literal = element(name === 'cat' ? 'strLit' : 'numLit');
          if (name === 'val') literal.appendChild(element('formatCode', 'General'));
          const count = element('ptCount');
          count.setAttribute('val', String(data.categories.length));
          literal.appendChild(count);
          (name === 'cat' ? data.categories : input.values).forEach((v, n) => {
            const point = element('pt');
            point.setAttribute('idx', String(n));
            point.appendChild(element('v', String(v)));
            literal.appendChild(point);
          });
          node.appendChild(literal);
        }
        ser.replaceChild(node, old);
      }
      // Point-specific formatting cannot safely follow resized/reordered data.
      for (const n of [...descendants(ser, c, 'dPt'), ...descendants(ser, c, 'dLbls')])
        n.parentNode!.removeChild(n);
      owner.insertBefore(ser, prototype);
    }
    for (const old of existing) owner.removeChild(old);
    const result = serialize(doc);
    if (result.length > limits.entry) throw new RenderError('LIMIT_EXCEEDED', 'Chart exceeds XML budget');
    parts.set(part, result);
  }
  if ([...parts.values()].reduce((n, b) => n + b.length, 0) > limits.expanded)
    throw new RenderError('LIMIT_EXCEEDED', 'Chart updates exceed package budget');
  return writePackage(parts);
}
