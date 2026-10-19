import type { Element, Node } from '@xmldom/xmldom';
import { createHash } from 'node:crypto';
import { openOfficePackage } from './ooxml.ts';
import { descendants, ns, parseXml } from './xml.ts';
import { formatNs } from './format-namespaces.ts';
import type { DocumentFormat } from './types.ts';
export interface DocumentUnit {
  location: string;
  kind: 'text' | 'cell' | 'row';
  value: string;
}
export interface DocumentSnapshot {
  format: DocumentFormat;
  sha256: string;
  units: DocumentUnit[];
}
export interface DocumentChange {
  location: string;
  kind: DocumentUnit['kind'];
  before?: string;
  after?: string;
}
export interface DocumentDiff {
  equal: boolean;
  binaryEqual: boolean;
  changes: DocumentChange[];
  masked: string[];
  totalChanges: number;
  truncated: false;
}
export const sha256 = (bytes: Uint8Array | string): string =>
  createHash('sha256').update(bytes).digest('hex');
const formats = new Set(['docx', 'xlsx', 'pptx', 'odt', 'ods', 'odp', 'txt', 'csv']);
function logicalText(root: Element, mode: 'office' | 'odf'): string {
  const stack: Node[] = [root],
    chunks: string[] = [];
  let length = 0,
    nodes = 0;
  const add = (text: string) => {
    length += text.length;
    if (length > 8 * 1024 * 1024) throw new Error('Semantic text exceeds budget');
    chunks.push(text);
  };
  while (stack.length) {
    const node = stack.pop()!;
    if (++nodes > 200_000) throw new Error('Semantic XML exceeds complexity budget');
    if (node.nodeType === 3 || node.nodeType === 4) {
      if (mode === 'odf') add(node.nodeValue ?? '');
      continue;
    }
    if (node.nodeType !== 1) continue;
    const e = node as Element;
    if (mode === 'office' && [ns.word, ns.drawing].includes(e.namespaceURI ?? '')) {
      if (e.localName === 't') {
        add(e.textContent ?? '');
        continue;
      }
      if (e.localName === 'tab') {
        add('\t');
        continue;
      }
      if (['br', 'cr'].includes(e.localName ?? '')) {
        add('\n');
        continue;
      }
    }
    if (mode === 'odf' && e.namespaceURI === formatNs.text) {
      if (e.localName === 'tab') {
        add('\t');
        continue;
      }
      if (e.localName === 'line-break') {
        add('\n');
        continue;
      }
      if (e.localName === 's') {
        const raw = e.getAttributeNS(formatNs.text, 'c') ?? '1';
        if (!/^[1-9]\d{0,4}$/.test(raw) || Number(raw) > 10000) throw new Error('Invalid ODF space repeat');
        add(' '.repeat(Number(raw)));
        continue;
      }
    }
    for (let i = node.childNodes.length - 1; i >= 0; i--) stack.push(node.childNodes.item(i)!);
  }
  return chunks.join('');
}
/** Text/cell comparison only: no page layout, image, accessibility or PDF verification. */
export async function snapshotDocument(bytes: Uint8Array, format: DocumentFormat): Promise<DocumentSnapshot> {
  bytes = Buffer.from(bytes);
  if (!formats.has(format)) throw new Error('Unsupported snapshot format');
  if (bytes.byteLength > 8 * 1024 * 1024) throw new Error('Snapshot exceeds 8 MiB');
  const units: DocumentUnit[] = [];
  const add = (location: string, kind: DocumentUnit['kind'], value: string) => {
    if (units.length >= 100_000) throw new Error('Snapshot exceeds 100000 units');
    units.push({ location, kind, value });
  };
  if (format === 'txt' || format === 'csv') {
    const source = new TextDecoder('utf-8', { fatal: true }).decode(bytes);
    // CSV records must preserve quoted newlines and escaped double quotes.
    if (format === 'txt')
      source.split(/\r\n|\n|\r/).forEach((value, i) => add(`text/line/${i + 1}`, 'text', value));
    else {
      let row: string[] = [],
        cell = '',
        quoted = false,
        closed = false,
        line = 1;
      const endCell = () => {
        row.push(cell);
        cell = '';
        closed = false;
      };
      const endRow = () => {
        endCell();
        add(`csv/row/${line++}`, 'row', JSON.stringify(row));
        row = [];
      };
      for (let i = 0; i < source.length; i++) {
        const c = source[i]!;
        if (quoted) {
          if (c === '"') {
            if (source[i + 1] === '"') {
              cell += '"';
              i++;
            } else {
              quoted = false;
              closed = true;
            }
          } else cell += c;
        } else if (c === ',') endCell();
        else if (c === '\r' || c === '\n') {
          if (c === '\r' && source[i + 1] === '\n') i++;
          endRow();
        } else if (c === '"' && !cell && !closed) quoted = true;
        else {
          if (closed || c === '"') throw new Error('Malformed CSV');
          cell += c;
        }
      }
      if (quoted) throw new Error('Unterminated CSV quote');
      if (cell || row.length || closed || !source.length) endRow();
    }
  } else {
    const office = await openOfficePackage(bytes);
    if (office.format !== format) throw new Error('Snapshot format does not match document');
    const parts = office.parts;
    const shared = parts.get('xl/sharedStrings.xml');
    const strings = shared
      ? descendants(parseXml(shared.toString(), 'xl/sharedStrings.xml'), ns.spreadsheet, 'si').map((e) =>
          descendants(e, ns.spreadsheet, 't')
            .map((t) => t.textContent ?? '')
            .join(''),
        )
      : [];
    for (const [part, content] of [...parts].sort(([a], [b]) => a.localeCompare(b))) {
      if (!part.endsWith('.xml')) continue;
      const doc = parseXml(content.toString(), part);
      if (format === 'xlsx' && /^xl\/worksheets\/[^/]+\.xml$/.test(part)) {
        for (const [i, cell] of descendants(doc, ns.spreadsheet, 'c').entries()) {
          const type = cell.getAttribute('t') ?? 'n';
          let value = descendants(cell, ns.spreadsheet, 'v')[0]?.textContent ?? '';
          if (type === 's') {
            if (!/^\d+$/.test(value) || strings[Number(value)] === undefined)
              throw new Error('Invalid shared string index');
            value = strings[Number(value)]!;
          }
          if (type === 'inlineStr')
            value = descendants(cell, ns.spreadsheet, 't')
              .map((t) => t.textContent ?? '')
              .join('');
          add(
            `${part}/cell/${cell.getAttribute('r') || i}`,
            'cell',
            JSON.stringify({
              type: ['s', 'inlineStr'].includes(type) ? 'string' : type,
              value,
              formula: descendants(cell, ns.spreadsheet, 'f')[0]?.textContent ?? null,
              formulaAttributes: Array.from(descendants(cell, ns.spreadsheet, 'f')[0]?.attributes ?? [])
                .map((a) => [a.name, a.value])
                .sort(([a], [b]) => a!.localeCompare(b!)),
            }),
          );
        }
      } else if (format === 'ods' && part === 'content.xml') {
        descendants(doc, formatNs.table, 'table-row').forEach((row, i) =>
          add(
            `${part}/row/${i + 1}`,
            'row',
            JSON.stringify({
              repeat: row.getAttributeNS(formatNs.table, 'number-rows-repeated'),
              cells: descendants(row, formatNs.table, 'table-cell').map((c) => ({
                text: logicalText(c, 'odf'),
                value: c.getAttributeNS(formatNs.office, 'value'),
                type: c.getAttributeNS(formatNs.office, 'value-type'),
                date: c.getAttributeNS(formatNs.office, 'date-value'),
                boolean: c.getAttributeNS(formatNs.office, 'boolean-value'),
                string: c.getAttributeNS(formatNs.office, 'string-value'),
                time: c.getAttributeNS(formatNs.office, 'time-value'),
                currency: c.getAttributeNS(formatNs.office, 'currency'),
                formula: c.getAttributeNS(formatNs.table, 'formula'),
                repeat: c.getAttributeNS(formatNs.table, 'number-columns-repeated'),
              })),
            }),
          ),
        );
      } else {
        const uri = format === 'docx' ? ns.word : format === 'pptx' ? ns.drawing : formatNs.text;
        descendants(doc, uri, 'p').forEach((p, i) =>
          add(
            `${part}/paragraph/${i + 1}`,
            'text',
            logicalText(p, format === 'docx' || format === 'pptx' ? 'office' : 'odf'),
          ),
        );
        if (format === 'odt' || format === 'odp')
          descendants(doc, formatNs.text, 'h').forEach((p, i) =>
            add(`${part}/heading/${i + 1}`, 'text', logicalText(p, 'odf')),
          );
      }
    }
  }
  if (new Set(units.map((u) => u.location)).size !== units.length)
    throw new Error('Duplicate semantic location');
  return { format, sha256: sha256(bytes), units };
}
/** Exact location masks; stale/duplicate masks fail instead of silently widening exclusions. */
export function compareDocuments(
  before: DocumentSnapshot,
  after: DocumentSnapshot,
  masks: string[] = [],
): DocumentDiff {
  for (const snapshot of [before, after]) {
    if (
      !snapshot ||
      !formats.has(snapshot.format) ||
      !/^[a-f0-9]{64}$/.test(snapshot.sha256) ||
      !Array.isArray(snapshot.units) ||
      snapshot.units.length > 100_000 ||
      snapshot.units.some(
        (u) =>
          !u ||
          typeof u.location !== 'string' ||
          !u.location ||
          u.location.length > 2048 ||
          typeof u.value !== 'string' ||
          u.value.length > 8 * 1024 * 1024 ||
          !['text', 'cell', 'row'].includes(u.kind),
      ) ||
      new Set(snapshot.units.map((u) => u.location)).size !== snapshot.units.length
    )
      throw new Error('Invalid document snapshot');
  }
  if (before.format !== after.format) throw new Error('Cannot compare different formats');
  if (
    !Array.isArray(masks) ||
    masks.length > 100 ||
    new Set(masks).size !== masks.length ||
    masks.some((m) => typeof m !== 'string' || !m || m.length > 1024)
  )
    throw new Error('Invalid comparison masks');
  const a = new Map(before.units.map((u) => [u.location, u])),
    b = new Map(after.units.map((u) => [u.location, u]));
  for (const mask of masks)
    if (!a.has(mask) || !b.has(mask)) throw new Error(`Mask must exist in both documents: ${mask}`);
  if (masks.length && masks.length >= new Set([...a.keys(), ...b.keys()]).size)
    throw new Error('Masks cannot exclude every semantic unit');
  const ignored = new Set(masks),
    changes: DocumentChange[] = [];
  for (const location of new Set([...a.keys(), ...b.keys()])) {
    const x = a.get(location),
      y = b.get(location);
    if (!ignored.has(location) && (x?.value !== y?.value || x?.kind !== y?.kind))
      changes.push({
        location,
        kind: y?.kind ?? x!.kind,
        ...(x ? { before: x.value } : {}),
        ...(y ? { after: y.value } : {}),
      });
  }
  return {
    equal: changes.length === 0,
    binaryEqual: before.sha256 === after.sha256,
    changes,
    masked: [...masks],
    totalChanges: changes.length,
    truncated: false,
  };
}
