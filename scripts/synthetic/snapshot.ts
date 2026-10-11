import { readPackage } from '../../src/package.ts';
import { parseXml, descendants, ns } from '../../src/xml.ts';
import type { Node } from '@xmldom/xmldom';
export async function snapshot(bytes: Uint8Array) {
  const parts = await readPackage(bytes), stories: Record<string, unknown> = {};
  const text = (node: Node): string => {
    if (node.nodeType === 1) {
      const element = node as Node & { namespaceURI: string; localName: string };
      if (element.namespaceURI === ns.word) {
        if (element.localName === 't') return node.textContent ?? '';
        if (element.localName === 'tab') return '\t';
        if (element.localName === 'br') return '\n';
      }
    }
    return Array.from(node.childNodes).map(text).join('');
  };
  for (const [name, value] of [...parts].sort()) if (/^word\/(document|header\d*|footer\d*)\.xml$/.test(name)) {
    const xml = parseXml(value.toString(), name);
    stories[name] = {
      paragraphs: descendants(xml, ns.word, 'p').map(text),
      tables: descendants(xml, ns.word, 'tbl').map(table => descendants(table, ns.word, 'tr').map(row => descendants(row, ns.word, 'tc').map(cell => descendants(cell, ns.word, 'p').map(text)))),
      pictures: descendants(xml, ns.wordDrawing, 'extent').map(node => [node.getAttribute('cx'), node.getAttribute('cy')]),
    };
  }
  return stories;
}
