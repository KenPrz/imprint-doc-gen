import { textMagic } from './text-template.ts';
import { formatNs as f } from './format-namespaces.ts';
import { openOfficePackage } from './ooxml.ts';
import { compile } from './index.ts';
import { standardBindings } from './standard-syntax.ts';
import { ns, parseXml, descendants } from './xml.ts';
import { RenderError } from './errors.ts';
import type { DocumentFormat } from './types.ts';
export interface TemplateField {
  expression: string;
  path: string;
  part: string;
  location: string;
  kind: 'value' | 'array' | 'image';
  required: boolean;
}
export interface Diagnostic {
  severity: 'error';
  code: string;
  message: string;
  part?: string;
  location?: string;
}
export interface Inspection {
  format?: DocumentFormat;
  valid: boolean;
  fields: TemplateField[];
  diagnostics: Diagnostic[];
}
export async function inspectTemplate(bytes: Uint8Array): Promise<Inspection> {
  const fields: TemplateField[] = [],
    diagnostics: Diagnostic[] = [];
  let format: DocumentFormat | undefined;
  const record = (error: unknown, part?: string, location?: string) =>
    diagnostics.push({
      severity: 'error',
      code: error instanceof RenderError ? error.code : 'INSPECTION_FAILED',
      message: error instanceof RenderError ? error.message : 'Template inspection failed',
      part: error instanceof RenderError ? (error.part ?? part) : part,
      location,
    });
  try {
    const inspect = (text: string, part: string, location: string, image = false) => {
      try {
        for (const binding of standardBindings(text, part))
          fields.push({
            expression: text.slice(binding.start, binding.end),
            path: binding.path,
            part,
            location,
            kind: image ? 'image' : binding.path.includes('[i') ? 'array' : 'value',
            required: binding.formatters?.some((f) => f.name === 'required') ?? false,
          });
      } catch (error) {
        record(error, part, location);
      }
    };
    if(Buffer.from(bytes).subarray(0,textMagic.length).toString()===textMagic){
      const compiled=await compile(bytes);format=compiled.format;
      const envelope=JSON.parse(Buffer.from(bytes).toString().slice(textMagic.length));
      inspect(envelope.source,'text','source');
      return {format,valid:diagnostics.length===0,fields,diagnostics};
    }
    const office = await openOfficePackage(bytes);
    format = office.format;
    for (const [part, content] of office.parts) {
      if (!part.endsWith('.xml')) continue;
      const doc = parseXml(content.toString(), part);
      if (format === 'docx' && part.startsWith('word/'))
        descendants(doc, ns.word, 'p').forEach((p, n) =>
          inspect(
            descendants(p, ns.word, 't')
              .map((t) => t.textContent)
              .join(''),
            part,
            `paragraph ${n + 1}`,
          ),
        );
      if (format === 'pptx')
        descendants(doc, ns.drawing, 'p').forEach((p, n) =>
          inspect(
            descendants(p, ns.drawing, 't')
              .map((t) => t.textContent)
              .join(''),
            part,
            `paragraph ${n + 1}`,
          ),
        );
      if (['odt', 'ods', 'odp'].includes(format)) {
        descendants(doc, f.text, 'p').forEach((p, n) =>
          inspect(p.textContent ?? '', part, `paragraph ${n + 1}`),
        );
        descendants(doc, f.text, 'h').forEach((p, n) =>
          inspect(p.textContent ?? '', part, `heading ${n + 1}`),
        );
        descendants(doc, f.svg, 'desc').forEach((p, n) =>
          inspect(p.textContent ?? '', part, `picture ${n + 1}`, true),
        );
      }
      if (format === 'xlsx') {
        const shared: string[] = [];
        for (const [path, value] of office.parts)
          if (/sharedStrings\.xml$/.test(path))
            shared.push(
              ...descendants(parseXml(value.toString(), path), ns.spreadsheet, 'si').map((si) =>
                descendants(si, ns.spreadsheet, 't')
                  .map((t) => t.textContent)
                  .join(''),
              ),
            );
        for (const c of descendants(doc, ns.spreadsheet, 'c')) {
          const text =
            c.getAttribute('t') === 's'
              ? (shared[Number(descendants(c, ns.spreadsheet, 'v')[0]?.textContent)] ?? '')
              : descendants(c, ns.spreadsheet, 't')
                  .map((t) => t.textContent)
                  .join('');
          inspect(text, part, `cell ${c.getAttribute('r') ?? '?'}`);
        }
      }
      for (const node of descendants(doc, '*', '*'))
        for (const name of ['descr', 'title'])
          if (node.hasAttribute(name))
            inspect(node.getAttribute(name)!, part, `picture ${node.getAttribute('id') ?? '?'}`, true);
    }
    if (!diagnostics.length)
      try {
        await compile(bytes);
      } catch (error) {
        record(error);
      }
  } catch (error) {
    record(error);
  }
  return { format, valid: diagnostics.length === 0, fields, diagnostics };
}
