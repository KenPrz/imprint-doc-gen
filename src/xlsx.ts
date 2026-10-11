import { standardImageResolver } from './image-binding.ts';
import { standardBindings, parseStandardPath } from './standard-syntax.ts';
import { compileRowRenderer } from './xlsx-rows.ts';
import { RenderError } from './errors.ts';
import { validateData } from './data.ts';
import { checkPath, lookup, scalar, type Scope } from './scope.ts';
import { createRenderPackage } from './render-package.ts';
import { limits, relsPath, resolvePart, type PackageParts } from './package.ts';
import type { OfficePackage } from './ooxml.ts';
import type { CompiledTemplate, RenderOptions } from './types.ts';
import { ns, indexXml, parseXml, descendants, ancestor, applyEdits, escapeXml, type ElementSpan, type AttributeSpan, type Edit } from './xml.ts';
import { compileCellText, renderCellText, readLeaf, attributeEdit, type CellText } from './xlsx-text.ts';

export interface CellPlan { node: ElementSpan; payload: ElementSpan; text: CellText; shared?: number }
export interface SheetPlan { part: string; name: string; source: string; nodes: ElementSpan[]; cells: CellPlan[] }
interface Picture { id: string; size?: { width: number; height: number }; standard?: boolean; path: string; description: AttributeSpan; reference: AttributeSpan }
interface DrawingPlan { part: string; source: string; pictures: Picture[] }
const attr = (node: ElementSpan, name: string) => node.attributes.find(a => a.name === name)?.value;
function linksFor(parts: PackageParts, owner: string) {
  const path = relsPath(owner), bytes = parts.get(path);
  return bytes ? descendants(parseXml(bytes.toString('utf8'), path), ns.packageRelationship, 'Relationship') : [];
}
function childrenOf(nodes: ElementSpan[]): Map<ElementSpan, ElementSpan[]> {
  const children = new Map<ElementSpan, ElementSpan[]>();
  for (const node of nodes) if (node.parent) { const list = children.get(node.parent) ?? []; list.push(node); children.set(node.parent, list); }
  return children;
}
function checkedRoot(nodes: ElementSpan[], uri: string, local: string, part: string) {
  const root = nodes[0];
  if (!root || root.uri !== uri || root.local !== local) throw new RenderError('INVALID_PACKAGE', `Expected ${local} root`, part);
  return root;
}
function compileDrawing(parts: PackageParts, part: string): DrawingPlan {
  const source = parts.get(part)!.toString('utf8'), nodes = indexXml(source, part);
  checkedRoot(nodes, ns.spreadsheetDrawing, 'wsDr', part);
  const pictures: Picture[] = [], owners = new Set<number>(), links = linksFor(parts, part);
  for (const node of nodes) for (const description of node.attributes.filter(a => a.value.includes('{{') || standardBindings(a.value, part).length)) {
    if (node.uri !== ns.spreadsheetDrawing || node.local !== 'cNvPr' || !['descr', 'title', 'name'].includes(description.local)) throw new RenderError('UNSUPPORTED_TEMPLATE', 'Only picture metadata bindings are supported in XLSX drawings', part);
    const match = /^\s*\{\{image:([^{}]+)\}\}\s*$/.exec(description.value);
    const standard = !match, single = /^\s*\{([^{}]+)\}\s*$/.exec(description.value);
    if (!match && !single) throw new RenderError('TEMPLATE_SYNTAX', 'Image binding must occupy the whole picture description', part);
    const path = match ? match[1]!.trim() : parseStandardPath(single![1]!, part); checkPath(path.replace(/\[i(?:\+1)?\]/g, ''), part);
    const owner = ancestor(node, ns.spreadsheetDrawing, 'pic');
    if (!owner || owners.has(owner.start) || owner.parent?.uri !== ns.spreadsheetDrawing || !['oneCellAnchor', 'twoCellAnchor', 'absoluteAnchor'].includes(owner.parent.local)) throw new RenderError('UNSUPPORTED_TEMPLATE', 'Expected one binding on an anchored spreadsheet picture', part);
    const blips = nodes.filter(n => n.uri === ns.drawing && n.local === 'blip' && n.start > owner.start && n.end < owner.end);
    const reference = blips[0]?.attributes.find(a => a.uri === ns.relationship && a.local === 'embed');
    if (blips.length !== 1 || !reference || blips[0]!.attributes.some(a => a.uri === ns.relationship && a.local === 'link')) throw new RenderError('UNSUPPORTED_TEMPLATE', 'Expected one embedded picture per binding', part);
    if (!links.some(link => link.getAttribute('Id') === reference.value && link.getAttribute('Type') === `${ns.relationship}/image` && link.getAttribute('TargetMode') !== 'External')) throw new RenderError('INVALID_PACKAGE', 'Picture references an undefined image relationship', part);
    const extent = nodes.find(n => n.local === 'ext' && n.start > owner.start && n.end < owner.end);
    const size = extent ? { width: Math.max(1, Math.round(Number(attr(extent, 'cx')) / 9525)), height: Math.max(1, Math.round(Number(attr(extent, 'cy')) / 9525)) } : undefined;
    owners.add(owner.start); pictures.push({ id: attr(node, 'id') ?? '', size, path, description, reference, standard });
  }
  return { part, source, pictures };
}
function recalculate(source: string, nodes: ElementSpan[]): string {
  const root = nodes[0]!, calc = nodes.find(n => n.parent === root && n.uri === ns.spreadsheet && n.local === 'calcPr');
  if (calc) return applyEdits(source, ['fullCalcOnLoad', 'forceFullCalc'].map(name => attributeEdit(calc, name, '1')).concat(attributeEdit(calc, 'calcMode', 'auto')), limits.entry);
  const prefix = root.name.includes(':') ? root.name.split(':')[0] + ':' : '';
  // calcPr follows definedNames and precedes these optional workbook children.
  const tail = nodes.find(n => n.parent === root && n.uri === ns.spreadsheet && ['oleSize', 'customWorkbookViews', 'pivotCaches', 'smartTagPr', 'smartTagTypes', 'webPublishing', 'fileRecoveryPr', 'webPublishObjects', 'extLst'].includes(n.local));
  const start = tail?.start ?? root.closeStart;
  return applyEdits(source, [{ start, end: start, value: `<${prefix}calcPr calcMode="auto" fullCalcOnLoad="1" forceFullCalc="1"/>` }], limits.entry);
}
export async function compileXlsx({ parts: original, mainPart }: OfficePackage): Promise<CompiledTemplate> {
  const workbook = original.get(mainPart)!.toString('utf8'), bookNodes = indexXml(workbook, mainPart), bookLinks = linksFor(original, mainPart);
  const sharedLinks = bookLinks.filter(l => l.getAttribute('Type') === `${ns.relationship}/sharedStrings`);
  if (sharedLinks.length > 1) throw new RenderError('INVALID_PACKAGE', 'Workbook has multiple shared-string tables', mainPart);
  const sharedPart = sharedLinks.length ? resolvePart(mainPart, sharedLinks[0]!.getAttribute('Target')!) : undefined;
  const sharedSource = sharedPart ? original.get(sharedPart)!.toString('utf8') : '';
  const sharedNodes = sharedPart ? indexXml(sharedSource, sharedPart) : [];
  const sharedRoot = sharedPart ? checkedRoot(sharedNodes, ns.spreadsheet, 'sst', sharedPart) : undefined;
  const sharedChildren = childrenOf(sharedNodes);
  const shared = sharedNodes.filter(n => n.parent === sharedRoot && n.uri === ns.spreadsheet && n.local === 'si').map(node => compileCellText(sharedSource, node, sharedChildren, sharedPart!));
  const sheets: SheetPlan[] = [], drawingParts = new Set<string>(), seenSheets = new Set<string>();
  let sharedReferences = 0;
  for (const sheet of bookNodes.filter(n => n.uri === ns.spreadsheet && n.local === 'sheet')) {
    const id = sheet.attributes.find(a => a.uri === ns.relationship && a.local === 'id')?.value;
    const link = bookLinks.find(l => l.getAttribute('Id') === id);
    if (!link || link.getAttribute('Type') !== `${ns.relationship}/worksheet`) throw new RenderError('UNSUPPORTED_TEMPLATE', 'Only worksheet sheets are supported', mainPart);
    const part = resolvePart(mainPart, link.getAttribute('Target')!);
    if (seenSheets.has(part)) throw new RenderError('INVALID_PACKAGE', 'Multiple sheets reference the same worksheet', mainPart);
    seenSheets.add(part);
    const source = original.get(part)!.toString('utf8'), nodes = indexXml(source, part), root = checkedRoot(nodes, ns.spreadsheet, 'worksheet', part);
    const children = childrenOf(nodes);
    const cells: CellPlan[] = [];
    for (const node of nodes) {
      if (node.uri === ns.spreadsheet && ['f', 'formula', 'formula1', 'formula2'].includes(node.local) && (readLeaf(source, node).includes('{{') || standardBindings(readLeaf(source, node), part).length)) throw new RenderError('UNSUPPORTED_TEMPLATE', 'Formula templating is not supported', part);
      if (node.uri !== ns.spreadsheet || node.local !== 'c') continue;
      const contents = children.get(node) ?? [], type = attr(node, 't');
      const values = contents.filter(n => n.uri === ns.spreadsheet && ['is', 'v'].includes(n.local));
      const formula = contents.some(n => n.uri === ns.spreadsheet && n.local === 'f');
      if (type !== 's' && type !== 'inlineStr') {
        if (values.some(n => (source.slice(n.openEnd, n.closeStart).includes('{{') || standardBindings(readLeaf(source, n), part).length))) throw new RenderError('UNSUPPORTED_TEMPLATE', 'Bindings require inline or shared-string cells', part);
        continue;
      }
      if (values.length !== 1 || values[0]!.local !== (type === 's' ? 'v' : 'is')) throw new RenderError('INVALID_PACKAGE', 'Invalid spreadsheet string cell payload', part);
      const payload = values[0]!;
      let text: CellText, sharedIndex: number | undefined;
      if (type === 's') {
        const raw = readLeaf(source, payload).trim();
        if (!/^\d+$/.test(raw) || !Number.isSafeInteger(Number(raw)) || !shared[Number(raw)]) throw new RenderError('INVALID_PACKAGE', 'Invalid shared-string index', part);
        sharedIndex = Number(raw); text = shared[sharedIndex]!; sharedReferences++;
      } else text = compileCellText(source, payload, children, part);
      if (!text.bindings.length) continue;
      if (formula) throw new RenderError('UNSUPPORTED_TEMPLATE', 'Formula cells cannot contain template bindings', part);
      if (contents.some(n => n.local === 'extLst')) throw new RenderError('UNSUPPORTED_TEMPLATE', 'Bound cell extensions require a later adapter', part);
      cells.push({ node, payload, text, shared: sharedIndex });
    }
    const links = linksFor(original, part);
    for (const drawing of nodes.filter(n => n.parent === root && n.uri === ns.spreadsheet && n.local === 'drawing')) {
      const id = drawing.attributes.find(a => a.uri === ns.relationship && a.local === 'id')?.value;
      const link = links.find(l => l.getAttribute('Id') === id && l.getAttribute('Type') === `${ns.relationship}/drawing`);
      if (!link) throw new RenderError('INVALID_PACKAGE', 'Worksheet references an undefined drawing', part);
      drawingParts.add(resolvePart(part, link.getAttribute('Target')!));
    }
    sheets.push({ part, name: attr(sheet, 'name') ?? '', source, nodes, cells });
  }
  if (!sheets.length) throw new RenderError('INVALID_PACKAGE', 'Workbook has no worksheets', mainPart);
  const drawings = [...drawingParts].map(part => compileDrawing(original, part));
  const rows = compileRowRenderer({ original, mainPart, workbook, bookNodes, sheets, sharedPart, sharedSource, sharedNodes, shared });
  if (!rows && drawings.some(d=>d.pictures.some(p=>p.path.includes('[i')))) throw new RenderError('TEMPLATE_EXPRESSION','Item picture bindings require a repeated row');
  if (!rows && sheets.some(s => s.cells.some(c => c.text.bindings.some(b => b.path.startsWith('../'))))) throw new RenderError('TEMPLATE_EXPRESSION', 'Parent paths require a repeated row');
  const renderPictures = async (output: ReturnType<typeof createRenderPackage>, attachments: Record<string, Uint8Array>, data: Record<string, unknown>, drawingSources = original, pictureScopes = new Map<string, Map<string, { scope: Scope; path: string }>>()) => {
    const imageKey = standardImageResolver(attachments);
    for (const initial of drawings) {
      const drawing = drawingSources === original ? initial : compileDrawing(drawingSources, initial.part);
      const edits: Edit[] = [];
      for (const picture of drawing.pictures) {
        const repeated = pictureScopes.get(drawing.part)?.get(picture.id), scope = repeated?.scope ?? [data];
        const path = repeated && picture.standard ? picture.path.startsWith(`${repeated.path}[i].`) ? picture.path.slice(repeated.path.length + 4) : '../' + picture.path : picture.path;
        const key = picture.standard ? imageKey(path, scope, drawing.part) : scope.length === 1 && Object.hasOwn(attachments, path) ? path : scalar(lookup(scope, path, drawing.part), path, drawing.part);
        if (!key && picture.standard) { edits.push({ start: picture.description.start, end: picture.description.end, value: 'Image placeholder' }); continue; }
        const id = await output.addImage(drawing.part, key!, picture.size);
        edits.push({ start: picture.reference.start, end: picture.reference.end, value: id }, { start: picture.description.start, end: picture.description.end, value: picture.standard ? 'Image attachment' : escapeXml(`Attachment ${key}`) });
      }
      if (edits.length) output.setPart(drawing.part, Buffer.from(applyEdits(drawing.source, edits, limits.entry)));
    }
  };
  return Object.freeze({ format: 'xlsx' as const, async render(data: Record<string, unknown>, options: RenderOptions = {}) {
    validateData(data);
    const attachments = Object.assign(Object.create(null), options.attachments ?? {}) as Record<string, Uint8Array>;
    if (rows) {
      const expanded = rows(data), output = createRenderPackage(expanded.parts, mainPart, attachments, options.images);
      const source = expanded.parts.get(mainPart)!.toString();
      output.setPart(mainPart, Buffer.from(recalculate(source, indexXml(source, mainPart))));
      await renderPictures(output, attachments, data, expanded.parts, expanded.pictureScopes);
      return output.finish();
    }
    const output = createRenderPackage(original, mainPart, attachments, options.images);
    let sharedInsertedBytes = 0;
    const renderedShared = shared.map(plan => {
      if (!plan.bindings.length) return undefined;
      const result = renderCellText(plan, data, sharedPart!);
      sharedInsertedBytes += result.edits.reduce((sum, edit) => sum + Buffer.byteLength(edit.value), 0);
      if (sharedInsertedBytes > limits.entry) throw new RenderError('LIMIT_EXCEEDED', 'Shared-string replacements exceed the part byte budget', sharedPart);
      return result;
    });
    let remainingReferences = sharedReferences, changed = false;
    for (const sheet of sheets) {
      const edits: Edit[] = []; let insertedBytes = 0;
      const add = (...values: Edit[]) => {
        for (const edit of values) {
          insertedBytes += Buffer.byteLength(edit.value);
          if (insertedBytes > limits.entry) throw new RenderError('LIMIT_EXCEEDED', 'Cell replacements exceed the part byte budget', sheet.part);
          edits.push(edit);
        }
      };
      for (const cell of sheet.cells) {
        const result = cell.shared === undefined ? renderCellText(cell.text, data, sheet.part) : renderedShared[cell.shared]!;
        if (cell.text.whole && (result.value === null || typeof result.value === 'number' || typeof result.value === 'boolean')) {
          const prefix = cell.node.name.includes(':') ? cell.node.name.split(':')[0] + ':' : '';
          add(attributeEdit(cell.node, 't', typeof result.value === 'boolean' ? 'b' : 'n'));
          add({ start: cell.payload.start, end: cell.payload.end, value: result.value === null ? '' : `<${prefix}v>${typeof result.value === 'boolean' ? Number(result.value) : result.value}</${prefix}v>` });
          if (cell.shared !== undefined) remainingReferences--;
        } else if (cell.shared === undefined) add(...result.edits);
      }
      if (edits.length) output.setPart(sheet.part, Buffer.from(applyEdits(sheet.source, edits, limits.entry)));
      if (sheet.cells.length) changed = true;
    }
    if (sharedPart && renderedShared.some(Boolean)) {
      const edits = renderedShared.flatMap(result => result?.edits ?? []);
      edits.push(attributeEdit(sharedRoot!, 'count', String(remainingReferences)));
      output.setPart(sharedPart, Buffer.from(applyEdits(sharedSource, edits, limits.entry)));
    }
    if (changed) output.setPart(mainPart, Buffer.from(recalculate(workbook, bookNodes)));
    await renderPictures(output, attachments, data);
    return output.finish();
  } });
}
