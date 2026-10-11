import { carboneLookup } from './carbone.ts';
import { checkPath } from './scope.ts';
import { RenderError } from './errors.ts';
import { lookup, type Scope } from './scope.ts';
import { limits, relsPath, resolvePart, type PackageParts } from './package.ts';
import { ns, indexXml, ancestor, applyEdits, escapeXml, parseXml, descendants, serialize, type ElementSpan, type Edit } from './xml.ts';
import { attributeEdit, renderCellText, readLeaf, decodeXstring, encodeXstring, type CellText } from './xlsx-text.ts';
import { address, a1, maxRow, RowMap } from './xlsx-references.ts';
import { parseFormula, rewriteFormula, evaluateFormula, type Value } from './xlsx-formula.ts';
import type { CellPlan, SheetPlan } from './xlsx.ts';

interface Row { node: ElementSpan; number: number; cells: ElementSpan[]; bindings: CellPlan[]; path?: string; carbone?: boolean; discard?: boolean }
interface Sheet extends SheetPlan { rows: Row[] }
interface LinkedPart { part: string; sheet: string; source: string; nodes: ElementSpan[] }
interface Context { original: PackageParts; mainPart: string; workbook: string; bookNodes: ElementSpan[]; sheets: SheetPlan[]; sharedPart?: string; sharedSource: string; sharedNodes: ElementSpan[]; shared: CellText[] }
const attr = (node: ElementSpan, name: string) => node.attributes.find(a => a.name === name)?.value;
const prefix = (node: ElementSpan) => node.name.includes(':') ? node.name.split(':')[0] + ':' : '';
const childIndexes = new WeakMap<ElementSpan[], Map<ElementSpan, ElementSpan[]>>();
function children(nodes: ElementSpan[], parent: ElementSpan): ElementSpan[] {
  let index = childIndexes.get(nodes);
  if (!index) {
    index = new Map();
    for (const node of nodes) if (node.parent) { const list = index.get(node.parent) ?? []; list.push(node); index.set(node.parent, list); }
    childIndexes.set(nodes, index);
  }
  return index.get(parent) ?? [];
}
function boundedSet(parts: PackageParts) {
  let total = [...parts.values()].reduce((sum, b) => sum + b.length, 0);
  return (part: string, source: string | Buffer) => {
    const bytes = Buffer.from(source); total += bytes.length - (parts.get(part)?.length ?? 0);
    if (bytes.length > limits.entry || total > limits.expanded) throw new RenderError('LIMIT_EXCEEDED', 'Expanded spreadsheet exceeds its byte budget', part);
    parts.set(part, bytes);
  };
}
function rangeEdits(source: string, nodes: ElementSpan[], map: RowMap, maps: Map<string, RowMap>, sheet: string): Edit[] {
  const edits: Edit[] = [], removed = new Set<ElementSpan>();
  const rangeAttributes: Record<string, string[]> = { dimension: ['ref'], mergeCell: ['ref'], autoFilter: ['ref'], conditionalFormatting: ['sqref'], dataValidation: ['sqref'], hyperlink: ['ref'], selection: ['sqref', 'activeCell'], pane: ['topLeftCell'], table: ['ref'] };
  for (const node of nodes) {
    if (node.uri !== ns.spreadsheet) continue;
    for (const name of rangeAttributes[node.local] ?? []) {
      const value = attr(node, name); if (!value) continue;
      let ranges: string[];
      if (['activeCell', 'topLeftCell'].includes(name)) {
        const reference = address(value); reference.row = map.map(reference.row, 'start') ?? 1; ranges = [a1(reference)];
      } else ranges = value.trim().split(/\s+/).map(range => map.range(range)).filter((range): range is string => !!range);
      if (!ranges.length && ['dimension', 'selection'].includes(node.local)) ranges = ['A1'];
      if (!ranges.length) {
        if (['conditionalFormatting', 'dataValidation', 'hyperlink', 'mergeCell', 'autoFilter'].includes(node.local)) { removed.add(node); break; }
        throw new RenderError('UNSUPPORTED_TEMPLATE', `Row removal leaves an empty ${node.local} range`);
      }
      if (ranges.join(' ') !== value) edits.push(attributeEdit(node, name, ranges.join(' ')));
    }
    if (['formula', 'formula1', 'formula2'].includes(node.local)) {
      const text = readLeaf(source, node);
      const result = rewriteFormula(parseFormula(text, true), maps, sheet);
      if (result !== text) edits.push({ start: node.openEnd, end: node.closeStart, value: escapeXml(result) });
    }
  }
  const kept = edits.filter(e => ![...removed].some(n => e.start >= n.start && e.end <= n.end));
  for (const node of removed) kept.push({ start: node.start, end: node.end, value: '' });
  for (const parent of new Set([...removed].map(n => n.parent!))) if (attr(parent, 'count') !== undefined) {
    kept.push(attributeEdit(parent, 'count', String(children(nodes, parent).filter(n => !removed.has(n)).length)));
  }
  return kept;
}
function inlineShared(source: string, entry: ElementSpan, result: ReturnType<typeof renderCellText>): string {
  const inner = applyEdits(source.slice(entry.openEnd, entry.closeStart), result.edits.map(e => ({ ...e, start: e.start - entry.openEnd, end: e.end - entry.openEnd })), limits.entry);
  // Shared strings can use different namespace aliases from the worksheet.
  const declarations = [...entry.namespaces].filter(([key]) => key !== 'xml').map(([key, uri]) => ` xmlns${key ? ':' + key : ''}="${escapeXml(uri)}"`).join('');
  let alias = 'imprintS'; while (entry.namespaces.has(alias)) alias += '_';
  return `<${alias}:is xmlns:${alias}="${ns.spreadsheet}"${declarations}>${inner}</${alias}:is>`;
}
export function compileRowRenderer(context: Context) {
  const { original, sheets: inputSheets, mainPart, workbook, bookNodes, shared, sharedSource, sharedNodes, sharedPart } = context;
  if (!inputSheets.some(s => s.cells.some(c => c.text.bindings.some(b => b.kind !== 'value' || (b.carbone && b.path.includes('[i')))))) return undefined;
  const remainingShared = shared.filter(text => !text.bindings.length);
  let nextShared = 0;
  const sharedRemap = shared.map(text => text.bindings.length ? -1 : nextShared++);
  const sharedEntries = sharedNodes.filter(n => n.uri === ns.spreadsheet && n.local === 'si');
  const sheets: Sheet[] = inputSheets.map(sheet => {
    const rows: Row[] = [], rowNodes = sheet.nodes.filter(n => n.uri === ns.spreadsheet && n.local === 'row');
    let previous = 0;
    for (const node of rowNodes) {
      const number = Number(attr(node, 'r'));
      if (!Number.isInteger(number) || number <= previous || number > maxRow) throw new RenderError('INVALID_PACKAGE', 'Loop workbooks require explicit ascending row numbers', sheet.part);
      previous = number;
      const cells = children(sheet.nodes, node).filter(n => n.uri === ns.spreadsheet && n.local === 'c'), seen = new Set<string>();
      for (const cell of cells) {
        const ref = address(attr(cell, 'r') ?? '');
        if (ref.row !== number || seen.has(ref.column)) throw new RenderError('INVALID_PACKAGE', 'Duplicate cell or mismatched row coordinate', sheet.part);
        seen.add(ref.column);
      }
      const bindings = sheet.cells.filter(c => c.node.parent === node), markers = bindings.flatMap(c => c.text.bindings.filter(b => b.kind !== 'value'));
      if (markers.length && (markers.length !== 2 || markers[0]!.kind !== 'open' || markers[1]!.kind !== 'close')) throw new RenderError('TEMPLATE_STRUCTURE', 'Each XLSX loop must open and close in the same single row', sheet.part);
      rows.push({ node, number, cells, bindings, path: markers[0]?.path });
    }
    if (sheet.cells.some(c => !rows.some(r => r.node === c.node.parent))) throw new RenderError('INVALID_PACKAGE', 'Cell binding is outside a worksheet row', sheet.part);
    const hasCarboneRepeat = rows.some(row => row.bindings.some(c => c.text.bindings.some(b => b.carbone && b.path.includes('[i'))));
    if (hasCarboneRepeat && rows.some(row => row.path)) throw new RenderError('UNSUPPORTED_TEMPLATE', 'Do not mix Carbone and legacy repeats in one worksheet', sheet.part);
    for (let index = 0; index < rows.length; index++) {
      const row = rows[index]!;
      const tags = row.bindings.flatMap(c => c.text.bindings).filter(b => b.carbone && b.path.includes('[i'));
      if (!tags.length || row.discard) continue;
      const paths = new Set(tags.map(tag => tag.path.split('[i]')[0]!));
      if (paths.size !== 1 || tags.some(tag => !/^[\w$.]+\[i\]\.[\w$.]+$/.test(tag.path))) throw new RenderError('UNSUPPORTED_CARBONE_TEMPLATE', 'XLSX repeats support one object array per row', sheet.part);
      const path = [...paths][0]!, marker = rows[index + 1];
      if (!marker || marker.number !== row.number + 1) throw new RenderError('TEMPLATE_STRUCTURE', 'Expected an adjacent [i+1] marker row', sheet.part);
      const markerTags = marker.bindings.flatMap(c => c.text.bindings);
      if (!markerTags.length || markerTags.some(tag => !tag.carbone || !(tag.path === `${path}[i+1]` || tag.path.startsWith(`${path}[i+1].`)))) throw new RenderError('TEMPLATE_STRUCTURE', 'Expected matching [i+1] bindings in the marker row', sheet.part);
      row.path = path; row.carbone = true; marker.discard = true;
      // Each shared-string use needs its own scope; never mutate the shared plan.
      row.bindings = row.bindings.map(cell => {
        const bindings = cell.text.bindings.map(binding => {
          if (!binding.carbone) return binding;
          const scoped = binding.path.startsWith(`${path}[i].`) ? binding.path.slice(path.length + 4) : '../' + binding.path;
          checkPath(scoped, sheet.part); return { ...binding, path: scoped };
        });
        return { ...cell, text: { ...cell.text, bindings, whole: cell.text.whole ? bindings.find(b => b.kind === 'value')!.path : undefined } };
      });
    }
    for (const row of rows) if (row.path && !row.carbone) row.bindings = row.bindings.map(cell => {
      const bindings = cell.text.bindings.map(binding => binding.carbone ? { ...binding, path: '../' + binding.path } : binding);
      return { ...cell, text: { ...cell.text, bindings, whole: cell.text.whole ? bindings.find(b => b.kind === 'value')!.path : undefined } };
    });
    for (const row of rows) if (!row.path && !row.discard) for (const cell of row.bindings) for (const binding of cell.text.bindings) if (binding.carbone) checkPath(binding.path, sheet.part);
    return { ...sheet, rows };
  });
  const knownNames = new Set<string>();
  for (const sheet of sheets) { if (knownNames.has(sheet.name.toLowerCase())) throw new RenderError('INVALID_PACKAGE', 'Duplicate worksheet name'); knownNames.add(sheet.name.toLowerCase()); }
  const tables: LinkedPart[] = [], drawings: LinkedPart[] = [], linkedOwners = new Set<string>();
  for (const sheet of sheets) {
    const loopRows = sheet.rows.filter(r => r.path || r.discard).map(r => r.number);
    for (const node of sheet.nodes) {
      if (node.uri !== ns.spreadsheet) continue;
      if (['extLst', 'rowBreaks', 'legacyDrawing', 'legacyDrawingHF', 'oleObjects', 'controls', 'pivotTableParts'].includes(node.local)) throw new RenderError('UNSUPPORTED_TEMPLATE', `XLSX loops with ${node.local} require a later adapter`, sheet.part);
      if (node.local === 'f') {
        if (node.attributes.some(a => a.name !== 't' || a.value !== 'normal')) throw new RenderError('UNSUPPORTED_FORMULA', 'Shared, array and data-table formulas require a later adapter', sheet.part);
        parseFormula(readLeaf(sheet.source, node));
      }
      if (node.local === 'mergeCell') {
        const refs = attr(node, 'ref')!.split(':').map(address);
        if (loopRows.some(r => r >= refs[0]!.row && r <= (refs[1] ?? refs[0])!.row)) throw new RenderError('UNSUPPORTED_TEMPLATE', 'Merged cells intersect a repeated row', sheet.part);
      }
      if (node.local === 'pane' && loopRows.some(r => r <= Number(attr(node, 'ySplit') ?? 0))) throw new RenderError('UNSUPPORTED_TEMPLATE', 'A repeated row intersects frozen panes', sheet.part);
    }
    const links = original.get(relsPath(sheet.part));
    if (!links) continue;
    for (const link of descendants(parseXml(links.toString(), sheet.part), ns.packageRelationship, 'Relationship')) {
      const type = link.getAttribute('Type')!.split('/').at(-1)!;
      if (['comments', 'threadedComment', 'vmlDrawing', 'pivotTable', 'queryTable'].includes(type)) throw new RenderError('UNSUPPORTED_TEMPLATE', `XLSX loops with ${type} require a later adapter`, sheet.part);
      if (!['table', 'drawing'].includes(type)) continue;
      const part = resolvePart(sheet.part, link.getAttribute('Target')!);
      if (linkedOwners.has(part)) throw new RenderError('UNSUPPORTED_TEMPLATE', 'Shared worksheet drawings or tables cannot be moved independently', part);
      linkedOwners.add(part);
      const source = original.get(part)!.toString(), nodes = indexXml(source, part), plan = { part, sheet: sheet.name, source, nodes };
      if (type === 'table') {
        if (nodes.some(n => ['calculatedColumnFormula', 'totalsRowFormula', 'extLst'].includes(n.local))) throw new RenderError('UNSUPPORTED_TEMPLATE', 'Calculated table columns and table extensions require a later adapter', part);
        const range = address(attr(nodes[0]!, 'ref')!.split(':')[0]!);
        if (loopRows.includes(range.row)) throw new RenderError('UNSUPPORTED_TEMPLATE', 'A table header cannot be repeated', part);
        tables.push(plan);
      } else {
        if (nodes.some(n => ['chart', 'grpSp', 'extLst'].includes(n.local))) throw new RenderError('UNSUPPORTED_TEMPLATE', 'Charts, grouped drawings and drawing extensions require a later loop adapter', part);
        for (const anchor of nodes.filter(n => n.parent === nodes[0])) {
          const positions = nodes.filter(n => n.start > anchor.start && n.end < anchor.end && n.uri === ns.spreadsheetDrawing && n.local === 'row').map(n => Number(readLeaf(source, n)) + 1);
          if (positions.length && loopRows.some(r => r >= Math.min(...positions) && r <= Math.max(...positions))) throw new RenderError('UNSUPPORTED_TEMPLATE', 'A drawing intersects a repeated row; picture cloning requires a later adapter', part);
        }
        drawings.push(plan);
      }
    }
  }
  for (const node of bookNodes.filter(n => n.uri === ns.spreadsheet && n.local === 'definedName')) parseFormula(readLeaf(workbook, node), true);
  if (bookNodes.some(n => n.local === 'externalReferences' || n.local === 'pivotCaches')) throw new RenderError('UNSUPPORTED_TEMPLATE', 'External references and pivot caches require a later loop adapter', mainPart);
  return (data: Record<string, unknown>) => {
    const maps = new Map<string, RowMap>(), scopes = new Map<Row, Scope[]>(); let iterations = 0;
    for (const sheet of sheets) {
      const expansions = [];
      for (const row of sheet.rows) if (row.path || row.discard) {
        const items = row.discard ? [] : row.carbone ? carboneLookup([data], row.path!, sheet.part) ?? [] : lookup([data], row.path!, sheet.part);
        if (!Array.isArray(items)) throw new RenderError('INVALID_DATA', `Expected an array at ${row.path}`, sheet.part);
        iterations += items.length;
        if (iterations > 10_000) throw new RenderError('LIMIT_EXCEEDED', 'Spreadsheet expansion exceeds 10000 iterations');
        const values: Scope[] = [];
        for (const item of items) {
          if (!item || typeof item !== 'object' || Array.isArray(item)) throw new RenderError('INVALID_DATA', `Expected object items at ${row.path}`, sheet.part);
          values.push([data, item]);
        }
        scopes.set(row, values); expansions.push({ row: row.number, count: values.length });
      }
      maps.set(sheet.name.toLowerCase(), new RowMap(expansions));
    }
    const parts = new Map(original), set = boundedSet(parts);
    for (const sheet of sheets) {
      const map = maps.get(sheet.name.toLowerCase())!, edits = rangeEdits(sheet.source, sheet.nodes, map, maps, sheet.name);
      let rowBytes = 0;
      for (const row of sheet.rows) {
        const rendered: string[] = [];
        for (const [index, scope] of (scopes.get(row) ?? [[data]]).entries()) {
          const newRow = map.map(row.number)! + index;
          if (newRow > maxRow) throw new RenderError('LIMIT_EXCEEDED', 'Expanded row exceeds Excel bounds', sheet.part);
          const local: Edit[] = []; let inserted = 0;
          const add = (...values: Edit[]) => { for (const value of values) { inserted += Buffer.byteLength(value.value); if (inserted > limits.entry) throw new RenderError('LIMIT_EXCEEDED', 'Row replacements exceed the part byte budget', sheet.part); local.push(value); } };
          add(attributeEdit(row.node, 'r', String(newRow)));
          for (const cell of row.cells) {
            const binding = row.bindings.find(c => c.node === cell);
            if (binding && binding.text.bindings.some(b => b.kind !== 'value') && binding.text.bindings.every(b => b.kind !== 'value') && applyEdits(binding.text.text, binding.text.bindings.map(b => ({ ...b, value: '' }))).trim() === '') {
              add({ start: cell.start, end: cell.end, value: '' }); continue;
            }
            add(attributeEdit(cell, 'r', address(attr(cell, 'r')!).column + newRow));
            if (binding) {
              const result = renderCellText(binding.text, data, sheet.part, scope);
              if (binding.text.whole && (result.value === null || typeof result.value === 'number' || typeof result.value === 'boolean')) {
                add(attributeEdit(cell, 't', typeof result.value === 'boolean' ? 'b' : 'n'));
                add({ start: binding.payload.start, end: binding.payload.end, value: result.value === null ? '' : `<${prefix(cell)}v>${typeof result.value === 'boolean' ? Number(result.value) : result.value}</${prefix(cell)}v>` });
              } else if (binding.shared !== undefined) {
                add(attributeEdit(cell, 't', 'inlineStr'));
                add({ start: binding.payload.start, end: binding.payload.end, value: inlineShared(sharedSource, sharedEntries[binding.shared]!, result) });
              } else add(...result.edits);
            }
            if (!binding && attr(cell, 't') === 's') {
              const value = children(sheet.nodes, cell).find(n => n.uri === ns.spreadsheet && n.local === 'v')!;
              const old = Number(readLeaf(sheet.source, value)), mapped = sharedRemap[old];
              if (mapped === undefined || mapped < 0) throw new RenderError('INVALID_PACKAGE', 'Unresolved shared-string cell', sheet.part);
              if (mapped !== old) add({ start: value.openEnd, end: value.closeStart, value: String(mapped) });
            }
            for (const formula of children(sheet.nodes, cell).filter(n => n.uri === ns.spreadsheet && n.local === 'f')) {
              add({ start: formula.openEnd, end: formula.closeStart, value: escapeXml(rewriteFormula(parseFormula(readLeaf(sheet.source, formula)), maps, sheet.name, row.path ? index : 0, row.path ? row.number : undefined)) });
            }
          }
          const value = applyEdits(sheet.source.slice(row.node.start, row.node.end), local.map(e => ({ ...e, start: e.start - row.node.start, end: e.end - row.node.start })), limits.entry);
          rowBytes += Buffer.byteLength(value);
          if (rowBytes > limits.entry) throw new RenderError('LIMIT_EXCEEDED', 'Expanded rows exceed the part byte budget', sheet.part);
          rendered.push(value);
        }
        edits.push({ start: row.node.start, end: row.node.end, value: rendered.join('') });
      }
      set(sheet.part, applyEdits(sheet.source, edits, limits.entry));
    }
    for (const table of tables) set(table.part, applyEdits(table.source, rangeEdits(table.source, table.nodes, maps.get(table.sheet.toLowerCase())!, maps, table.sheet), limits.entry));
    const names: Edit[] = [];
    for (const node of bookNodes.filter(n => n.uri === ns.spreadsheet && n.local === 'definedName')) {
      const localSheet = attr(node, 'localSheetId'), sheet = localSheet === undefined ? '' : sheets[Number(localSheet)]?.name ?? '';
      names.push({ start: node.openEnd, end: node.closeStart, value: escapeXml(rewriteFormula(parseFormula(readLeaf(workbook, node), true), maps, sheet)) });
    }
    if (names.length) set(mainPart, applyEdits(workbook, names, limits.entry));
    if (sharedPart) {
      const edits: Edit[] = [];
      for (const [index, text] of shared.entries()) if (text.bindings.length) {
        const node = sharedEntries[index]!; edits.push({ start: node.start, end: node.end, value: '' });
      }
      let count = 0;
      for (const sheet of sheets) for (const n of indexXml(parts.get(sheet.part)!.toString(), sheet.part)) if (n.uri === ns.spreadsheet && n.local === 'c' && attr(n, 't') === 's') count++;
      edits.push(attributeEdit(sharedNodes[0]!, 'count', String(count)), attributeEdit(sharedNodes[0]!, 'uniqueCount', String(remainingShared.length)));
      set(sharedPart, applyEdits(sharedSource, edits, limits.entry));
    }
    recalculateCells(parts, sheets, remainingShared, set);
    removeCalculationChain(parts, mainPart, set);
    const drawingEdits = new Map<string, Edit[]>();
    for (const drawing of drawings) {
      const map = maps.get(drawing.sheet.toLowerCase())!, edits: Edit[] = [];
      for (const node of drawing.nodes.filter(n => n.uri === ns.spreadsheetDrawing && n.local === 'row')) {
        const anchor = ancestor(node, ns.spreadsheetDrawing, 'twoCellAnchor');
        if (anchor && attr(anchor, 'editAs') === 'absolute') continue;
        const row = Number(readLeaf(drawing.source, node)) + 1, mapped = map.map(row);
        if (mapped === undefined) throw new RenderError('UNSUPPORTED_TEMPLATE', 'Drawing anchor was removed by row expansion', drawing.part);
        if (mapped !== row) edits.push({ start: node.openEnd, end: node.closeStart, value: String(mapped - 1) });
      }
      drawingEdits.set(drawing.part, edits);
    }
    return { parts, drawingEdits };
  };
}
function recalculateCells(parts: PackageParts, sheets: Sheet[], shared: CellText[], set: (part: string, source: string | Buffer) => void) {
  const cells = new Map<string, { sheet: string; cell: ElementSpan; nodes: ElementSpan[]; source: string; part: string; formula?: ReturnType<typeof parseFormula> }>();
  const perSheet = new Map<string, string>(), known = new Set(sheets.map(s => s.name.toLowerCase()));
  for (const sheet of sheets) {
    const source = parts.get(sheet.part)!.toString(), nodes = indexXml(source, sheet.part); perSheet.set(sheet.part, source);
    for (const cell of nodes.filter(n => n.uri === ns.spreadsheet && n.local === 'c')) {
      const child = children(nodes, cell), formula = child.find(n => n.uri === ns.spreadsheet && n.local === 'f');
      cells.set(`${sheet.name.toLowerCase()}!${attr(cell, 'r')}`, { sheet: sheet.name, cell, nodes: child, source, part: sheet.part, formula: formula ? parseFormula(readLeaf(source, formula)) : undefined });
    }
  }
  const cached = new Map<string, Value>(), active = new Set<string>(), budget = { operations: 0 };
  const read = (sheet: string, reference: string): Value => {
    if (!known.has(sheet.toLowerCase())) return { error: '#REF!' };
    const key = `${sheet.toLowerCase()}!${reference.toUpperCase()}`;
    if (cached.has(key)) return cached.get(key)!;
    if (active.has(key)) throw new RenderError('FORMULA_CYCLE', `Circular formula dependency at ${sheet}!${reference}`);
    if (active.size > 128) throw new RenderError('LIMIT_EXCEEDED', 'Formula dependency depth exceeds 128 levels');
    const cell = cells.get(key); if (!cell) return null;
    active.add(key);
    let value: Value;
    if (cell.formula) value = evaluateFormula(cell.formula, cell.sheet, read, budget);
    else {
      const type = attr(cell.cell, 't'), raw = cell.nodes.find(n => n.local === 'v'), text = raw ? readLeaf(cell.source, raw) : '';
      if (type === 's') value = shared[Number(text)]?.text ?? '';
      else if (type === 'inlineStr') {
        const content = cell.nodes.find(n => n.local === 'is');
        const doc = content ? parseXml(`<wrapper xmlns="${ns.spreadsheet}" ${[...cell.cell.namespaces].filter(([p]) => p && p !== 'xml').map(([p, uri]) => `xmlns:${p}="${escapeXml(uri)}"`).join(' ')}>${cell.source.slice(content.start, content.end)}</wrapper>`, cell.part) : undefined;
        value = doc ? descendants(doc, ns.spreadsheet, 't').map(n => decodeXstring(n.textContent ?? '')).join('') : '';
      } else if (type === 'b') value = text === '1';
      else if (type === 'e') value = { error: text };
      else if (type === 'str' || type === 'd') value = decodeXstring(text);
      else value = raw ? Number(text) : null;
    }
    active.delete(key); cached.set(key, value); return value;
  };
  const edits = new Map<string, Edit[]>();
  for (const [key, cell] of cells) if (cell.formula) {
    const value = read(cell.sheet, attr(cell.cell, 'r')!), list = edits.get(cell.part) ?? [];
    const type = value !== null && typeof value === 'object' ? 'e' : typeof value === 'boolean' ? 'b' : typeof value === 'string' ? 'str' : 'n';
    list.push(attributeEdit(cell.cell, 't', type));
    const xml = `<${prefix(cell.cell)}v>${typeof value === 'object' && value !== null ? value.error : typeof value === 'boolean' ? Number(value) : typeof value === 'string' ? encodeXstring(value) : value ?? 0}</${prefix(cell.cell)}v>`;
    const existing = cell.nodes.find(n => n.local === 'v'), formula = cell.nodes.find(n => n.local === 'f')!;
    list.push({ start: existing?.start ?? formula.end, end: existing?.end ?? formula.end, value: xml }); edits.set(cell.part, list);
  }
  for (const [part, list] of edits) set(part, applyEdits(perSheet.get(part)!, list, limits.entry));
}
function removeCalculationChain(parts: PackageParts, mainPart: string, set: (part: string, source: string | Buffer) => void) {
  const path = relsPath(mainPart), bytes = parts.get(path); if (!bytes) return;
  const document = parseXml(bytes.toString(), path), types = parseXml(parts.get('[Content_Types].xml')!.toString(), '[Content_Types].xml');
  let changed = false;
  for (const link of descendants(document, ns.packageRelationship, 'Relationship')) if (link.getAttribute('Type') === `${ns.relationship}/calcChain`) {
    const part = resolvePart(mainPart, link.getAttribute('Target')!); set(part, Buffer.alloc(0)); parts.delete(part); link.parentNode!.removeChild(link);
    for (const entry of descendants(types, ns.contentTypes, 'Override')) if (entry.getAttribute('PartName') === `/${part}`) entry.parentNode!.removeChild(entry);
    changed = true;
  }
  if (changed) { set(path, serialize(document)); set('[Content_Types].xml', serialize(types)); }
}
