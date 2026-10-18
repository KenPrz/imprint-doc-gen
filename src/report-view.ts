import { createHash } from 'node:crypto';
import { validateData } from './data.ts';
import { render } from './index.ts';
import type { RenderOptions } from './types.ts';
import { RenderError } from './errors.ts';
export type ReportValue = string | number | boolean | null;
export interface ReportColumn {
  key: string;
  label: string;
  type: 'string' | 'number' | 'boolean';
  sum?: boolean;
}
export interface ReportDefinition {
  title: string;
  columns: ReportColumn[];
  rows: Record<string, ReportValue>[];
}
export interface ReportFilter {
  key: string;
  op: 'eq' | 'contains' | 'gte' | 'lte';
  value: ReportValue;
}
export interface ReportSnapshot {
  version: 1;
  title: string;
  columns: ReportColumn[];
  rows: Record<string, ReportValue>[];
  filters: ReportFilter[];
  totals: Record<string, number>;
  sha256: string;
}
function fail(message: string): never {
  throw new RenderError('REPORT_DEFINITION', message);
}
export function createReportSnapshot(
  definition: ReportDefinition,
  filters: ReportFilter[] = [],
): ReportSnapshot {
  validateData({ definition, filters });
  if (Buffer.byteLength(JSON.stringify({ definition, filters })) > 4 * 1024 * 1024)
    fail('Report exceeds 4 MiB');
  if (
    !definition ||
    Object.keys(definition).sort().join(',') !== 'columns,rows,title' ||
    typeof definition.title !== 'string' ||
    definition.title.length > 200
  )
    fail('Invalid report title or properties');
  if (
    !Array.isArray(definition.columns) ||
    !definition.columns.length ||
    definition.columns.length > 50 ||
    !Array.isArray(definition.rows) ||
    definition.rows.length > 10000
  )
    fail('Expected 1–50 columns and at most 10000 rows');
  const columns = structuredClone(definition.columns),
    keys = new Set<string>();
  for (const c of columns) {
    if (
      !c ||
      Object.keys(c).some((k) => !['key', 'label', 'type', 'sum'].includes(k)) ||
      typeof c.key !== 'string' ||
      !/^[A-Za-z][A-Za-z0-9_]{0,63}$/.test(c.key) ||
      keys.has(c.key) ||
      typeof c.label !== 'string' ||
      c.label.length > 200 ||
      !['string', 'number', 'boolean'].includes(c.type) ||
      (c.sum !== undefined && typeof c.sum !== 'boolean') ||
      (c.sum && c.type !== 'number')
    )
      fail('Invalid or duplicate column');
    keys.add(c.key);
  }
  for (const row of definition.rows) {
    if (
      !row ||
      Array.isArray(row) ||
      typeof row !== 'object' ||
      Object.keys(row).length !== columns.length ||
      Object.keys(row).some((k) => !keys.has(k))
    )
      fail('Each row must have exactly the declared columns');
    for (const c of columns)
      if (!Object.hasOwn(row, c.key) || (row[c.key] !== null && typeof row[c.key] !== c.type))
        fail(`Wrong value type for ${c.key}`);
  }
  if (!Array.isArray(filters) || filters.length > 20) fail('At most 20 filters');
  for (const f of filters) {
    if (!f || Object.keys(f).sort().join(',') !== 'key,op,value') fail('Invalid filter');
    const column = columns.find((c) => c.key === f.key);
    if (!column || !['eq', 'contains', 'gte', 'lte'].includes(f.op))
      fail('Unknown filter column or operator');
    if (f.value !== null && typeof f.value !== column.type) fail('Filter type mismatch');
    if (
      (f.op === 'contains' && (column.type !== 'string' || typeof f.value !== 'string')) ||
      (['gte', 'lte'].includes(f.op) && (column.type !== 'number' || typeof f.value !== 'number'))
    )
      fail('Operator is incompatible with column');
  }
  const rows = definition.rows
    .filter((row) =>
      filters.every((f) => {
        const value = row[f.key];
        if (f.op === 'eq') return value === f.value;
        if (value === null) return false;
        if (f.op === 'contains') return (value as string).includes(f.value as string);
        return f.op === 'gte'
          ? (value as number) >= (f.value as number)
          : (value as number) <= (f.value as number);
      }),
    )
    .map((row) => Object.fromEntries(columns.map((c) => [c.key, row[c.key]])));
  const totals: Record<string, number> = {};
  for (const c of columns.filter((c) => c.sum)) {
    let sum = 0,
      correction = 0;
    for (const row of rows) {
      const next = (row[c.key] as number | null) ?? 0;
      const adjusted = next - correction;
      const result = sum + adjusted;
      correction = result - sum - adjusted;
      sum = result;
    }
    if (!Number.isFinite(sum)) fail('Aggregate exceeds finite number range');
    totals[c.key] = sum;
  }
  const content = {
    version: 1 as const,
    title: definition.title,
    columns,
    rows,
    filters: structuredClone(filters),
    totals,
  };
  return { ...content, sha256: createHash('sha256').update(JSON.stringify(content)).digest('hex') };
}
/** Recompute all content: never trust caller-provided totals or a stale digest. */
export function verifyReportSnapshot(snapshot: ReportSnapshot): ReportSnapshot {
  validateData({ snapshot });
  if (
    !snapshot ||
    snapshot.version !== 1 ||
    Object.keys(snapshot).sort().join(',') !== 'columns,filters,rows,sha256,title,totals,version'
  )
    fail('Invalid snapshot');
  const expected = createReportSnapshot(
    { title: snapshot.title, columns: snapshot.columns, rows: snapshot.rows },
    snapshot.filters,
  );
  if (
    expected.sha256 !== snapshot.sha256 ||
    JSON.stringify(expected.totals) !== JSON.stringify(snapshot.totals)
  )
    fail('Snapshot integrity check failed');
  return expected;
}
export async function renderReportSnapshot(
  template: Uint8Array,
  snapshot: ReportSnapshot,
  options?: RenderOptions,
): Promise<Buffer> {
  const verified = verifyReportSnapshot(snapshot);
  return render(template, { report: verified }, options);
}
const escape = (s: string) =>
  s
    .replaceAll('&', '&amp;')
    .replaceAll('<', '&lt;')
    .replaceAll('>', '&gt;')
    .replaceAll('"', '&quot;')
    .replaceAll("'", '&#39;');
/** Self-contained view contains only the supplied snapshot. Do not use filtering as access control. */
export function reportSnapshotHtml(input: ReportSnapshot): string {
  const snapshot = verifyReportSnapshot(input);
  const encoded = JSON.stringify(snapshot)
    .replaceAll('<', '\\u003c')
    .replaceAll('>', '\\u003e')
    .replaceAll('&', '\\u0026')
    .replaceAll('\u2028', '\\u2028')
    .replaceAll('\u2029', '\\u2029');
  const script = `const s=JSON.parse(document.getElementById('snapshot').textContent);let selected=s.rows;
const body=document.querySelector('tbody'),summary=document.getElementById('summary'),detail=document.getElementById('detail');
function draw(){body.replaceChildren();selected.forEach((r,i)=>{const tr=document.createElement('tr');s.columns.forEach(c=>{const td=document.createElement('td');td.textContent=r[c.key]===null?'':String(r[c.key]);tr.append(td)});const td=document.createElement('td'),b=document.createElement('button');b.textContent='Details';b.setAttribute('aria-label','Details for row '+(i+1));b.onclick=()=>{detail.textContent=JSON.stringify(r,null,2)};td.append(b);tr.append(td);body.append(tr)});summary.textContent=selected.length+' of '+s.rows.length+' rows';const totals=s.columns.filter(c=>c.sum).map(c=>{let total=0,correction=0;for(const row of selected){const adjusted=(row[c.key]??0)-correction,next=total+adjusted;correction=(next-total)-adjusted;total=next}return c.label+': '+(Number.isFinite(total)?String(total):'Overflow')});document.getElementById('totals').textContent=totals.join(' · ');}
document.getElementById('search').oninput=e=>{const q=e.target.value;selected=s.rows.filter(r=>s.columns.some(c=>String(r[c.key]??'').includes(q)));detail.textContent='';draw()};
function download(text,type,name){const url=URL.createObjectURL(new Blob([text],{type}));const a=document.createElement('a');a.href=url;a.download=name;a.click();setTimeout(()=>URL.revokeObjectURL(url),1000)}
document.getElementById('csv').onclick=()=>{function quote(v){let x=String(v??'');if(/^[\\s\\u0000-\\u001f]*[=+@-]/.test(x))x="'"+x;return '"'+x.replaceAll('"','""')+'"'}download([s.columns.map(c=>quote(c.label)).join(','),...selected.map(r=>s.columns.map(c=>quote(r[c.key])).join(','))].join('\\r\\n'),'text/csv;charset=utf-8','report.csv')};
document.getElementById('json').onclick=()=>download(JSON.stringify({title:s.title,columns:s.columns,rows:selected},null,2),'application/json','report-selection.json');
document.getElementById('print').onclick=()=>window.print();draw();`;
  const scriptHash = createHash('sha256').update(script).digest('base64');
  return `<!doctype html><html lang="en"><head><meta charset="utf-8"><meta name="viewport" content="width=device-width,initial-scale=1"><meta http-equiv="Content-Security-Policy" content="default-src 'none'; script-src 'sha256-${scriptHash}'; style-src 'unsafe-inline'; img-src data:; base-uri 'none'; form-action 'none'"><title>${escape(snapshot.title)}</title><style>body{font:16px system-ui;margin:2rem;color:#162230}table{border-collapse:collapse;width:100%}th,td{padding:.6rem;border:1px solid #bbb;text-align:left}input,button{font:inherit;margin:.3rem;padding:.5rem}.table{overflow:auto}pre{white-space:pre-wrap}p,pre{overflow-wrap:anywhere}.controls input{box-sizing:border-box;max-width:100%} @media print{.controls,td:last-child,th:last-child,#detail{display:none}body{margin:0}}</style></head><body><h1>${escape(snapshot.title)}</h1><div class="controls"><label>Filter visible rows <input id="search" type="search"></label><button id="csv">Download selected CSV</button><button id="json">Download selected data</button><button id="print">Print / save PDF</button></div><p id="summary" aria-live="polite"></p><p id="totals" aria-live="polite"></p><div class="table"><table><thead><tr>${snapshot.columns.map((c) => `<th scope="col">${escape(c.label)}</th>`).join('')}<th scope="col">Inspect</th></tr></thead><tbody></tbody></table></div><pre id="detail" aria-live="polite"></pre><p>Source snapshot: ${snapshot.sha256}</p><script id="snapshot" type="application/json">${encoded}</script><script>${script}</script></body></html>`;
}
