import sharp from 'sharp';

export const scenarios = ['normal', 'empty', 'multilingual', 'large'] as const;
export type Scenario = typeof scenarios[number];
export type Data = Record<string, any>;
const money = (value: number) => value.toLocaleString('en-US', { minimumFractionDigits: 2, maximumFractionDigits: 2 });
const amount = (value: unknown) => Number(String(value ?? '0').replace(/[^\d.-]/g, '')) || 0;
const sum = (rows: Data[], key: string) => rows.reduce((total, row) => total + amount(row[key]), 0);
function seedNumber(text: string) { let value = 2166136261; for (const c of text) value = Math.imul(value ^ c.charCodeAt(0), 16777619); return value >>> 0; }

export function syntheticData(report: string, sample: Data, scenario: Scenario, seed = 20261011): Data {
  const id = seedNumber(`${seed}:${report}`) % 900 + 100;
  const names = scenario === 'multilingual' ? ['Zoë Núñez', '山田 花子', 'François Müller', 'Αλέξανδρος Test'] : ['Alex Example', 'Morgan Sample', 'Casey Demo', 'Taylor Fixture'];
  const count = scenario === 'empty' ? 0 : scenario === 'large' ? 15 : 3;
  function visit(value: any, key = '', index = 0, depth = 0): any {
    if (Array.isArray(value)) {
      const length = key === 'sheets' ? Math.min(count, 4) : key === 'rooms' && depth >= 1 ? (scenario === 'large' ? [0, 1, 6, 12][index % 4] : index + 1) : count;
      return value.length ? Array.from({ length }, (_, n) => visit(value[n % value.length], key, n, depth + 1)) : [];
    }
    if (value && typeof value === 'object') return Object.fromEntries(Object.entries(value).map(([k, v]) => [k, visit(v, k, index, depth)]));
    if (typeof value !== 'string') return scenario === 'empty' && typeof value === 'number' ? 0 : value;
    if (key === 'property_name') return `Imprint Test Hotel - SYNTHETIC ${id}`;
    if (key === 'property_code') return 'TEST';
    if (key === 'rendered_on') return `11 Oct 2026, ${sample.report_time ?? sample.current_time ?? '14:32'}`;
    if (key === 'first_name' || key === 'comp_f_name') return names[index % names.length]!.split(' ')[0];
    if (key === 'last_name' || key === 'comp_l_name') return names[index % names.length]!.split(' ').slice(1).join(' ');
    if (/email|display_contact/.test(key)) return `guest-${id}-${index}@imprint.example`;
    if (/phone|contact_no/.test(key)) return `+1 202 555 ${String(100 + index).padStart(4, '0')}`;
    if (/company|organization|account_name/.test(key)) return `Example Travel ${id} & Co.`;
    if (/^(guest|guest_name|name|recipient_name|attendant|finder|emergency_contact_name|window_name)$/.test(key)) return `${names[index % names.length]} ${index + 1}`;
    if (/address|street_name/.test(key)) return scenario === 'multilingual' ? '12 Rue Démo & Example Avenue\nTest District' : '100 Example Avenue, Test District';
    if (/^(city|guest_city|recipient_city|property_city)$/.test(key)) return 'Sample City';
    if (/^(id_no|comp_id_no|reservation_id|confirmation_no|receipt_no|folio_no|bill_no|invoice_no|invoice_number|account_number|statement_number|reference_number)$/.test(key)) return `TEST-${id}-${String(index + 1).padStart(4, '0')}`;
    if (/^(room|room_no|room_number)$/.test(key)) return String(300 + index + 1);
    if (/^(remarks|instructions|body)$/.test(key)) return scenario === 'large' ? 'SYNTHETIC NOTE: verify wrapping and pagination. '.repeat(5) : scenario === 'multilingual' ? 'Café & tea <demo>\nSecond line\tTabbed note — Ω' : 'Synthetic test note; no real guest information.';
    if (/^\d{2}\/\d{2}\/\d{2}$/.test(value)) return /departure|(?:^|_)to$|due/.test(key) ? '14/10/26' : '11/10/26';
    if (/^2026-\d{2}-\d{2}$/.test(value)) return '2026-10-11';
    if (/^[A-Z][a-z]+ \d{1,2}, 2026$/.test(value)) return /departure|(?:^|_)to$|due/.test(key) ? 'October 14, 2026' : 'October 11, 2026';
    if (/^(row_count|total_rows|total_rules|total_sheets)$/.test(key)) return String(count);
    if (scenario === 'empty' && /^(?:(?:PHP|USD) )?-?[\d,]+\.\d{2}$/.test(value)) return value.replace(/-?[\d,]+\.\d{2}$/, '0.00');
    return value;
  }
  const data = visit(sample) as Data;
  const rows = (key: string): Data[] => data[key] ?? [];
  const total = (key: string, source: string, field: string) => { data[key] = money(sum(rows(source), field)); };
  if (report === 'registration_form') data.res_adr = data.room_rate = scenario === 'empty' ? '0.00' : '6,500.00';
  if (report === 'folio_statement') {
    data.windows = count ? [{ window_number: '1', window_name: names[0], window_balance: '0.00' }] : [];
    data.transactions = Array.from({ length: count }, (_, n) => ({ window_number: '1', date: '11/10/26', code: '1000', description: `Synthetic stay charge ${n + 1}`, reference: `TEST-${id}-${n + 1}`, charges: money(1250 + n * 25), payments: '', amount_raw: (1250 + n * 25).toFixed(2), currency: 'PHP' }));
    const charges = sum(data.transactions, 'charges');
    if (count) data.transactions.push({ window_number: '1', date: '14/10/26', code: '9000', description: 'Synthetic settlement', reference: 'TEST-PAYMENT', charges: '', payments: money(charges), amount_raw: (-charges).toFixed(2), currency: 'PHP' });
    data.total_charges = data.total_payments = money(charges); data.balance = data.balance_raw = '0.00';
  }
  if (report === 'payment_receipt' || report === 'deposit_receipt') {
    const key = report === 'payment_receipt' ? 'payments' : 'deposits';
    rows(key).forEach((row, n) => { row.amount_raw = 1250 + n * 100; row.amount = money(row.amount_raw); });
    total('total', key, 'amount'); data.total_raw = sum(rows(key), 'amount');
    if (report === 'payment_receipt') { data.tendered_currency = 'PHP'; data.tendered_amount = data.total; data.exchange_rate = '1.0000'; data.is_foreign_tender = 'N'; }
  }
  if (report === 'ar_settlements') { total('total', 'rows', 'amount'); data.total_raw = sum(rows('rows'), 'amount'); }
  if (report === 'cash_report') { const income = rows('rows').filter(r => amount(r.amount) >= 0), expense = rows('rows').filter(r => amount(r.amount) < 0); data.total_income_local = money(sum(income, 'amount')); data.total_expense_local = money(-sum(expense, 'amount')); data.balance_amount_local_raw = amount(data.starting_amount_local) + sum(rows('rows'), 'amount'); data.balance_amount_local = money(data.balance_amount_local_raw); }
  if (report === 'check_report') { data.total_checks_local_raw = sum(rows('rows'), 'credit') - sum(rows('rows'), 'debit'); data.total_checks_local = money(data.total_checks_local_raw); }
  if (report === 'revenue_summary') { total('total_mtd', 'rows', 'mtd'); total('total_ytd', 'rows', 'ytd'); }
  if (report === 'cashier_closure') for (const [target, source] of [['charges_total', 'charge_rows'], ['settlements_total', 'settlement_rows'], ['paid_outs_total', 'paid_out_rows']]) total(target!, source!, 'amount');
  if (report === 'trial_balance') {
    for (const key of ['revenue', 'paid_outs', 'non_revenue', 'payment']) total(`${key}_total`, key, 'amount');
    const activity = ['revenue', 'paid_outs', 'non_revenue', 'payment'].reduce((n, key) => n + sum(rows(key), 'amount'), 0);
    data.transaction_total_today = money(activity); data.grand_total = data.balance_carried_forward = money(amount(data.balance_brought_forward) + activity);
  }
  if (report === 'task_sheet') for (const [n, sheet] of rows('sheets').entries()) { sheet.sheet_number = String(n + 1); sheet.room_count = String(sheet.rooms.length); sheet.total_credits = String(sum(sheet.rooms, 'credits')); }
  return data;
}

export function syntheticInvoice(scenario: Scenario): Data {
  const count = scenario === 'empty' ? 0 : scenario === 'large' ? 15 : 3;
  const items = Array.from({ length: count }, (_, n) => ({ description: `Synthetic service ${n + 1}`, quantity: n % 3 + 1, price: 500 + n * 25, amount: '' }));
  for (const item of items) item.amount = money(item.quantity * item.price);
  return { hotel: 'Imprint Test Hotel - SYNTHETIC', invoiceNumber: 'TEST-INVOICE-001', guest: scenario === 'multilingual' ? 'Zoë Núñez 山田' : 'Alex Example', room: '301', arrival: '2026-10-11', departure: '2026-10-14', currency: 'PHP', total: money(sum(items, 'amount')), items };
}

export async function syntheticAssets(): Promise<Record<string, Buffer>> {
  const signature = `<svg xmlns="http://www.w3.org/2000/svg" width="720" height="180"><path d="M30 100 Q90 5 75 115 Q120 35 150 95 Q175 130 195 50 Q220 135 275 80 Q320 50 350 95 Q400 10 430 105 Q480 125 530 70 L670 90" fill="none" stroke="#163f72" stroke-width="5"/><text x="30" y="165" font-size="25" font-family="sans-serif" fill="#444">SYNTHETIC SIGNATURE - TEST ONLY</text></svg>`;
  const logo = `<svg xmlns="http://www.w3.org/2000/svg" width="480" height="160"><rect width="480" height="160" rx="16" fill="#164b61"/><path d="M30 110 V60 L65 30 L100 60 V110 Z" fill="#e6b957"/><text x="125" y="76" font-family="sans-serif" font-size="35" fill="white">IMPRINT TEST</text><text x="125" y="114" font-family="sans-serif" font-size="22" fill="white">SYNTHETIC HOTEL</text></svg>`;
  return { 'signature.png': await sharp(Buffer.from(signature)).png().toBuffer(), 'signature.jpg': await sharp(Buffer.from(signature)).flatten({ background: '#fff' }).jpeg({ quality: 90 }).toBuffer(), 'logo.png': await sharp(Buffer.from(logo)).png().toBuffer() };
}
