import sharp from 'sharp';

export const scenarios = ['normal', 'empty', 'multilingual', 'large'] as const;
export type Scenario = typeof scenarios[number];
export type Data = Record<string, any>;
const money = (value: number) => value.toLocaleString('en-US', { minimumFractionDigits: 2, maximumFractionDigits: 2 });

export function syntheticInvoice(scenario: Scenario, seed = 20261011): Data {
  const count = scenario === 'empty' ? 0 : scenario === 'large' ? 15 : 3;
  const items = Array.from({ length: count }, (_, n) => ({ description: `Synthetic service ${n + 1}`, quantity: n % 3 + 1, price: 500 + n * 25, amount: '' }));
  for (const item of items) item.amount = money(item.quantity * item.price);
  return { company: 'Imprint Test Studio', invoiceNumber: `TEST-INVOICE-${seed}`, customer: scenario === 'multilingual' ? 'Zoë Núñez 山田' : 'Alex Example', orderNumber: `TEST-ORDER-${seed}`, issuedOn: '2026-10-11', dueOn: '2026-10-25', currency: 'USD', total: money(items.reduce((sum, item) => sum + item.quantity * item.price, 0)), items };
}

export async function syntheticAssets(): Promise<Record<string, Buffer>> {
  const signature = `<svg xmlns="http://www.w3.org/2000/svg" width="720" height="180"><path d="M30 100 Q90 5 75 115 Q120 35 150 95 Q175 130 195 50 Q220 135 275 80 Q320 50 350 95 Q400 10 430 105 Q480 125 530 70 L670 90" fill="none" stroke="#163f72" stroke-width="5"/><text x="30" y="165" font-size="25" font-family="sans-serif" fill="#444">SYNTHETIC SIGNATURE - TEST ONLY</text></svg>`;
  const logo = `<svg xmlns="http://www.w3.org/2000/svg" width="480" height="160"><rect width="480" height="160" rx="16" fill="#164b61"/><path d="M30 110 V60 L65 30 L100 60 V110 Z" fill="#e6b957"/><text x="125" y="76" font-family="sans-serif" font-size="35" fill="white">IMPRINT TEST</text><text x="125" y="114" font-family="sans-serif" font-size="22" fill="white">SYNTHETIC STUDIO</text></svg>`;
  return { 'signature.png': await sharp(Buffer.from(signature)).png().toBuffer(), 'signature.jpg': await sharp(Buffer.from(signature)).flatten({ background: '#fff' }).jpeg({ quality: 90 }).toBuffer(), 'logo.png': await sharp(Buffer.from(logo)).png().toBuffer() };
}
