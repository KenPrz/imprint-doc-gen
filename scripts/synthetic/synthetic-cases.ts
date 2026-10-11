import { readFile } from 'node:fs/promises';
import { readPackage, writePackage } from '../../src/package.ts';
import { scenarios, syntheticInvoice, syntheticAssets, type Data, type Scenario } from './synthetic-data.ts';
export interface SyntheticCase { id: string; report: string; scenario: Scenario; format: 'docx' | 'xlsx'; template: Buffer; source: string; data: Data; attachments: Record<string, Buffer>; expectedImage?: Buffer; imageMode: string }
export async function syntheticCases(seed = 20261011) {
  const assets = await syntheticAssets(), cases: SyntheticCase[] = [];
  for (const format of ['docx', 'xlsx'] as const) {
    const source = `examples/invoice/invoice.${format}`, template = await readFile(source);
    for (const scenario of scenarios) cases.push({ id: `invoice-${format}-${scenario}`, report: `invoice-${format}`, scenario, format, template, source, data: syntheticInvoice(scenario, seed), attachments: { logo: assets['logo.png']! }, expectedImage: assets['logo.png'], imageMode: 'png-logo' });
  }
  // A generated invoice variant isolates optional signature binding behavior.
  const parts = await readPackage(cases[0]!.template);
  parts.set('word/document.xml', Buffer.from(parts.get('word/document.xml')!.toString().replaceAll('{d.logo}', '{d.signature}')));
  const template = await writePackage(parts);
  for (const scenario of scenarios) {
    const data = syntheticInvoice(scenario, seed), attachments: Record<string, Buffer> = {};
    let expectedImage: Buffer | undefined, imageMode: string;
    if (scenario === 'empty') imageMode = 'optional-signature-omitted';
    else if (scenario === 'multilingual') { expectedImage = assets['signature.jpg']; data.signature = `data:image/jpeg;base64,${expectedImage!.toString('base64')}`; imageMode = 'jpeg-data-url'; }
    else { expectedImage = assets['signature.png']; const key = scenario === 'large' ? 'approval' : 'signature'; attachments[key] = expectedImage!; if (scenario === 'large') data.signature = key; imageMode = scenario === 'large' ? 'named-png-attachment' : 'direct-png-attachment'; }
    cases.push({ id: `signature-docx-${scenario}`, report: 'signature-docx', scenario, format: 'docx', template, source: 'generated:signature-invoice', data, attachments, expectedImage, imageMode });
  }
  return { cases, assets, seed };
}
