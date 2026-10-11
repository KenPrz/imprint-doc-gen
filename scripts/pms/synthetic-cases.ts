import { readFile } from 'node:fs/promises';
import { createHash } from 'node:crypto';
import assert from 'node:assert/strict';
import { scenarios, syntheticData, syntheticInvoice, syntheticAssets, type Data, type Scenario } from './synthetic-data.ts';
export interface SyntheticCase { id: string; report: string; scenario: Scenario; format: 'docx' | 'xlsx'; template: Buffer; source: string; data: Data; attachments: Record<string, Buffer>; expectedImage?: Buffer; imageMode: string }
export async function syntheticCases(seed = 20261011) {
  const assets = await syntheticAssets(), cases: SyntheticCase[] = [];
  const root = new URL('../../tests/fixtures/pms/', import.meta.url);
  const manifest = JSON.parse(await readFile(new URL('manifest.json', root), 'utf8'));
  for (const item of manifest.templates as { file: string; sha256: string }[]) {
    const template = await readFile(new URL(item.file, root));
    assert.equal(createHash('sha256').update(template).digest('hex'), item.sha256, `Original PMS template changed: ${item.file}`);
    const report = item.file.replace('-template.docx', ''), sample = JSON.parse(await readFile(new URL(`${report}.json`, root), 'utf8'));
    for (const scenario of scenarios) {
      const data = syntheticData(report, sample, scenario, seed), attachments: Record<string, Buffer> = {};
      let expectedImage: Buffer | undefined, imageMode = 'no-image-field';
      if (report === 'registration_form') {
        if (scenario === 'empty') imageMode = 'optional-signature-omitted';
        else if (scenario === 'multilingual') { expectedImage = assets['signature.jpg']; data.guest_signature = `data:image/jpeg;base64,${expectedImage!.toString('base64')}`; imageMode = 'jpeg-data-url'; }
        else { expectedImage = assets['signature.png']; const key = scenario === 'large' ? 'signature' : 'guest_signature'; attachments[key] = expectedImage!; if (scenario === 'large') data.guest_signature = key; imageMode = scenario === 'large' ? 'named-png-attachment' : 'direct-png-attachment'; }
      }
      cases.push({ id: `${report}-${scenario}`, report, scenario, format: 'docx', template, source: `tests/fixtures/pms/${item.file}`, data, attachments, expectedImage, imageMode });
    }
  }
  for (const format of ['docx', 'xlsx'] as const) {
    const source = `examples/pms/invoice.${format}`, template = await readFile(source);
    for (const scenario of scenarios) cases.push({ id: `invoice-${format}-${scenario}`, report: `invoice-${format}`, scenario, format, template, source, data: syntheticInvoice(scenario), attachments: { logo: assets['logo.png']! }, expectedImage: assets['logo.png'], imageMode: 'png-logo' });
  }
  return { cases, assets, sourceCommit: manifest.sourceCommit, seed };
}
