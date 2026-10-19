import test from 'node:test';
import assert from 'node:assert/strict';
import { readFile } from 'node:fs/promises';
import { executeAssurance } from '../src/assurance-task.ts';
import { snapshotDocument, artifactDigest } from '../src/index.ts';
test('JSON assurance assembles pinned real DOCX components and rejects tampered pins and extra properties', async () => {
  const template = await readFile(new URL('./fixtures/docx-components/host.docx', import.meta.url));
  const document = await readFile(new URL('./fixtures/docx-components/clause.docx', import.meta.url));
  const payload = {
    operation: 'docx-components',
    template: template.toString('base64'),
    components: [
      { id: 'clause', version: '1', document: document.toString('base64'), sha256: artifactDigest(document) },
    ],
  };
  const result = (await executeAssurance(payload)) as { artifact: string; lock: { occurrences: number }[] };
  assert.equal(result.lock[0]!.occurrences, 2);
  const snapshot = await snapshotDocument(Buffer.from(result.artifact, 'base64'), 'docx');
  assert.equal(snapshot.units.filter((u) => u.value.includes('First requirement')).length, 2);
  await assert.rejects(
    executeAssurance({ ...payload, components: [{ ...payload.components[0], sha256: '0'.repeat(64) }] }),
  );
  await assert.rejects(executeAssurance({ ...payload, command: 'arbitrary' }));
});
test('JSON Crystal inspection rejects non-CFB artifacts and caller paths', async () => {
  await assert.rejects(
    executeAssurance({ operation: 'crystal-artifact', artifact: Buffer.from('not rpt').toString('base64') }),
  );
  await assert.rejects(
    executeAssurance({ operation: 'crystal-artifact', artifact: '', path: '/tmp/report.rpt' }),
  );
});
