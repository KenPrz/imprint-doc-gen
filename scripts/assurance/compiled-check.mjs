import assert from 'node:assert/strict';
import { createApp } from '../../dist/app.js';
import { encodeTextTemplate } from '../../dist/index.js';
const apiKey = 'compiled-assurance-smoke-key-123456';
const app = createApp({ apiKey });
const call = async (payload) => {
  const response = await app.inject({ method: 'POST', url: '/v1/assurance', headers: { authorization: `Bearer ${apiKey}` }, payload });
  assert.equal(response.statusCode, 200, response.body);
  return response.json();
};
try {
  const template = encodeTextTemplate('txt', 'Hi {d.name}').toString('base64');
  const data = { name: 'compiled' };
  const trace = await call({ operation: 'trace', template, data });
  assert.equal(Buffer.from(trace.artifact, 'base64').toString(), 'Hi compiled');
  assert.equal(trace.trace.events.length, 1);
  const bundle = await call({ operation: 'replay-create', template, data });
  const replay = await call({ operation: 'replay', bundle });
  assert.equal(Buffer.from(replay.artifact, 'base64').toString(), 'Hi compiled');
  console.log('Compiled assurance worker: trace and strict replay passed');
} finally { await app.close(); }
