import test from 'node:test';
import assert from 'node:assert/strict';
import { createServer, type RequestListener } from 'node:http';
import { convertToPdf } from '../src/converter.ts';
import { RenderError } from '../src/errors.ts';

async function withServer(listener: RequestListener, run: (endpoint: string) => Promise<void>) {
  const server = createServer(listener);
  await new Promise<void>(resolve => server.listen(0, '127.0.0.1', resolve));
  try { await run(`http://127.0.0.1:${(server.address() as { port: number }).port}`); }
  finally { server.closeAllConnections(); await new Promise<void>(resolve => server.close(() => resolve())); }
}
const code = (value: string) => (error: unknown) => error instanceof RenderError && error.code === value;

test('converter sends multipart bytes and returns a PDF', async () => {
  const input = Buffer.from('fixture-document-bytes');
  await withServer(async (request, response) => {
    assert.equal(request.url, '/forms/libreoffice/convert');
    assert.match(request.headers['content-type']!, /multipart\/form-data; boundary=/);
    const chunks = []; for await (const chunk of request) chunks.push(chunk);
    assert(Buffer.concat(chunks).includes(input));
    response.end('%PDF-1.7\nfixture');
  }, async endpoint => assert.equal((await convertToPdf(input, endpoint)).toString(), '%PDF-1.7\nfixture'));
});

test('converter distinguishes Gotenberg deadlines from capacity failures', async () => {
  let deadline = true;
  await withServer((_request, response) => {
    response.writeHead(503); response.end(deadline ? 'The request exceeded the time limit. Increase it with --api-timeout, or reduce the workload.' : 'Maximum queue size exceeded');
  }, async endpoint => {
    await assert.rejects(convertToPdf(Buffer.from('test'), endpoint), code('RENDER_TIMEOUT'));
    deadline = false;
    await assert.rejects(convertToPdf(Buffer.from('test'), endpoint), code('CONVERSION_FAILED'));
  });
});

for (const stage of ['headers', 'body']) test(`converter deadline covers stalled ${stage} and the next request succeeds`, async () => {
  let stall = true;
  await withServer((_request, response) => {
    if (stall) { if (stage === 'body') { response.writeHead(200); response.write('%PDF-'); } return; }
    response.end('%PDF-1.7\nrecovered');
  }, async endpoint => {
    await assert.rejects(convertToPdf(Buffer.from('test'), endpoint, 'test.docx', { timeoutMs: 100 }), code('RENDER_TIMEOUT'));
    stall = false;
    assert.equal((await convertToPdf(Buffer.from('test'), endpoint)).toString(), '%PDF-1.7\nrecovered');
  });
});

test('converter rejects truncated PDF streams with a typed error', async () => {
  await withServer((_request, response) => {
    response.writeHead(200, { 'content-length': '500' }); response.write('%PDF-');
    setImmediate(() => response.destroy());
  }, async endpoint => { await assert.rejects(convertToPdf(Buffer.from('test'), endpoint), code('CONVERSION_FAILED')); });
});

test('converter bounds output and can convert again after rejecting it', async () => {
  await withServer((_request, response) => response.end('%PDF-' + 'x'.repeat(1000)), async endpoint => {
    await assert.rejects(convertToPdf(Buffer.from('test'), endpoint, 'test.docx', { maxOutputBytes: 100 }), code('LIMIT_EXCEEDED'));
    assert.equal((await convertToPdf(Buffer.from('test'), endpoint)).length, 1005);
  });
});

test('converter cancels non-success bodies without waiting for a stalled body', async () => {
  await withServer((_request, response) => { response.writeHead(500); response.write('failure'); }, async endpoint => {
    await assert.rejects(convertToPdf(Buffer.from('test'), endpoint, 'test.docx', { timeoutMs: 1000 }), code('CONVERSION_FAILED'));
  });
});

test('converter rejects non-PDF success responses', async () => {
  await withServer((_request, response) => response.end('<html>failure</html>'), async endpoint => {
    await assert.rejects(convertToPdf(Buffer.from('test'), endpoint), code('CONVERSION_FAILED'));
  });
});
