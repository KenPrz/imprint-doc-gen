import { createServer } from 'node:http';
import { timingSafeEqual } from 'node:crypto';
import type { RenderService } from './render-service.ts';
export function workerMetricsServer(
  renderer: RenderService,
  queued: () => number | Promise<number>,
  token: string,
) {
  if (token.length < 24) throw new Error('Worker metrics token must contain at least 24 characters');
  const expected = Buffer.from('Bearer ' + token);
  return createServer(async (request, response) => {
    const actual = Buffer.from(request.headers.authorization ?? '');
    if (actual.length !== expected.length || !timingSafeEqual(actual, expected)) {
      response.writeHead(401).end();
      return;
    }
    if (request.url !== '/metrics' || request.method !== 'GET') {
      response.writeHead(404).end();
      return;
    }
    try {
      const count = await queued();
      response
        .writeHead(200, { 'Content-Type': 'text/plain; version=0.0.4' })
        .end(renderer.metrics.text(renderer.active, renderer.concurrency, count));
    } catch {
      response.writeHead(503).end('Metrics unavailable');
    }
  });
}
