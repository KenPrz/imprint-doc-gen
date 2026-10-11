import Fastify from 'fastify';
import { timingSafeEqual } from 'node:crypto';
import { RenderError } from './errors.ts';
import { mediaTypes } from './types.ts';
import { StandaloneStore } from './store.ts';
import { DurableQueue } from './queue.ts';
import { ServiceError } from './service-error.ts';
import { decodeBase64, RenderService, type RenderInput } from './render-service.ts';
import { registerServiceRoutes, inputProperties, type Principal } from './service-routes.ts';

interface AppConfig { apiKey: string; converterUrl?: string; logger?: boolean; concurrency?: number; dataDir?: string; runQueue?: boolean }
export function createApp(config: AppConfig) {
  if (config.apiKey.length < 24) throw new Error('IMPRINT_API_KEY must contain at least 24 characters');
  if (config.concurrency !== undefined && (!Number.isSafeInteger(config.concurrency) || config.concurrency < 1 || config.concurrency > 8)) throw new Error('Concurrency must be between 1 and 8');
  const app = Fastify({ bodyLimit: 16 * 1024 * 1024, requestTimeout: 100_000, logger: config.logger ? { redact: ['req.headers.authorization'] } : false });
  const secret = Buffer.from(`Bearer ${config.apiKey}`);
  const renderer = new RenderService(config.concurrency, config.converterUrl);
  const store = config.dataDir ? new StandaloneStore(config.dataDir) : undefined;
  const identities = new WeakMap<object, Principal>();
  app.addHook('onSend', async (_request, reply) => { reply.header('Cache-Control', 'no-store'); });
  app.setErrorHandler((error, request, reply) => {
    reply.type('application/json');
    if (error instanceof ServiceError) {
      if (error.status === 503) reply.header('Retry-After', '1');
      return reply.code(error.status).send({ error: error.code, message: error.message });
    }
    if (error instanceof RenderError) {
      const status = error.code === 'LIMIT_EXCEEDED' ? 413 : error.code === 'RENDER_TIMEOUT' ? 504 : error.code.startsWith('CONVER') ? 502 : 422;
      return reply.code(status).send({ error: error.code, message: error.message, part: error.part });
    }
    if (error instanceof Error && 'statusCode' in error && typeof error.statusCode === 'number' && error.statusCode >= 400 && error.statusCode < 500) return reply.code(error.statusCode).send({ error: 'code' in error ? error.code : 'INVALID_REQUEST' });
    request.log.error({ err: error }, 'Request failed');
    return reply.code(500).send({ error: 'INTERNAL_ERROR' });
  });
  app.get('/health', async () => ({ status: 'ok', service: 'imprint-doc-gen', milestone: 7, storage: store ? 'sqlite' : 'disabled', formats: ['docx', 'xlsx'], features: ['text', 'pictures', 'headers', 'footers', 'loops', 'nested-scopes', 'typed-cells', 'spreadsheet-pictures', 'spreadsheet-rows', 'formula-calculation', 'standard-docx-tags', 'standard-xlsx-tags', ...(store ? ['tenants', 'template-versions', 'durable-jobs'] : [])] }));
  app.get('/ready', async (_request, reply) => {
    if (store) store.db.prepare('SELECT 1').get();
    if (config.converterUrl) {
      try {
        const response = await fetch(new URL('/health', config.converterUrl), { signal: AbortSignal.timeout(2000) });
        await response.body?.cancel();
        if (!response.ok) throw new Error('Converter not ready');
      } catch { return reply.code(503).send({ status: 'unavailable' }); }
    }
    return { status: 'ready' };
  });
  app.addHook('onRequest', async (request, reply) => {
    if (['/health', '/ready'].includes(request.routeOptions.url ?? '')) return;
    const supplied = Buffer.from(request.headers.authorization ?? '');
    if (supplied.length === secret.length && timingSafeEqual(supplied, secret)) { identities.set(request, { tenant: 'default', admin: true }); return; }
    const authorization = request.headers.authorization ?? '';
    const tenant = authorization.startsWith('Bearer ') ? store?.authenticate(authorization.slice(7)) : undefined;
    if (!tenant) return reply.code(401).send({ error: 'UNAUTHORIZED' });
    identities.set(request, { tenant, admin: false });
  });
  app.post<{ Body: RenderInput & { template: string } }>('/v1/render', {
    schema: { body: { type: 'object', additionalProperties: false, required: ['template', 'data'], properties: {
      template: { type: 'string', maxLength: 12 * 1024 * 1024 }, ...inputProperties,
    } } },
  }, async (request, reply) => {
    const output = await renderer.render(decodeBase64(request.body.template), request.body);
    return reply.type(output.format === 'pdf' ? 'application/pdf' : mediaTypes[output.format]).header('Content-Disposition', `attachment; filename="document.${output.format}"`).send(output.bytes);
  });
  // One-shot document rendering shares authentication
  // and resource limits; only the request/response protocol is adapted here.
  app.post<{ Body: RenderInput & { template: string; convertTo?: 'docx' | 'xlsx' | 'pdf' }; Querystring: { download: string } }>('/render/template', {
    schema: {
      querystring: { type: 'object', required: ['download'], additionalProperties: false, properties: { download: { type: 'string', const: 'true' } } },
      body: { type: 'object', additionalProperties: false, required: ['template', 'data'], properties: {
        template: { type: 'string', maxLength: 12 * 1024 * 1024 }, data: inputProperties.data, attachments: inputProperties.attachments, convertTo: inputProperties.output,
      } },
    },
  }, async (request, reply) => {
    const output = await renderer.render(decodeBase64(request.body.template), { data: request.body.data, attachments: request.body.attachments, output: request.body.convertTo });
    return reply.type(output.format === 'pdf' ? 'application/pdf' : mediaTypes[output.format]).header('Content-Disposition', `attachment; filename="document.${output.format}"`).send(output.bytes);
  });
  if (store) {
    registerServiceRoutes(app, store, renderer, request => identities.get(request)!);
    const queue = new DurableQueue(store, renderer, error => app.log.error({ err: error }, 'Queue failed'));
    if (config.runQueue !== false) app.addHook('onReady', async () => queue.start());
    app.addHook('onClose', async () => { await queue.stop(); store.close(); });
  }
  return app;
}
