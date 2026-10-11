import Fastify from 'fastify';
import { timingSafeEqual } from 'node:crypto';
import { RenderError } from './errors.ts';
import { renderInWorker } from './jobs.ts';
import { convertToPdf } from './converter.ts';

interface RenderBody { template: string; data: Record<string, unknown>; attachments?: Record<string, string>; output?: 'docx' | 'pdf' }
export function createApp(config: { apiKey: string; converterUrl?: string; logger?: boolean; concurrency?: number }) {
  if (config.apiKey.length < 24) throw new Error('IMPRINT_API_KEY must contain at least 24 characters');
  const app = Fastify({ bodyLimit: 16 * 1024 * 1024, requestTimeout: 100_000, logger: config.logger ? { redact: ['req.headers.authorization'] } : false });
  const secret = Buffer.from(`Bearer ${config.apiKey}`); let active = 0;
  const decode = (value: string) => {
    if (!value || value.length % 4 || !/^(?:[A-Za-z0-9+/]{4})*(?:[A-Za-z0-9+/]{2}==|[A-Za-z0-9+/]{3}=)?$/.test(value)) throw new RenderError('INVALID_DATA', 'Expected canonical base64 bytes');
    return Buffer.from(value, 'base64');
  };
  app.get('/health', async () => ({ status: 'ok', service: 'imprint-doc-gen', milestone: 1, formats: ['docx'], features: ['text', 'pictures', 'headers', 'footers'] }));
  app.addHook('onRequest', async (request, reply) => {
    if (request.routeOptions.url === '/health') return;
    const supplied = Buffer.from(request.headers.authorization ?? '');
    if (supplied.length !== secret.length || !timingSafeEqual(supplied, secret)) return reply.code(401).send({ error: 'UNAUTHORIZED' });
  });
  app.post<{ Body: RenderBody }>('/v1/render', {
    schema: { body: { type: 'object', additionalProperties: false, required: ['template', 'data'], properties: {
      template: { type: 'string', maxLength: 12 * 1024 * 1024 }, data: { type: 'object', additionalProperties: true },
      attachments: { type: 'object', maxProperties: 32, additionalProperties: { type: 'string', maxLength: 12 * 1024 * 1024 } }, output: { type: 'string', enum: ['docx', 'pdf'] },
    } } },
  }, async (request, reply) => {
    if (active >= (config.concurrency ?? 2)) return reply.code(503).header('Retry-After', '1').send({ error: 'BUSY' });
    active++;
    try {
      const attachments: Record<string, Uint8Array> = Object.create(null);
      for (const [key, value] of Object.entries(request.body.attachments ?? {})) attachments[key] = decode(value);
      const output = await renderInWorker(decode(request.body.template), request.body.data, attachments);
      if (request.body.output === 'pdf') {
        if (!config.converterUrl) throw new RenderError('CONVERTER_UNAVAILABLE', 'PDF conversion is not configured');
        const pdf = await convertToPdf(output, config.converterUrl);
        return reply.type('application/pdf').header('Content-Disposition', 'attachment; filename="document.pdf"').send(pdf);
      }
      return reply.type('application/vnd.openxmlformats-officedocument.wordprocessingml.document').header('Content-Disposition', 'attachment; filename="document.docx"').send(output);
    } catch (error) {
      reply.type('application/json');
      if (error instanceof RenderError) {
        const status = error.code === 'LIMIT_EXCEEDED' ? 413 : error.code === 'RENDER_TIMEOUT' ? 504 : error.code.startsWith('CONVER') ? 502 : 422;
        return reply.code(status).send({ error: error.code, message: error.message, part: error.part });
      }
      request.log.error({ err: error }, 'Render failed');
      return reply.code(500).send({ error: 'RENDER_FAILED' });
    } finally { active--; }
  });
  return app;
}
