import type { FastifyInstance, FastifyRequest } from 'fastify';
import type { ServiceStore } from './store-interface.ts';
import { StandaloneStore, type Job, type JobInput } from './store.ts';
import { ServiceError } from './service-error.ts';
import { decodeBase64, RenderService } from './render-service.ts';
import { renderDocumentInWorker } from './jobs.ts';
import { mediaTypes } from './types.ts';

export interface Principal { tenant: string; admin: boolean }
export const inputProperties = {
  restartNumbering: { type: 'boolean' },
  charts: { type: 'object', maxProperties: 32, additionalProperties: { type: 'object', additionalProperties: false, required: ['categories','series'], properties: { categories: { type: 'array', maxItems: 10000, items: { type: 'string', maxLength: 32767 } }, series: { type: 'array', minItems: 1, maxItems: 100, items: { type: 'object', additionalProperties: false, required: ['name','values'], properties: { name: { type: 'string', maxLength: 32767 }, values: { type: 'array', maxItems: 10000, items: { type: 'number' } } } } } } } },
  locale: { type: 'string', maxLength: 80 }, timeZone: { type: 'string', maxLength: 80 },
  images: { type: 'object', maxProperties: 32, additionalProperties: { type: 'object', additionalProperties: false, properties: {
    fit: { type: 'string', enum: ['contain','cover','fill'] }, width: { type: 'integer', minimum: 1, maximum: 4096 }, height: { type: 'integer', minimum: 1, maximum: 4096 }, rotate: { type: 'number', minimum: -360, maximum: 360 }, background: { type: 'string', pattern: '^#[0-9a-fA-F]{6}([0-9a-fA-F]{2})?$' },
    crop: { type: 'object', additionalProperties: false, required: ['left','top','width','height'], properties: { left: { type: 'integer', minimum: 0 }, top: { type: 'integer', minimum: 0 }, width: { type: 'integer', minimum: 1 }, height: { type: 'integer', minimum: 1 } } },
  } } },
  data: { type: 'object', additionalProperties: true },
  attachments: { type: 'object', maxProperties: 32, additionalProperties: { type: 'string', maxLength: 12 * 1024 * 1024 } },
  output: { type: 'string', enum: ['docx', 'xlsx', 'pptx', 'odt', 'ods', 'odp', 'txt', 'csv', 'pdf'] },
};
const id = { type: 'string', pattern: '^[a-zA-Z0-9-]{1,64}$' };
const body = (required: string[], properties: object) => ({ body: { type: 'object', additionalProperties: false, required, properties } });
const publicTemplate = (template: ReturnType<StandaloneStore['template']>) => ({ id: template.id, name: template.name, publishedVersion: template.published_version, createdAt: template.created_at });
const templateBytes = { type: 'string', maxLength: 12 * 1024 * 1024 };
const name = { type: 'string', minLength: 1, maxLength: 120 };
export const publicJob = (job: Job) => ({
  id: job.id, templateId: job.template_id, versionId: job.version_id, status: job.status,
  attempts: job.attempts, createdAt: job.created_at, updatedAt: job.updated_at, deadline: job.deadline,
  error: job.error, format: job.format,
  downloadUrl: job.status === 'succeeded' ? `/v1/jobs/${job.id}/result` : null,
});

export function registerServiceRoutes(app: FastifyInstance, store: ServiceStore, renderer: RenderService, principal: (request: FastifyRequest) => Principal) {
  const tenant = (request: FastifyRequest) => principal(request).tenant;
  const admin = (request: FastifyRequest) => { if (!principal(request).admin) throw new ServiceError(403, 'FORBIDDEN'); };
  app.get('/v1/tenants', async request => { admin(request); return { tenants: await store.listTenants() }; });
  app.post<{ Body: { name: string } }>('/v1/tenants', { schema: body(['name'], { name }) }, async (request, reply) => {
    admin(request); return reply.code(201).send(await store.createTenant(request.body.name));
  });
  app.post<{ Params: { id: string } }>('/v1/tenants/:id/key', async request => { admin(request); return store.rotateKey(request.params.id); });
  app.delete<{ Params: { id: string } }>('/v1/tenants/:id/key', async request => { admin(request); return store.rotateKey(request.params.id, true); });
  app.get('/v1/templates', async request => ({ templates: (await store.listTemplates(tenant(request))).map(publicTemplate) }));
  // Upload validation shares the render admission limit and executes in a bounded worker.
  const upload = async (owner: string, title: string, base64: string, templateId?: string) => {
    if (!renderer.available) throw new ServiceError(503, 'BUSY');
    renderer.active++;
    try {
      if (templateId) await store.template(owner, templateId);
      const bytes = decodeBase64(base64);
      const validated = await renderDocumentInWorker(bytes, {}, {}, 30_000, undefined, true);
      return await store.upload(owner, title, bytes, validated.format, templateId);
    } finally { renderer.active--; }
  };
  app.post<{ Body: { name: string; template: string } }>('/v1/templates', { schema: body(['name', 'template'], { name, template: templateBytes }) }, async (request, reply) => {
    return reply.code(201).send(await upload(tenant(request), request.body.name, request.body.template));
  });
  app.get<{ Params: { id: string } }>('/v1/templates/:id', async request => ({ ...publicTemplate(await store.template(tenant(request), request.params.id)), versions: (await store.versions(tenant(request), request.params.id)).map(version => ({ id: version.id, templateId: version.template_id, format: version.format, createdAt: version.created_at })) }));
  app.post<{ Params: { id: string }; Body: { template: string } }>('/v1/templates/:id/versions', { schema: body(['template'], { template: templateBytes }) }, async (request, reply) => {
    return reply.code(201).send(await upload(tenant(request), '', request.body.template, request.params.id));
  });
  app.post<{ Params: { id: string }; Body: { versionId: string } }>('/v1/templates/:id/publish', { schema: body(['versionId'], { versionId: id }) }, async request => store.publish(tenant(request), request.params.id, request.body.versionId));
  app.delete<{ Params: { id: string } }>('/v1/templates/:id', async (request, reply) => { await store.deleteTemplate(tenant(request), request.params.id); return reply.code(204).send(); });
  app.post<{ Body: JobInput }>('/v1/jobs', { schema: body(['templateId', 'data'], { templateId: id, versionId: id, ...inputProperties }) }, async (request, reply) => {
    const key = request.headers['idempotency-key'];
    if (key !== undefined && (typeof key !== 'string' || !/^[\x21-\x7e]{1,128}$/.test(key))) throw new ServiceError(400, 'INVALID_IDEMPOTENCY_KEY');
    if (request.body.output === 'pdf' && !renderer.converterUrl) throw new ServiceError(502, 'CONVERTER_UNAVAILABLE');
    for (const value of Object.values(request.body.attachments ?? {})) decodeBase64(value);
    const { job, reused } = await store.enqueue(tenant(request), request.body, key);
    return reply.code(reused ? 200 : 202).header('Location', `/v1/jobs/${job.id}`).send(publicJob(job));
  });
  app.get<{ Params: { id: string } }>('/v1/jobs/:id', async request => publicJob(await store.job(tenant(request), request.params.id)));
  app.get<{ Params: { id: string } }>('/v1/jobs/:id/result', async (request, reply) => {
    const result = await store.result(tenant(request), request.params.id);
    return reply.type(result.format === 'pdf' ? 'application/pdf' : mediaTypes[result.format]).header('Content-Disposition', `attachment; filename="document.${result.format}"`).send(result.bytes);
  });
  app.delete<{ Params: { id: string } }>('/v1/jobs/:id', async (request, reply) => { await store.deleteJob(tenant(request), request.params.id); return reply.code(204).send(); });
}
