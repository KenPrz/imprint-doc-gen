import type { FastifyInstance } from 'fastify';
import type { RenderService } from './render-service.ts';
import { ServiceError } from './service-error.ts';
import { runTask } from './task-runner.ts';

/** Registered under the application's authentication hook; operations run in bounded workers. */
export function registerAssuranceRoutes(app: FastifyInstance, renderer: RenderService) {
  app.post('/v1/assurance', { bodyLimit: 12 * 1024 * 1024 }, async (request) => {
    if (!renderer.available) throw new ServiceError(503, 'BUSY');
    renderer.active++;
    try {
      return await runTask('assurance', request.body);
    } finally {
      renderer.active--;
    }
  });
}
