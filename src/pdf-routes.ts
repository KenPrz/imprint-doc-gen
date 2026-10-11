import type { FastifyInstance } from 'fastify';
import { decodeBase64, RenderService } from './render-service.ts';
import { runTask } from './task-runner.ts';
import { ServiceError } from './service-error.ts';
import type { PdfOptions } from './pdf.ts';
import type { BarcodeOptions } from './barcode.ts';
const bytes = { type: 'string', maxLength: 12 * 1024 * 1024 };
const pages = {
  type: 'array',
  minItems: 1,
  maxItems: 1000,
  items: { type: 'integer', minimum: 1, maximum: 1000 },
};
export function registerPdfRoutes(app: FastifyInstance, renderer: RenderService) {
  const execute = async <T>(task: 'pdf' | 'barcode', payload: unknown) => {
    if (!renderer.available) throw new ServiceError(503, 'BUSY');
    renderer.active++;
    try {
      return await runTask<T>(task, payload);
    } finally {
      renderer.active--;
    }
  };
  app.post<{
    Body: Omit<PdfOptions, 'inputs' | 'attachments'> & {
      inputs: { bytes: string; pages?: number[] }[];
      attachments?: { name: string; bytes: string; mimeType?: string }[];
    };
  }>(
    '/v1/pdf',
    {
      schema: {
        body: {
          type: 'object',
          additionalProperties: false,
          required: ['inputs'],
          properties: {
            inputs: {
              type: 'array',
              minItems: 1,
              maxItems: 20,
              items: {
                type: 'object',
                additionalProperties: false,
                required: ['bytes'],
                properties: { bytes, pages },
              },
            },
            watermark: {
              type: 'object',
              additionalProperties: false,
              required: ['text'],
              properties: {
                text: { type: 'string', minLength: 1, maxLength: 200 },
                opacity: { type: 'number', minimum: 0, maximum: 1 },
                rotation: { type: 'number', minimum: -360, maximum: 360 },
              },
            },
            pageNumbers: { type: 'boolean' },
            attachments: {
              type: 'array',
              maxItems: 20,
              items: {
                type: 'object',
                additionalProperties: false,
                required: ['name', 'bytes'],
                properties: {
                  name: { type: 'string', maxLength: 100 },
                  bytes,
                  mimeType: { type: 'string', maxLength: 100 },
                },
              },
            },
          },
        },
      },
    },
    async (request, reply) => {
      const body = request.body;
      const output = await execute<Uint8Array>('pdf', {
        ...body,
        inputs: body.inputs.map((i) => ({ ...i, bytes: decodeBase64(i.bytes) })),
        attachments: body.attachments?.map((i) => ({ ...i, bytes: decodeBase64(i.bytes) })),
      });
      return reply
        .type('application/pdf')
        .header('Content-Disposition', 'attachment; filename="processed.pdf"')
        .send(Buffer.from(output));
    },
  );
  app.post<{ Body: { bytes: string; ranges: number[][] } }>(
    '/v1/pdf/split',
    {
      schema: {
        body: {
          type: 'object',
          additionalProperties: false,
          required: ['bytes', 'ranges'],
          properties: { bytes, ranges: { type: 'array', minItems: 1, maxItems: 20, items: pages } },
        },
      },
    },
    async (request) => ({
      documents: (
        await execute<Uint8Array[]>('pdf', {
          bytes: decodeBase64(request.body.bytes),
          ranges: request.body.ranges,
        })
      ).map((b) => Buffer.from(b).toString('base64')),
    }),
  );
  app.post<{ Body: BarcodeOptions }>(
    '/v1/barcode',
    {
      schema: {
        body: {
          type: 'object',
          additionalProperties: false,
          required: ['type', 'text'],
          properties: {
            type: { type: 'string', enum: ['qrcode', 'code128', 'ean13', 'datamatrix'] },
            text: { type: 'string', minLength: 1, maxLength: 512 },
            scale: { type: 'integer', minimum: 1, maximum: 6 },
            includeText: { type: 'boolean' },
          },
        },
      },
    },
    async (request, reply) =>
      reply.type('image/png').send(Buffer.from(await execute<Uint8Array>('barcode', request.body))),
  );
}
