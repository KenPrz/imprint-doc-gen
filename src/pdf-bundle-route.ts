import { execFile } from 'node:child_process';
import { createWriteStream } from 'node:fs';
import { mkdir, mkdtemp, open, readdir, rm, stat, statfs } from 'node:fs/promises';
import { tmpdir } from 'node:os';
import { join } from 'node:path';
import { pipeline } from 'node:stream/promises';
import { promisify } from 'node:util';
import multipart from '@fastify/multipart';
import type { FastifyInstance } from 'fastify';
import { RenderError } from './errors.ts';
import type { RenderService } from './render-service.ts';
import { ServiceError } from './service-error.ts';
import { runTask } from './task-runner.ts';

const exec = promisify(execFile);
const flags = ['decrypt', 'flatten', 'stripActiveContent'] as const;
const MiB = 1024 * 1024;

/** Reads a positive integer setting. Operators may choose any value; only malformed input is refused. */
export function envInt(env: NodeJS.ProcessEnv, key: string, fallback: number) {
  const raw = env[key];
  if (raw === undefined || raw === '') return fallback;
  const value = Number(raw);
  if (!Number.isSafeInteger(value) || value < 1) throw new Error(`${key} must be a positive integer`);
  return value;
}

/**
 * Bundle limits. Defaults are the measured bounds (scripts/bench-pdf-bundle.ts, 2 cores): 20 × 10 MB mixed in
 * ~10 s, worst case 50 × 12 MP photos in ~25 s at ~560 MiB API peak. Values past them are allowed, with a warning.
 */
export function bundleConfig(env: NodeJS.ProcessEnv = process.env) {
  const config = {
    root: env.IMPRINT_PDF_BUNDLE_DIR || tmpdir(),
    timeout: envInt(env, 'IMPRINT_PDF_BUNDLE_TIMEOUT_MS', 60_000),
    maxFiles: envInt(env, 'IMPRINT_PDF_BUNDLE_MAX_FILES', 50),
    maxFileBytes: envInt(env, 'IMPRINT_PDF_BUNDLE_MAX_FILE_MB', 10) * MiB,
    maxTotalBytes: envInt(env, 'IMPRINT_PDF_BUNDLE_MAX_TOTAL_MB', 200) * MiB,
    warnings: [] as string[],
  };
  const untested = 'beyond the measured bounds; re-run npm run bench:pdf-bundle on the target host';
  if (config.maxFiles > 50) config.warnings.push(`IMPRINT_PDF_BUNDLE_MAX_FILES=${config.maxFiles} is ${untested}`);
  if (config.maxFileBytes > 10 * MiB)
    config.warnings.push(
      `IMPRINT_PDF_BUNDLE_MAX_FILE_MB=${config.maxFileBytes / MiB} is ${untested}; each input is processed in a 256 MiB worker heap`,
    );
  if (config.maxTotalBytes > 200 * MiB)
    config.warnings.push(
      `IMPRINT_PDF_BUNDLE_MAX_TOTAL_MB=${config.maxTotalBytes / MiB} is ${untested}; staging needs ~3× and the pdfcpu merge holds the whole output`,
    );
  return config;
}

/** Warns when the staging directory cannot hold a full bundle, or holds it in RAM. */
export async function stagingWarnings(root: string, maxTotalBytes: number) {
  const warnings: string[] = [];
  await mkdir(root, { recursive: true, mode: 0o700 });
  const fs = await statfs(root);
  const free = fs.bavail * fs.bsize;
  if (free < 3 * maxTotalBytes)
    warnings.push(
      `PDF bundle staging ${root} has ${Math.round(free / MiB)} MiB free; a full bundle stages up to ${Math.round((3 * maxTotalBytes) / MiB)} MiB`,
    );
  if (fs.type === 0x01021994)
    warnings.push(`PDF bundle staging ${root} is tmpfs (RAM-backed); set IMPRINT_PDF_BUNDLE_DIR to a disk volume`);
  return warnings;
}
const badRequest = (message: string): never => {
  throw new ServiceError(400, 'INVALID_REQUEST', message);
};

/** Runs pdfcpu with its config dir disabled (read-only containers), killed when the bundle deadline expires. */
async function pdfcpu(args: string[], timeout: number) {
  try {
    await exec('pdfcpu', ['-c', 'disable', '-q', ...args], { timeout, killSignal: 'SIGKILL', maxBuffer: 1024 * 1024 });
  } catch (error) {
    const failure = error as NodeJS.ErrnoException & { killed?: boolean; stderr?: string };
    if (failure.killed) throw new RenderError('RENDER_TIMEOUT', 'Operation exceeded its time budget');
    // A missing binary is a deployment fault: surface it as a logged 500, not a client error.
    if (failure.code === 'ENOENT') throw new Error('pdfcpu is not installed');
    throw Object.assign(new RenderError('INVALID_PDF_OPERATION', 'Cannot process PDF'), { stderr: failure.stderr ?? '' });
  }
}
/** Removes an owner password. Returns false when the file is not encrypted. */
async function decrypt(input: string, output: string, timeout: number) {
  try {
    await pdfcpu(['decrypt', input, output], timeout);
    return true;
  } catch (error) {
    const stderr = (error as { stderr?: string }).stderr ?? '';
    if (stderr.includes('not encrypted')) return false;
    if (stderr.includes('password')) throw new RenderError('PDF_PASSWORD_REQUIRED', 'PDF needs a password to open');
    throw error;
  }
}
const isPdf = async (path: string) => {
  const file = await open(path);
  try {
    const { buffer, bytesRead } = await file.read(Buffer.alloc(5), 0, 5, 0);
    return bytesRead === 5 && buffer.toString('latin1') === '%PDF-';
  } finally {
    await file.close();
  }
};

export function registerPdfBundleRoute(app: FastifyInstance, renderer: RenderService) {
  const { root, timeout, maxFiles, maxFileBytes, maxTotalBytes, warnings } = bundleConfig();
  for (const warning of warnings) app.log.warn(warning);
  void stagingWarnings(root, maxTotalBytes)
    .then((found) => found.forEach((warning) => app.log.warn(warning)))
    .catch((error: unknown) => app.log.warn({ err: error }, 'PDF bundle staging directory is not usable'));
  // Staging may be disk-backed and survive a crash. Remove leftovers older than any possible request.
  void readdir(root)
    .then((names) =>
      Promise.all(
        names
          .filter((name) => name.startsWith('imprint-bundle-'))
          .map(async (name) => {
            if (Date.now() - (await stat(join(root, name))).mtimeMs > 2 * timeout + 600_000)
              await rm(join(root, name), { recursive: true, force: true });
          }),
      ),
    )
    .catch(() => undefined);
  // ponytail: one bundle at a time per process. Each stages up to ~3× maxTotalBytes; raise when staging is sized for more.
  let busy = false;
  app.register(async (scope) => {
    await scope.register(multipart, {
      limits: { fileSize: maxFileBytes, files: maxFiles, fields: 4, fieldSize: 64, parts: maxFiles + 4, headerPairs: 100 },
    });
    scope.post('/v1/pdf/bundle', async (request, reply) => {
      if (!request.isMultipart()) throw new ServiceError(415, 'UNSUPPORTED_MEDIA_TYPE');
      if (busy || !renderer.available) throw new ServiceError(503, 'BUSY');
      busy = true;
      renderer.active++;
      // The deadline covers processing; the upload itself is bounded by the server's requestTimeout.
      let end = Infinity;
      const left = () => {
        const ms = end - Date.now();
        if (ms <= 0) throw new RenderError('RENDER_TIMEOUT', 'Operation exceeded its time budget');
        return ms;
      };
      let work: string | undefined;
      try {
        await mkdir(root, { recursive: true, mode: 0o700 });
        work = await mkdtemp(join(root, 'imprint-bundle-'));
        const options: Record<(typeof flags)[number], boolean> = { decrypt: false, flatten: false, stripActiveContent: false };
        let pageSize: 'A4' | 'Letter' = 'A4';
        const files: string[] = [];
        let total = 0;
        for await (const part of request.parts()) {
          if (part.type === 'file') {
            if (part.fieldname !== 'files') badRequest('Upload inputs as "files" parts');
            const path = join(work, `in-${files.length}.pdf`); // pdfcpu requires the .pdf extension
            const file = createWriteStream(path, { flags: 'wx', mode: 0o600 });
            await pipeline(part.file, file);
            files.push(path);
            // ponytail: checked per completed file, so staging can overshoot by at most one file.
            if ((total += file.bytesWritten) > maxTotalBytes)
              throw new RenderError('LIMIT_EXCEEDED', `Bundle inputs exceed ${maxTotalBytes / MiB} MiB in total`);
          } else if (flags.includes(part.fieldname as (typeof flags)[number]) && ['true', 'false'].includes(String(part.value)))
            options[part.fieldname as (typeof flags)[number]] = part.value === 'true';
          else if (part.fieldname === 'pageSize' && (part.value === 'A4' || part.value === 'Letter')) pageSize = part.value;
          else badRequest(`Unknown or invalid field "${part.fieldname}"`);
        }
        if (!files.length) badRequest('Supply at least one file');
        end = Date.now() + timeout;
        const outputs: string[] = [];
        const stripped: string[] = [];
        let pages = 0;
        for (const [index, path] of files.entries()) {
          let source = path;
          if (options.decrypt && (await isPdf(path)) && (await decrypt(path, `${path}.plain.pdf`, left()))) source = `${path}.plain.pdf`;
          const output = join(work, `out-${index}.pdf`);
          const result = await runTask<{ pages: number; stripped: string[] }>(
            'bundle',
            { path: source, out: output, flatten: options.flatten, stripActiveContent: options.stripActiveContent, pageSize },
            left(),
          ).catch((error: unknown) => {
            if (error instanceof RenderError) throw new RenderError(error.code, `Input ${index + 1}: ${error.message}`, error.part);
            throw error;
          });
          await rm(path, { force: true });
          await rm(`${path}.plain.pdf`, { force: true });
          pages += result.pages;
          if (pages > 1000) throw new RenderError('LIMIT_EXCEEDED', 'Output exceeds 1000 pages');
          if (result.stripped.length) stripped.push(`${index + 1}:${result.stripped.join('+')}`);
          outputs.push(output);
        }
        const merged = join(work, 'bundle.pdf');
        await pdfcpu(['merge', '-b=false', merged, ...outputs], left());
        // The open handle keeps the file readable after the staging directory is removed below.
        const handle = await open(merged);
        if (stripped.length) reply.header('Imprint-Stripped-Active-Content', stripped.join(', '));
        return reply
          .type('application/pdf')
          .header('Content-Disposition', 'attachment; filename="bundle.pdf"')
          .send(handle.createReadStream());
      } finally {
        // Release admission before the async cleanup: a fast response can finish first, and the client's next
        // request must not see a stale BUSY. The next request stages in its own directory.
        renderer.active--;
        busy = false;
        if (work) await rm(work, { recursive: true, force: true });
      }
    });
  });
}
