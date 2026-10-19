import { execFile } from 'node:child_process';
import { promisify } from 'node:util';
import { types as utilTypes } from 'node:util';
import { mkdtemp, writeFile, rm } from 'node:fs/promises';
import { tmpdir } from 'node:os';
import { isAbsolute, join } from 'node:path';
import { PDFDocument, PDFName, PDFNumber } from 'pdf-lib';
import { cloneAssuranceJson } from './assurance-input.ts';
import { comparePageImages } from './visual-diff.ts';
import type { VisualDiff, VisualMask } from './visual-diff.ts';
const execute = promisify(execFile);
export interface PdfRegressionOptions {
  /** Trusted deployment configuration, never an untrusted request parameter. */
  pdftoppmPath: string;
  dpi?: number;
  timeoutMs?: number;
  maxPages?: number;
  maxTotalPixels?: number;
  channelTolerance?: number;
  maxChangedRatio?: number;
  pageMasks?: { page: number; masks: VisualMask[] }[];
}
export interface PdfRegressionResult {
  passed: boolean;
  beforePages: number;
  afterPages: number;
  renderer: { executable: string; version: string; dpi: number };
  pages: { page: number; passed: boolean; diff?: VisualDiff; reason?: string }[];
  reason?: string;
}
/** Visual comparison against an externally approved PDF. Does not approve its own baseline or validate accessibility. */
export async function comparePdfDocuments(
  before: Uint8Array,
  after: Uint8Array,
  input: PdfRegressionOptions,
): Promise<PdfRegressionResult> {
  for (const bytes of [before, after])
    if (
      utilTypes.isProxy(bytes) ||
      !(bytes instanceof Uint8Array) ||
      bytes.byteLength > 16 * 1024 * 1024 ||
      !Buffer.from(bytes.subarray(0, 5)).equals(Buffer.from('%PDF-'))
    )
      throw new Error('Expected PDF inputs up to 16 MiB each');
  const approved = Buffer.from(before),
    candidate = Buffer.from(after),
    options = cloneAssuranceJson(input);
  if (
    !options ||
    typeof options !== 'object' ||
    Array.isArray(options) ||
    Object.keys(options).some(
      (k) =>
        ![
          'pdftoppmPath',
          'dpi',
          'timeoutMs',
          'maxPages',
          'maxTotalPixels',
          'channelTolerance',
          'maxChangedRatio',
          'pageMasks',
        ].includes(k),
    )
  )
    throw new Error('Invalid PDF comparison options');
  if (
    typeof options.pdftoppmPath !== 'string' ||
    !isAbsolute(options.pdftoppmPath) ||
    options.pdftoppmPath.includes('\0')
  )
    throw new Error('A trusted absolute pdftoppm executable path is required');
  const dpi = options.dpi ?? 96,
    timeoutMs = options.timeoutMs ?? 30_000,
    maxPages = options.maxPages ?? 20,
    maxTotalPixels = options.maxTotalPixels ?? 64_000_000;
  for (const [value, min, max] of [
    [dpi, 36, 200],
    [timeoutMs, 1, 120_000],
    [maxPages, 1, 100],
    [maxTotalPixels, 1, 256_000_000],
  ])
    if (!Number.isSafeInteger(value) || value < min || value > max)
      throw new Error('Invalid PDF comparison budget');
  if (
    options.channelTolerance !== undefined &&
    (!Number.isInteger(options.channelTolerance) ||
      options.channelTolerance < 0 ||
      options.channelTolerance >= 255)
  )
    throw new Error('Invalid PDF channel tolerance');
  if (
    options.maxChangedRatio !== undefined &&
    (!Number.isFinite(options.maxChangedRatio) || options.maxChangedRatio < 0 || options.maxChangedRatio >= 1)
  )
    throw new Error('Invalid PDF change threshold');
  const masks = options.pageMasks ?? [];
  if (
    !Array.isArray(masks) ||
    masks.length > maxPages ||
    masks.some(
      (m) =>
        !m ||
        Object.keys(m).sort().join(',') !== 'masks,page' ||
        !Number.isSafeInteger(m.page) ||
        m.page < 1 ||
        !Array.isArray(m.masks) ||
        m.masks.length > 100,
    ) ||
    new Set(masks.map((m) => m.page)).size !== masks.length
  )
    throw new Error('Invalid PDF page masks');
  const deadline = Date.now() + timeoutMs;
  const run = async (args: string[], maxBuffer: number) => {
    const remaining = deadline - Date.now();
    if (remaining <= 0) throw new Error('PDF comparison deadline exceeded');
    try {
      return await execute(options.pdftoppmPath, args, {
        encoding: 'buffer',
        timeout: remaining,
        killSignal: 'SIGKILL',
        maxBuffer,
        windowsHide: true,
      });
    } catch (error) {
      const e = error as { killed?: boolean; code?: string };
      throw new Error(
        e.killed
          ? 'PDF comparison deadline exceeded'
          : e.code === 'ERR_CHILD_PROCESS_STDIO_MAXBUFFER'
            ? 'PDF raster output exceeds byte budget'
            : 'Cannot execute trusted PDF rasterizer',
      );
    }
  };
  const a = await PDFDocument.load(approved, { throwOnInvalidObject: true, updateMetadata: false }),
    b = await PDFDocument.load(candidate, { throwOnInvalidObject: true, updateMetadata: false });
  const beforePages = a.getPageCount(),
    afterPages = b.getPageCount();
  if (!beforePages || !afterPages || Math.max(beforePages, afterPages) > maxPages)
    throw new Error('PDF page budget exceeded');
  if (masks.some((m) => m.page > Math.min(beforePages, afterPages)))
    throw new Error('PDF mask references a missing page');
  let pixels = 0;
  const sizes = [a, b].map((doc) =>
    doc.getPages().map((page) => {
      const { width, height } = page.getCropBox(),
        angle = page.getRotation().angle;
      if (
        !Number.isFinite(width) ||
        !Number.isFinite(height) ||
        width <= 0 ||
        height <= 0 ||
        !Number.isFinite(angle) ||
        angle % 90 !== 0
      )
        throw new Error('Invalid PDF page dimensions or rotation');
      const userUnit = page.node.get(PDFName.of('UserUnit'));
      if (userUnit !== undefined && (!(userUnit instanceof PDFNumber) || userUnit.asNumber() !== 1))
        throw new Error('Nonstandard PDF UserUnit is unsupported');
      const w = Math.ceil((width * dpi) / 72),
        h = Math.ceil((height * dpi) / 72);
      if (!Number.isSafeInteger(w * h) || w * h > 16_000_000 || w > 4096 || h > 4096)
        throw new Error('PDF page pixel budget exceeded');
      pixels += w * h;
      return {
        sourceWidth: w,
        sourceHeight: h,
        ...(angle % 180 ? { width: h, height: w } : { width: w, height: h }),
      };
    }),
  );
  if (pixels > maxTotalPixels) throw new Error('PDF total pixel budget exceeded');
  const version = await run(['-v'], 16 * 1024);
  const versionText = Buffer.concat([version.stdout, version.stderr]).toString().trim();
  if (!/^pdftoppm version \S+/m.test(versionText))
    throw new Error('Configured executable does not identify as pdftoppm');
  const result: PdfRegressionResult = {
    passed: false,
    beforePages,
    afterPages,
    renderer: { executable: options.pdftoppmPath, version: versionText.split('\n')[0]!, dpi },
    pages: [],
  };
  if (beforePages !== afterPages) return { ...result, reason: 'Page counts differ' };
  const directory = await mkdtemp(join(tmpdir(), 'imprint-pdf-regression-'));
  try {
    const beforePath = join(directory, 'approved.pdf'),
      afterPath = join(directory, 'candidate.pdf');
    await writeFile(beforePath, approved, { mode: 0o600 });
    await writeFile(afterPath, candidate, { mode: 0o600 });
    for (let i = 0; i < beforePages; i++) {
      const page = i + 1;
      if (sizes[0]![i]!.width !== sizes[1]![i]!.width || sizes[0]![i]!.height !== sizes[1]![i]!.height) {
        result.pages.push({ page, passed: false, reason: 'Page dimensions differ' });
        continue;
      }
      const raster = async (path: string, side: number) =>
        (
          await run(
            [
              '-f',
              String(page),
              '-l',
              String(page),
              '-r',
              String(dpi),
              '-singlefile',
              '-cropbox',
              '-scale-to-x',
              String(sizes[side]![i]!.sourceWidth),
              '-scale-to-y',
              String(sizes[side]![i]!.sourceHeight),
              '-png',
              path,
            ],
            16 * 1024 * 1024,
          )
        ).stdout;
      const left = await raster(beforePath, 0),
        right = await raster(afterPath, 1);
      const pageOptions = {
        ...(options.channelTolerance !== undefined ? { channelTolerance: options.channelTolerance } : {}),
        ...(options.maxChangedRatio !== undefined ? { maxChangedRatio: options.maxChangedRatio } : {}),
        masks: masks.find((m) => m.page === page)?.masks ?? [],
      };
      const diff = await comparePageImages(left, right, pageOptions);
      if (diff.width !== sizes[0]![i]!.width || diff.height !== sizes[0]![i]!.height)
        throw new Error('Raster dimensions do not match PDF page geometry');
      if (Date.now() > deadline) throw new Error('PDF comparison deadline exceeded');
      result.pages.push({ page, passed: diff.passed, diff });
    }
    result.passed = result.pages.every((page) => page.passed);
    return result;
  } finally {
    await rm(directory, { recursive: true, force: true });
  }
}
