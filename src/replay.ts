import { types as utilTypes } from 'node:util';
import { readdir, readFile } from 'node:fs/promises';
import { render } from './index.ts';
import { sha256 } from './document-diff.ts';
import type { RenderOptions } from './types.ts';
const maxBundle = 64 * 1024 * 1024;
interface ReplayPayload {
  schema: 'imprint-replay-v1';
  engine: string;
  template: string;
  data: Record<string, unknown>;
  options: Record<string, unknown>;
  outputSha256: string;
}
export interface ReplayBundle {
  payload: ReplayPayload;
  sha256: string;
}
/** Binds replay to source/build bytes, lockfile, Node version and OS/architecture. No PDF converter is involved. */
export async function replayEngineFingerprint(): Promise<string> {
  const root = new URL('./', import.meta.url);
  const names = (await readdir(root)).filter((n) => /\.(ts|js)$/.test(n) && !n.endsWith('.d.ts')).sort();
  const files = await Promise.all(
    names.map(async (name) => [name, sha256(await readFile(new URL(name, root)))]),
  );
  const lock = await readFile(new URL('../package-lock.json', root));
  return sha256(
    JSON.stringify({
      files,
      lock: sha256(lock),
      node: process.version,
      platform: process.platform,
      arch: process.arch,
    }),
  );
}
function jsonClone(value: unknown): unknown {
  let count = 0;
  const seen = new Set<object>();
  function walk(v: unknown, depth: number): unknown {
    if (utilTypes.isProxy(v)) throw new Error('Replay does not support proxies');
    if (++count > 200_000 || depth > 64) throw new Error('Replay JSON complexity limit exceeded');
    if (v === null || typeof v === 'string' || typeof v === 'boolean') return v;
    if (typeof v === 'number' && Number.isFinite(v)) return v;
    if (
      typeof v !== 'object' ||
      !v ||
      (!Array.isArray(v) &&
        Object.getPrototypeOf(v) !== Object.prototype &&
        Object.getPrototypeOf(v) !== null)
    )
      throw new Error('Replay supports finite plain JSON data only');
    if (
      Object.getOwnPropertyNames(v).some(
        (k) => !(Array.isArray(v) && k === 'length') && !Object.getOwnPropertyDescriptor(v, k)!.enumerable,
      )
    )
      throw new Error('Replay does not support hidden properties');
    if (Object.getOwnPropertySymbols(v).length) throw new Error('Replay does not support symbol keys');
    if (seen.has(v)) throw new Error('Replay does not support circular data');
    seen.add(v);
    let result: unknown;
    if (Array.isArray(v)) {
      if (Object.keys(v).length !== v.length) throw new Error('Replay does not support sparse arrays');
      result = Array.from({ length: v.length }, (_, index) => {
        const descriptor = Object.getOwnPropertyDescriptor(v, String(index));
        if (!descriptor || !('value' in descriptor))
          throw new Error('Replay does not support getters or sparse arrays');
        return walk(descriptor.value, depth + 1);
      });
    } else {
      const out: Record<string, unknown> = Object.create(null);
      for (const key of Object.keys(v).sort()) {
        const descriptor = Object.getOwnPropertyDescriptor(v, key)!;
        if (!('value' in descriptor)) throw new Error('Replay does not support getters');
        out[key] = walk(descriptor.value, depth + 1);
      }
      result = out;
    }
    seen.delete(v);
    return result;
  }
  const out = walk(value, 0);
  if (Buffer.byteLength(JSON.stringify(out)) > maxBundle) throw new Error('Replay exceeds 64 MiB');
  return out;
}
function object(value: unknown): value is Record<string, unknown> {
  return !!value && typeof value === 'object' && !Array.isArray(value);
}
function decode(value: unknown): Buffer {
  if (
    typeof value !== 'string' ||
    value.length > maxBundle ||
    value.length % 4 !== 0 ||
    !/^[A-Za-z0-9+/]*={0,2}$/.test(value)
  )
    throw new Error('Invalid replay base64');
  const bytes = Buffer.from(value, 'base64');
  if (bytes.toString('base64') !== value || bytes.byteLength > 8 * 1024 * 1024)
    throw new Error('Invalid replay base64');
  return bytes;
}
function decodeOptions(value: Record<string, unknown>): RenderOptions {
  const allowed = new Set(['restartNumbering', 'charts', 'attachments', 'images', 'locale', 'timeZone']);
  if (Object.keys(value).some((k) => !allowed.has(k))) throw new Error('Unsupported replay option');
  if (value.restartNumbering !== undefined && typeof value.restartNumbering !== 'boolean')
    throw new Error('Invalid replay numbering option');
  for (const key of ['locale', 'timeZone'])
    if (value[key] !== undefined && (typeof value[key] !== 'string' || value[key].length > 200))
      throw new Error('Invalid replay locale option');
  new Intl.DateTimeFormat(value.locale as string | undefined, {
    timeZone: value.timeZone as string | undefined,
  });
  for (const key of ['images', 'charts'])
    if (value[key] !== undefined && !object(value[key])) throw new Error('Invalid replay option map');
  if (object(value.images))
    for (const image of Object.values(value.images)) {
      if (
        !object(image) ||
        Object.keys(image).some(
          (k) => !['fit', 'width', 'height', 'rotate', 'background', 'crop'].includes(k),
        )
      )
        throw new Error('Invalid replay image options');
      if (image.fit !== undefined && !['contain', 'cover', 'fill'].includes(image.fit as string))
        throw new Error('Invalid replay image fit');
      for (const key of ['width', 'height'])
        if (
          image[key] !== undefined &&
          (!Number.isSafeInteger(image[key]) || (image[key] as number) < 1 || (image[key] as number) > 4096)
        )
          throw new Error('Invalid replay image dimension');
      if (image.rotate !== undefined && (typeof image.rotate !== 'number' || Math.abs(image.rotate) > 360))
        throw new Error('Invalid replay rotation');
      if (
        image.background !== undefined &&
        (typeof image.background !== 'string' || !/^#[0-9a-f]{6}(?:[0-9a-f]{2})?$/i.test(image.background))
      )
        throw new Error('Invalid replay image background');
      if (
        image.crop !== undefined &&
        (!object(image.crop) ||
          Object.keys(image.crop).sort().join(',') !== 'height,left,top,width' ||
          Object.values(image.crop).some((n) => !Number.isSafeInteger(n) || (n as number) < 0) ||
          !image.crop.width ||
          !image.crop.height)
      )
        throw new Error('Invalid replay image crop');
    }
  if (object(value.charts))
    for (const [path, chart] of Object.entries(value.charts)) {
      if (
        !/^(word|xl|ppt)\/charts\/chart[\w-]*\.xml$/.test(path) ||
        !object(chart) ||
        !Array.isArray(chart.categories) ||
        !Array.isArray(chart.series) ||
        !chart.series.length ||
        chart.series.length > 100 ||
        chart.categories.length > 10000 ||
        chart.categories.some((v) => typeof v !== 'string' || v.length > 32767)
      )
        throw new Error('Invalid replay chart');
      const length = chart.categories.length;
      for (const series of chart.series)
        if (
          !object(series) ||
          typeof series.name !== 'string' ||
          !Array.isArray(series.values) ||
          series.values.length !== length ||
          series.values.some((v) => typeof v !== 'number')
        )
          throw new Error('Invalid replay chart series');
    }
  const options = { ...value } as RenderOptions;
  if (value.attachments !== undefined) {
    if (!object(value.attachments)) throw new Error('Invalid replay attachments');
    const attachments: Record<string, Uint8Array> = Object.create(null);
    for (const [name, bytes] of Object.entries(value.attachments)) attachments[name] = decode(bytes);
    options.attachments = attachments;
  }
  return options;
}
/** Bundle includes potentially sensitive source data. Hashes detect corruption, not a malicious author who recomputes them. */
export async function createReplayBundle(
  template: Uint8Array,
  data: Record<string, unknown>,
  options: RenderOptions = {},
): Promise<ReplayBundle> {
  if (utilTypes.isProxy(template)) throw new Error('Replay does not support proxies');
  if (!(template instanceof Uint8Array) || template.byteLength > 8 * 1024 * 1024)
    throw new Error('Invalid replay template');
  if (!object(data)) throw new Error('Replay data must be an object');
  function ownData(value: unknown): Record<string, unknown> {
    if (utilTypes.isProxy(value)) throw new Error('Replay does not support proxies');
    if (
      !object(value) ||
      (Object.getPrototypeOf(value) !== Object.prototype && Object.getPrototypeOf(value) !== null) ||
      Object.getOwnPropertySymbols(value).length
    )
      throw new Error('Replay options must be plain objects');
    const result: Record<string, unknown> = Object.create(null);
    for (const key of Object.getOwnPropertyNames(value)) {
      if (!Object.getOwnPropertyDescriptor(value, key)!.enumerable)
        throw new Error('Replay does not support hidden options');
      const descriptor = Object.getOwnPropertyDescriptor(value, key)!;
      if (!('value' in descriptor)) throw new Error('Replay does not support getters');
      result[key] = descriptor.value;
    }
    return result;
  }
  const encoded = ownData(options);
  if (encoded.attachments !== undefined) {
    const attachments: Record<string, string> = Object.create(null);
    for (const [name, bytes] of Object.entries(ownData(encoded.attachments))) {
      if (!(bytes instanceof Uint8Array)) throw new Error('Invalid replay attachment');
      attachments[name] = Buffer.from(bytes).toString('base64');
    }
    encoded.attachments = attachments;
  }
  const safeData = jsonClone(data) as Record<string, unknown>,
    safeOptions = jsonClone(encoded) as Record<string, unknown>;
  const safeTemplate = Buffer.from(template);
  const output = await render(safeTemplate, safeData, decodeOptions(safeOptions));
  const payload: ReplayPayload = {
    schema: 'imprint-replay-v1',
    engine: await replayEngineFingerprint(),
    template: safeTemplate.toString('base64'),
    data: safeData,
    options: safeOptions,
    outputSha256: sha256(output),
  };
  const canonical = JSON.stringify(jsonClone(payload));
  if (Buffer.byteLength(canonical) > maxBundle) throw new Error('Replay exceeds 64 MiB');
  return { payload, sha256: sha256(canonical) };
}
/** Strict replay refuses engine mismatches and output differences; it never silently updates a baseline. */
export async function replayBundle(input: unknown): Promise<Buffer> {
  const bundle = jsonClone(input);
  if (!object(bundle) || Object.keys(bundle).sort().join(',') !== 'payload,sha256' || !object(bundle.payload))
    throw new Error('Invalid replay envelope');
  const p = bundle.payload;
  if (
    Object.keys(p).sort().join(',') !== 'data,engine,options,outputSha256,schema,template' ||
    p.schema !== 'imprint-replay-v1' ||
    !object(p.data) ||
    !object(p.options) ||
    typeof p.outputSha256 !== 'string' ||
    !/^[a-f0-9]{64}$/.test(p.outputSha256)
  )
    throw new Error('Invalid replay manifest');
  if (bundle.sha256 !== sha256(JSON.stringify(p))) throw new Error('Replay integrity check failed');
  if (p.engine !== (await replayEngineFingerprint()))
    throw new Error('Replay engine mismatch: exact build, dependencies and runtime required');
  const output = await render(decode(p.template), p.data, decodeOptions(p.options));
  if (sha256(output) !== p.outputSha256) throw new Error('Replay output differs from recorded output');
  return output;
}
