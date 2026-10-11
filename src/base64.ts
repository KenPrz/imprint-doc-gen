import { RenderError } from './errors.ts';

export function decodeBase64(value: string): Buffer {
  if (!value || value.length % 4 || !/^(?:[A-Za-z0-9+/]{4})*(?:[A-Za-z0-9+/]{2}==|[A-Za-z0-9+/]{3}=)?$/.test(value)) throw new RenderError('INVALID_DATA', 'Expected canonical base64 bytes');
  const bytes = Buffer.from(value, 'base64');
  if (bytes.toString('base64') !== value) throw new RenderError('INVALID_DATA', 'Expected canonical base64 bytes');
  return bytes;
}

