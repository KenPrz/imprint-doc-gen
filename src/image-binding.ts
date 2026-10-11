import { carboneLookup } from './carbone.ts';
import { decodeBase64 } from './base64.ts';
import { RenderError } from './errors.ts';
import type { Scope } from './scope.ts';

// Per-render state: data URLs must never reuse another placement's cached image.
export function carboneImageResolver(attachments: Record<string, Uint8Array>) {
  let nextImage = 0;
  return (path: string, scope: Scope, part: string): string | undefined => {
    const value = carboneLookup(scope, path, part);
    const rootPath = path.replace(/^(?:\.\.\/)+/, '');
    const atRoot = (path.length - rootPath.length) / 3 === scope.length - 1;
    let key = atRoot && Object.hasOwn(attachments, rootPath) ? rootPath : typeof value === 'string' ? value : '';
    if (!key && value != null && value !== '') throw new RenderError('INVALID_DATA', 'Expected an image attachment key or data URL', part);
    if (!key) return undefined;
    if (key.startsWith('data:')) {
      const match = /^data:image\/(?:png|jpeg);base64,([\s\S]+)$/.exec(key);
      if (!match || match[1]!.length > 12 * 1024 * 1024) throw new RenderError('UNSUPPORTED_IMAGE', 'Expected a bounded PNG/JPEG data URL', part);
      do { key = `__imprintCarboneImage${nextImage++}`; } while (Object.hasOwn(attachments, key));
      attachments[key] = decodeBase64(match[1]!);
    }
    if (/^https?:/i.test(key)) throw new RenderError('UNSUPPORTED_IMAGE', 'Remote image fetches are not supported; use attachments or data URLs', part);
    return key;
  };
}
