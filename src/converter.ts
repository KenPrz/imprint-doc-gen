import { RenderError } from './errors.ts';

export async function convertToPdf(document: Uint8Array, endpoint: string, filename = 'document.docx'): Promise<Buffer> {
  const form = new FormData();
  form.append('files', new Blob([new Uint8Array(document)]), filename);
  const response = await fetch(new URL('/forms/libreoffice/convert', endpoint), { method: 'POST', body: form, signal: AbortSignal.timeout(65_000) }).catch(error => {
    throw new RenderError(error?.name === 'TimeoutError' ? 'RENDER_TIMEOUT' : 'CONVERSION_FAILED', error?.name === 'TimeoutError' ? 'PDF conversion exceeded its time budget' : 'Cannot reach PDF converter');
  });
  if (!response.ok || !response.body) throw new RenderError('CONVERSION_FAILED', `PDF converter returned HTTP ${response.status}`);
  const chunks: Buffer[] = []; let total = 0;
  for await (const chunk of response.body) {
    total += chunk.byteLength;
    if (total > 32 * 1024 * 1024) throw new RenderError('LIMIT_EXCEEDED', 'Converted PDF exceeds 32 MiB');
    chunks.push(Buffer.from(chunk));
  }
  const output = Buffer.concat(chunks);
  if (!output.subarray(0, 5).equals(Buffer.from('%PDF-'))) throw new RenderError('CONVERSION_FAILED', 'Converter did not return a PDF');
  return output;
}
