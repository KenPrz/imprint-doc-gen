import { RenderError } from './errors.ts';

interface ConversionOptions { timeoutMs?: number; maxOutputBytes?: number }

async function readBounded(response: Response, limit: number, truncate = false): Promise<Buffer> {
  const chunks: Buffer[] = []; let total = 0;
  for await (const chunk of response.body!) {
    if (total + chunk.byteLength > limit) {
      if (!truncate) throw new RenderError('LIMIT_EXCEEDED', 'Converted PDF exceeds its byte budget');
      chunks.push(Buffer.from(chunk.subarray(0, limit - total)));
      return Buffer.concat(chunks); // Returning cancels the response stream.
    }
    chunks.push(Buffer.from(chunk)); total += chunk.byteLength;
  }
  return Buffer.concat(chunks);
}

/** Converts one document, or several merged in order into one PDF (Gotenberg merges by filename). */
export async function convertToPdf(document: Uint8Array | Uint8Array[], endpoint: string, filename = 'document.docx', options: ConversionOptions = {}): Promise<Buffer> {
  const controller = new AbortController();
  const timer = setTimeout(() => controller.abort(new DOMException('Conversion deadline exceeded', 'TimeoutError')), options.timeoutMs ?? 65_000);
  try {
    const form = new FormData();
    if (Array.isArray(document)) {
      const extension = filename.slice(filename.lastIndexOf('.'));
      document.forEach((sheet, index) => form.append('files', new Blob([new Uint8Array(sheet)]), `sheet-${String(index + 1).padStart(4, '0')}${extension}`));
      form.append('merge', 'true');
    } else form.append('files', new Blob([new Uint8Array(document)]), filename);
    const response = await fetch(new URL('/forms/libreoffice/convert', endpoint), { method: 'POST', body: form, signal: controller.signal });
    if (!response.ok || !response.body) {
      // Gotenberg uses 503 for both deadlines and capacity errors. Only its
      // specific timeout message maps to 504; other failures remain 502.
      const message = response.status === 503 && response.body ? (await readBounded(response, 1024, true)).toString() : '';
      if (response.status === 503 && message.startsWith('The request exceeded the time limit.')) throw new RenderError('RENDER_TIMEOUT', 'PDF conversion exceeded its time budget');
      throw new RenderError('CONVERSION_FAILED', `PDF converter returned HTTP ${response.status}`);
    }
    const output = await readBounded(response, options.maxOutputBytes ?? 32 * 1024 * 1024);
    if (!output.subarray(0, 5).equals(Buffer.from('%PDF-'))) throw new RenderError('CONVERSION_FAILED', 'Converter did not return a PDF');
    return output;
  } catch (error) {
    if (controller.signal.aborted) throw new RenderError('RENDER_TIMEOUT', 'PDF conversion exceeded its time budget');
    if (error instanceof RenderError) throw error;
    throw new RenderError('CONVERSION_FAILED', 'Cannot complete PDF conversion');
  } finally {
    clearTimeout(timer);
    controller.abort(); // Cancel unread error bodies and release the HTTP connection.
  }
}
