import bwipjs from 'bwip-js';
import { RenderError } from './errors.ts';
export interface BarcodeOptions {
  type: 'qrcode' | 'code128' | 'ean13' | 'datamatrix';
  text: string;
  scale?: number;
  includeText?: boolean;
}
export async function barcode(options: BarcodeOptions): Promise<Buffer> {
  if (
    !['qrcode', 'code128', 'ean13', 'datamatrix'].includes(options.type) ||
    typeof options.text !== 'string' ||
    !options.text.length ||
    Buffer.byteLength(options.text) > 512 ||
    !Number.isInteger(options.scale ?? 3) ||
    (options.scale ?? 3) < 1 ||
    (options.scale ?? 3) > 6
  )
    throw new RenderError('INVALID_BARCODE', 'Unsupported barcode or text/scale limit exceeded');
  if (options.type === 'code128' && options.text.length > 100)
    throw new RenderError('INVALID_BARCODE', 'Code 128 text exceeds 100 characters');
  try {
    return await bwipjs.toBuffer({
      bcid: options.type,
      text: options.text,
      scale: options.scale ?? 3,
      height: 10,
      includetext: options.includeText ?? false,
      padding: 4,
      backgroundcolor: 'FFFFFF',
    });
  } catch {
    throw new RenderError('INVALID_BARCODE', 'Text cannot be encoded using the selected barcode');
  }
}
