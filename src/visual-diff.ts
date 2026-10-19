import { cloneAssuranceJson } from './assurance-input.ts';
import sharp from 'sharp';
export interface VisualMask {
  left: number;
  top: number;
  width: number;
  height: number;
}
export interface VisualDiffOptions {
  masks?: VisualMask[];
  channelTolerance?: number;
  maxChangedRatio?: number;
}
export interface VisualDiff {
  passed: boolean;
  width: number;
  height: number;
  comparedPixels: number;
  changedPixels: number;
  changedRatio: number;
  maskedPixels: number;
}
/** Compares supplied PNG page images; does not render documents or certify PDF accessibility. */
export async function comparePageImages(
  before: Uint8Array,
  after: Uint8Array,
  options: VisualDiffOptions = {},
): Promise<VisualDiff> {
  options = cloneAssuranceJson(options);
  before = Buffer.from(before);
  after = Buffer.from(after);
  if (
    !options ||
    typeof options !== 'object' ||
    Object.keys(options).some((k) => !['masks', 'channelTolerance', 'maxChangedRatio'].includes(k))
  )
    throw new Error('Invalid visual comparison options');
  const tolerance = options.channelTolerance ?? 0,
    maxRatio = options.maxChangedRatio ?? 0,
    masks = options.masks ?? [];
  if (
    !Number.isInteger(tolerance) ||
    tolerance < 0 ||
    tolerance >= 255 ||
    !Number.isFinite(maxRatio) ||
    maxRatio < 0 ||
    maxRatio >= 1 ||
    !Array.isArray(masks) ||
    masks.length > 100
  )
    throw new Error('Invalid visual comparison thresholds');
  const decode = async (bytes: Uint8Array) => {
    if (!(bytes instanceof Uint8Array) || bytes.byteLength > 16 * 1024 * 1024)
      throw new Error('PNG exceeds byte budget');
    const image = sharp(bytes, { limitInputPixels: 16_000_000, failOn: 'warning' });
    const metadata = await image.metadata();
    if (metadata.format !== 'png' || (metadata.pages ?? 1) !== 1) throw new Error('Expected one PNG page');
    return image.toColourspace('srgb').ensureAlpha().raw().toBuffer({ resolveWithObject: true });
  };
  const a = await decode(before),
    b = await decode(after);
  if (a.info.width !== b.info.width || a.info.height !== b.info.height)
    throw new Error('Page dimensions differ');
  const { width, height } = a.info,
    count = width * height,
    excluded = new Uint8Array(count);
  for (const mask of masks) {
    if (
      !mask ||
      Object.keys(mask).sort().join(',') !== 'height,left,top,width' ||
      Object.values(mask).some((v) => !Number.isSafeInteger(v) || v < 0) ||
      !mask.width ||
      !mask.height ||
      mask.left + mask.width > width ||
      mask.top + mask.height > height
    )
      throw new Error('Mask must be a nonempty rectangle inside the page');
    for (let y = mask.top; y < mask.top + mask.height; y++)
      excluded.fill(1, y * width + mask.left, y * width + mask.left + mask.width);
  }
  const maskedPixels = excluded.reduce((sum, n) => sum + n, 0),
    comparedPixels = count - maskedPixels;
  if (!comparedPixels) throw new Error('Masks cannot exclude the entire page');
  let changedPixels = 0;
  for (let i = 0; i < count; i++)
    if (!excluded[i]) {
      for (let channel = 0; channel < 4; channel++)
        if (Math.abs(a.data[i * 4 + channel]! - b.data[i * 4 + channel]!) > tolerance) {
          changedPixels++;
          break;
        }
    }
  const changedRatio = changedPixels / comparedPixels;
  return {
    passed: changedRatio <= maxRatio,
    width,
    height,
    comparedPixels,
    changedPixels,
    changedRatio,
    maskedPixels,
  };
}
