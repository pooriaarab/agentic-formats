import type { ImageMime } from "./model.js";

export interface SniffedImage {
  mime: ImageMime;
  ext: "png" | "jpg" | "gif";
  /** Pixel size from the file header; 0 when the header does not give one. */
  width: number;
  height: number;
}

/**
 * Tell a PNG, JPEG, or GIF by its first bytes, never by a name or a declared type. Anything else
 * returns undefined. SVG is not on the list: it can carry script.
 */
export function sniffImage(b: Uint8Array): SniffedImage | undefined {
  const v = new DataView(b.buffer, b.byteOffset, b.byteLength);
  if (b.length >= 24 && v.getUint32(0) === 0x89504e47 && v.getUint32(4) === 0x0d0a1a0a) {
    return { mime: "image/png", ext: "png", width: v.getUint32(16), height: v.getUint32(20) };
  }
  if (b.length >= 10 && v.getUint32(0) === 0x47494638 && (b[4] === 0x37 || b[4] === 0x39) && b[5] === 0x61) {
    return { mime: "image/gif", ext: "gif", width: v.getUint16(6, true), height: v.getUint16(8, true) };
  }
  if (b.length >= 4 && v.getUint16(0) === 0xffd8) {
    let width = 0;
    let height = 0;
    for (let i = 2; i + 9 < b.length && b[i] === 0xff; i += 2 + v.getUint16(i + 2)) {
      if (b[i + 1] >= 0xc0 && b[i + 1] <= 0xcf && ![0xc4, 0xc8, 0xcc].includes(b[i + 1])) {
        height = v.getUint16(i + 5);
        width = v.getUint16(i + 7);
        break;
      }
    }
    return { mime: "image/jpeg", ext: "jpg", width, height };
  }
  return undefined;
}

/** The widest or tallest an image frame gets: 22 inches. */
const MAX_POINTS = 1584;
const FALLBACK: [number, number] = [240, 180];

/**
 * The frame size in points for an image. The model size wins when it is a usable number;
 * otherwise the header size at 96 dpi; otherwise a fixed size. The result keeps its shape and
 * fits inside 22 inches, so a hostile header or model cannot ask for a huge frame.
 */
export function frameSize(width: number | undefined, height: number | undefined, sniffed: SniffedImage): [number, number] {
  const usable = (n: number | undefined): n is number => typeof n === "number" && Number.isFinite(n) && n > 0;
  const w = usable(width) ? width : sniffed.width > 0 ? sniffed.width * 0.75 : FALLBACK[0];
  const h = usable(height) ? height : sniffed.height > 0 ? sniffed.height * 0.75 : FALLBACK[1];
  const scale = Math.min(1, MAX_POINTS / w, MAX_POINTS / h);
  const fit = (n: number) => Math.max(1, Math.round(n * scale * 10) / 10);
  return [fit(w), fit(h)];
}

const UNIT_POINTS: Record<string, number> = { cm: 72 / 2.54, mm: 72 / 25.4, in: 72, pt: 1, pc: 12, px: 0.75 };

/** A length such as `2.5cm` or `10pt` in points, to 0.1 point. Anything else returns undefined. */
export function lengthPt(value: string | undefined): number | undefined {
  const found = /^(-?[0-9]*\.?[0-9]+)(cm|mm|in|pt|pc|px)$/.exec((value ?? "").trim());
  if (!found) return undefined;
  const points = Number(found[1]) * UNIT_POINTS[found[2]];
  return Number.isFinite(points) ? Math.round(points * 10) / 10 : undefined;
}

export const cm = (points: number): string => `${(points / UNIT_POINTS.cm).toFixed(4)}cm`;
