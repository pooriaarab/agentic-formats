/*
 * PDF writer for `Deck` (one page per slide) and `TextDoc` (Letter pages). Plain JavaScript, no font files,
 * no DOM: it runs in Node and in a Worker. Text uses the standard Helvetica and Courier faces with WinAnsi
 * encoding, so a character outside Latin-1 and the Windows punctuation is written as "?" and reported.
 * The output is deterministic: a fixed creation date and no random ids.
 */
import { Unzlib, zlibSync } from "fflate";
import { checkDeck } from "./deck.js";
import { frameSize, sniffImage } from "./image.js";
import { FormatError, capped, withLimits, type Limits } from "./limits.js";
import { checkList, checkRun, headingLevel, textLength, type Block, type Deck, type Inline, type List, type TextDoc, type TextRun, type WriteResult } from "./model.js";
import { safeHref } from "./xml.js";

// ---------- text encoding and metrics ----------

const WIN_ANSI = new Map<number, number>(
  [0x20ac, 0, 0x201a, 0x0192, 0x201e, 0x2026, 0x2020, 0x2021, 0x02c6, 0x2030, 0x0160, 0x2039, 0x0152, 0, 0x017d, 0, 0, 0x2018, 0x2019, 0x201c, 0x201d, 0x2022, 0x2013, 0x2014, 0x02dc, 0x2122, 0x0161, 0x203a, 0x0153, 0, 0x017e, 0x0178].flatMap((cp, i) => (cp ? [[cp, 0x80 + i] as [number, number]] : [])),
);

/** Widths of ASCII 32..126 in Helvetica and Helvetica-Bold, in 1/1000 em. */
const HELVETICA = [278, 278, 355, 556, 556, 889, 667, 191, 333, 333, 389, 584, 278, 333, 278, 278, 556, 556, 556, 556, 556, 556, 556, 556, 556, 556, 278, 278, 584, 584, 584, 556, 1015, 667, 667, 722, 722, 667, 611, 778, 722, 278, 500, 667, 556, 833, 722, 778, 667, 778, 722, 667, 611, 722, 667, 944, 667, 667, 611, 278, 278, 278, 469, 556, 333, 556, 556, 500, 556, 556, 278, 556, 556, 222, 222, 500, 222, 833, 556, 556, 556, 556, 333, 500, 278, 556, 500, 722, 500, 500, 500, 334, 260, 334, 584];
const HELVETICA_BOLD = [278, 333, 474, 556, 556, 889, 722, 238, 333, 333, 389, 584, 278, 333, 278, 278, 556, 556, 556, 556, 556, 556, 556, 556, 556, 556, 333, 333, 584, 584, 584, 611, 975, 722, 722, 722, 722, 667, 611, 778, 722, 278, 556, 722, 611, 833, 722, 778, 667, 778, 722, 667, 611, 722, 667, 944, 667, 667, 611, 333, 278, 333, 584, 556, 333, 556, 611, 556, 611, 556, 333, 611, 611, 278, 278, 556, 278, 889, 611, 611, 611, 611, 389, 556, 333, 611, 556, 778, 556, 556, 500, 389, 280, 389, 584];
const WIDE: Record<number, number> = { 0x95: 350, 0x97: 1000, 0x85: 1000, 0x99: 1000, 0x92: 222, 0x91: 222, 0x93: 333, 0x94: 333, 0xa0: 278 };

type Face = "F1" | "F2" | "F3" | "F4" | "F5";
const FACES: Record<Face, string> = { F1: "Helvetica", F2: "Helvetica-Bold", F3: "Helvetica-Oblique", F4: "Helvetica-BoldOblique", F5: "Courier" };
const faceOf = (marks: readonly string[] | undefined, bold: boolean): Face => {
  if (marks?.includes("code")) return "F5";
  const b = bold || !!marks?.includes("bold");
  const i = !!marks?.includes("italic");
  return b ? (i ? "F4" : "F2") : i ? "F3" : "F1";
};
const charWidth = (byte: number, face: Face): number => {
  if (face === "F5") return 600;
  if (byte >= 32 && byte <= 126) return (face === "F2" || face === "F4" ? HELVETICA_BOLD : HELVETICA)[byte - 32];
  return WIDE[byte] ?? (face === "F2" || face === "F4" ? 611 : 556);
};

// ---------- the file ----------

const latin1 = (s: string): Uint8Array => {
  const out = new Uint8Array(s.length);
  for (let i = 0; i < s.length; i++) out[i] = s.charCodeAt(i) & 0xff;
  return out;
};
/** A PDF literal string of WinAnsi characters. */
const lit = (s: string): string => `(${s.replace(/[\\()]/g, "\\$&").replace(/[^\x20-\x7e]/g, (c) => `\\${c.charCodeAt(0).toString(8).padStart(3, "0")}`)})`;
const num = (x: number): string => {
  const r = Math.round(x * 100) / 100;
  return Object.is(r, -0) ? "0" : String(r);
};
const rgb = (hex: string): string => [1, 3, 5].map((i) => num(parseInt(hex.slice(i, i + 2), 16) / 255)).join(" ");

class Page {
  ops: string[] = [];
  annots: string[] = [];
  constructor(
    readonly width: number,
    readonly height: number,
  ) {}
}

interface Img {
  name: string;
  /** Object id of the image. */
  id: number;
}

class Pdf {
  objects: (Uint8Array | undefined)[] = [undefined];
  pages: Page[] = [];
  images = new Map<Uint8Array, Img | null>();
  /** Images by their bytes, so two copies of one file are stored once. */
  byContent = new Map<string, { data: Uint8Array; img: Img | null }>();
  warnings = new Set<string>();
  charge: (n: number) => void;
  bytes: (n: number) => void;
  /** Bytes of pixels unpacked so far, for every image together. */
  decoded: (n: number) => void;
  constructor(readonly caps: Limits) {
    this.decoded = capped(caps.maxTotalBytes, `the images unpack to more than ${caps.maxTotalBytes} bytes`);
    this.charge = capped(caps.maxTextChars, `the file has more than ${caps.maxTextChars} characters of text`);
    this.bytes = capped(caps.maxTotalBytes, `the PDF passes ${caps.maxTotalBytes} bytes while it is written`);
  }
  alloc(): number {
    this.objects.push(undefined);
    return this.objects.length - 1;
  }
  set(id: number, dict: string, stream?: Uint8Array): void {
    const head = latin1(stream ? `${dict.slice(0, -2)} /Length ${stream.length}>>\nstream\n` : dict);
    const tail = stream ? latin1("\nendstream") : new Uint8Array();
    const out = new Uint8Array(head.length + (stream?.length ?? 0) + tail.length);
    out.set(head);
    if (stream) out.set(stream, head.length);
    out.set(tail, head.length + (stream?.length ?? 0));
    this.bytes(out.length);
    this.objects[id] = out;
  }
  newPage(width: number, height: number): Page {
    if (this.pages.length >= this.caps.maxEntries) throw new FormatError("LIMIT_EXCEEDED", `the PDF has more than ${this.caps.maxEntries} pages`);
    const page = new Page(width, height);
    this.pages.push(page);
    return page;
  }

  /** WinAnsi text for a string; each character it cannot hold becomes "?" and is reported once. */
  encode(text: string): string {
    this.charge(text.length);
    let out = "";
    for (const ch of text) {
      const cp = ch.codePointAt(0)!;
      if (cp === 9) out += "    ";
      else if (cp === 10 || cp === 13) out += " ";
      else if (cp >= 32 && cp <= 126) out += ch;
      else if (cp >= 0xa0 && cp <= 0xff) out += ch;
      else if (WIN_ANSI.has(cp)) out += String.fromCharCode(WIN_ANSI.get(cp)!);
      else if (cp < 32 || (cp >= 0x7f && cp < 0xa0) || (cp >= 0x200b && cp <= 0x200f) || cp === 0xfeff) continue;
      else {
        out += "?";
        this.warnings.add("a character that the PDF fonts cannot hold was written as ?");
      }
    }
    return out;
  }

  /** Register an image and return its name, or undefined with a warning when it cannot be embedded. */
  image(data: Uint8Array): Img | undefined {
    if (!(data instanceof Uint8Array)) throw new FormatError("INVALID_MODEL", "image data must be a Uint8Array");
    const known = this.images.get(data);
    if (known !== undefined) return known ?? undefined;
    const key = `${data.length}:${hash(data)}`;
    const same = this.byContent.get(key);
    if (same && same.data.every((b, i) => b === data[i])) {
      this.images.set(data, same.img);
      return same.img ?? undefined;
    }
    const made = this.embed(data);
    this.images.set(data, made ?? null);
    if (!same) this.byContent.set(key, { data, img: made ?? null });
    return made;
  }

  private embed(data: Uint8Array): Img | undefined {
    const sniffed = sniffImage(data);
    const skip = (why: string): undefined => {
      this.warnings.add(`an image ${why}: not written`);
      return undefined;
    };
    if (!sniffed) return skip("is not a PNG, JPEG, or GIF");
    if (sniffed.mime === "image/gif") return skip("is a GIF, which the PDF writer does not decode");
    if (sniffed.width <= 0 || sniffed.height <= 0) return skip("has no readable size");
    const id = this.alloc();
    const img: Img = { name: `Im${id}`, id };
    if (sniffed.mime === "image/jpeg") {
      const comps = jpegComponents(data);
      if (comps !== 1 && comps !== 3) return skip("is a JPEG that is not gray or RGB");
      this.set(id, `<</Type /XObject /Subtype /Image /Width ${sniffed.width} /Height ${sniffed.height} /ColorSpace /${comps === 1 ? "DeviceGray" : "DeviceRGB"} /BitsPerComponent 8 /Filter /DCTDecode>>`, data);
      return img;
    }
    return this.embedPng(data, id, img, skip);
  }

  private embedPng(data: Uint8Array, id: number, img: Img, skip: (why: string) => undefined): Img | undefined {
    const v = new DataView(data.buffer, data.byteOffset, data.byteLength);
    const width = v.getUint32(16);
    const height = v.getUint32(20);
    const depth = data[24];
    const ctype = data[25];
    if (data.length < 29 || data[28] !== 0) return skip("is an interlaced PNG");
    let palette: Uint8Array | undefined;
    const idat: Uint8Array[] = [];
    let transparent = false;
    for (let i = 8; i + 8 <= data.length; ) {
      const len = v.getUint32(i);
      const type = String.fromCharCode(data[i + 4], data[i + 5], data[i + 6], data[i + 7]);
      if (len > data.length - i - 12) break;
      const body = data.subarray(i + 8, i + 8 + len);
      if (type === "PLTE") palette = body;
      else if (type === "IDAT") idat.push(body);
      else if (type === "tRNS") transparent = true;
      else if (type === "IEND") break;
      i += 12 + len;
    }
    const joined = new Uint8Array(idat.reduce((n, c) => n + c.length, 0));
    idat.reduce((at, c) => (joined.set(c, at), at + c.length), 0);
    if (joined.length === 0) return skip("is a PNG with no pixels");
    if (transparent && (ctype === 0 || ctype === 2 || ctype === 3)) this.warnings.add("PNG transparency from a tRNS chunk is not written");
    const colors = ctype === 2 ? 3 : ctype === 6 ? 3 : 1;
    if (ctype === 0 || ctype === 2 || ctype === 3) {
      if (ctype === 3 && (!palette || palette.length < 3 || palette.length > 768)) return skip("is a PNG with no palette");
      if ((ctype === 2 && depth < 8) || (ctype === 3 && depth > 8) || ![1, 2, 4, 8, 16].includes(depth)) return skip("is an unusual PNG");
      // The viewer unpacks this stream, so its size counts against the same caps as an image that is unpacked here.
      const unpacked = height * (1 + Math.ceil((width * colors * depth) / 8));
      if (unpacked > this.caps.maxEntryBytes) throw new FormatError("LIMIT_EXCEEDED", "an image is larger than the entry limit when unpacked");
      this.decoded(unpacked);
      // Unpack it here, bounded, so a damaged or bomb stream is reported and not handed to the viewer.
      const rows = inflateUpTo(joined, unpacked);
      const stride = 1 + Math.ceil((width * colors * depth) / 8);
      if (!rows || rows.length !== unpacked) return skip("is a PNG with damaged pixels");
      for (let y = 0; y < height; y++) if (rows[y * stride] > 4) return skip("is a PNG with damaged pixels");
      const space = ctype === 3 ? `[/Indexed /DeviceRGB ${Math.floor(palette!.length / 3) - 1} <${[...palette!.subarray(0, Math.floor(palette!.length / 3) * 3)].map((b) => b.toString(16).padStart(2, "0")).join("")}>]` : ctype === 2 ? "/DeviceRGB" : "/DeviceGray";
      this.set(id, `<</Type /XObject /Subtype /Image /Width ${width} /Height ${height} /ColorSpace ${space} /BitsPerComponent ${depth} /Filter /FlateDecode /DecodeParms <</Predictor 15 /Colors ${colors} /BitsPerComponent ${depth} /Columns ${width}>>>>`, joined);
      return img;
    }
    if ((ctype !== 4 && ctype !== 6) || depth !== 8) return skip("is an unusual PNG");
    // Gray or RGB with an alpha channel: unfilter the rows and write the alpha as a soft mask.
    const bpp = ctype === 4 ? 2 : 4;
    const stride = width * bpp;
    const expected = height * (stride + 1);
    if (expected > this.caps.maxEntryBytes) throw new FormatError("LIMIT_EXCEEDED", "an image is larger than the entry limit when unpacked");
    this.decoded(expected);
    const raw = inflateUpTo(joined, expected);
    if (!raw || raw.length !== expected) return skip("is a PNG with damaged pixels");
    const px = unfilter(raw, width, height, bpp);
    if (!px) return skip("is a PNG with damaged pixels");
    const n = width * height;
    const body = new Uint8Array(n * (bpp - 1));
    const alpha = new Uint8Array(n);
    for (let i = 0; i < n; i++) {
      for (let c = 0; c < bpp - 1; c++) body[i * (bpp - 1) + c] = px[i * bpp + c];
      alpha[i] = px[i * bpp + bpp - 1];
    }
    const mask = this.alloc();
    this.set(mask, `<</Type /XObject /Subtype /Image /Width ${width} /Height ${height} /ColorSpace /DeviceGray /BitsPerComponent 8 /Filter /FlateDecode>>`, zlibSync(alpha));
    this.set(id, `<</Type /XObject /Subtype /Image /Width ${width} /Height ${height} /ColorSpace /${ctype === 4 ? "DeviceGray" : "DeviceRGB"} /BitsPerComponent 8 /SMask ${mask} 0 R /Filter /FlateDecode>>`, zlibSync(body));
    return img;
  }

  finish(title?: string): Uint8Array {
    if (this.pages.length === 0) this.newPage(612, 792);
    const fonts = Object.keys(FACES).map((key) => {
      const id = this.alloc();
      this.set(id, `<</Type /Font /Subtype /Type1 /BaseFont /${FACES[key as Face]} /Encoding /WinAnsiEncoding>>`);
      return `/${key} ${id} 0 R`;
    });
    const xobjects = [...this.images.values()].filter((i): i is Img => !!i).map((i) => `/${i.name} ${i.id} 0 R`);
    const resources = this.alloc();
    this.set(resources, `<</Font <<${fonts.join(" ")}>> /XObject <<${xobjects.join(" ")}>>>>`);
    const catalog = this.alloc();
    const tree = this.alloc();
    const info = this.alloc();
    const pageIds = this.pages.map((page) => {
      const content = this.alloc();
      this.set(content, "<</Filter /FlateDecode>>", zlibSync(latin1(page.ops.join("\n"))));
      const id = this.alloc();
      const annots = page.annots.map((a) => {
        const aid = this.alloc();
        this.set(aid, a);
        return `${aid} 0 R`;
      });
      this.set(id, `<</Type /Page /Parent ${tree} 0 R /MediaBox [0 0 ${num(page.width)} ${num(page.height)}] /Resources ${resources} 0 R /Contents ${content} 0 R${annots.length ? ` /Annots [${annots.join(" ")}]` : ""}>>`);
      return id;
    });
    this.set(tree, `<</Type /Pages /Kids [${pageIds.map((i) => `${i} 0 R`).join(" ")}] /Count ${pageIds.length}>>`);
    this.set(catalog, `<</Type /Catalog /Pages ${tree} 0 R>>`);
    if (title && [...title].length > 500) this.warnings.add("the title was cut to 500 characters");
    const titleHex = title ? `/Title <FEFF${[...title].slice(0, 500).map((c) => c.codePointAt(0)!).flatMap((cp) => (cp > 0xffff ? [0xd800 + ((cp - 0x10000) >> 10), 0xdc00 + ((cp - 0x10000) & 0x3ff)] : [cp])).map((u) => u.toString(16).padStart(4, "0")).join("")}> ` : "";
    this.set(info, `<<${titleHex}/Producer (agentic-formats) /CreationDate (D:20000101000000Z)>>`);

    const chunks: Uint8Array[] = [latin1("%PDF-1.5\n%\xe2\xe3\xcf\xd3\n")];
    const offsets: number[] = [];
    let at = chunks[0].length;
    for (let id = 1; id < this.objects.length; id++) {
      offsets[id] = at;
      const body = this.objects[id] ?? latin1("null");
      for (const c of [latin1(`${id} 0 obj\n`), body, latin1("\nendobj\n")]) {
        chunks.push(c);
        at += c.length;
      }
    }
    const xref = [`xref\n0 ${this.objects.length}\n0000000000 65535 f \n`, ...offsets.slice(1).map((o) => `${String(o).padStart(10, "0")} 00000 n \n`)].join("");
    chunks.push(latin1(`${xref}trailer\n<</Size ${this.objects.length} /Root ${catalog} 0 R /Info ${info} 0 R>>\nstartxref\n${at}\n%%EOF\n`));
    const out = new Uint8Array(chunks.reduce((n, c) => n + c.length, 0));
    chunks.reduce((o, c) => (out.set(c, o), o + c.length), 0);
    return out;
  }
}

/** A 64-bit hash (cyrb53 style) of bytes, to share two copies of one image. The length is part of the key too. */
function hash(b: Uint8Array): string {
  let h1 = 0xdeadbeef;
  let h2 = 0x41c6ce57;
  for (const x of b) {
    h1 = Math.imul(h1 ^ x, 2654435761);
    h2 = Math.imul(h2 ^ x, 1597334677);
  }
  h1 = Math.imul(h1 ^ (h1 >>> 16), 2246822507) ^ Math.imul(h2 ^ (h2 >>> 13), 3266489909);
  h2 = Math.imul(h2 ^ (h2 >>> 16), 2246822507) ^ Math.imul(h1 ^ (h1 >>> 13), 3266489909);
  return `${h1 >>> 0}.${h2 >>> 0}`;
}

/** Number of color components of a JPEG, from its first frame header. */
function jpegComponents(b: Uint8Array): number {
  const v = new DataView(b.buffer, b.byteOffset, b.byteLength);
  for (let i = 2; i + 9 < b.length && b[i] === 0xff; i += 2 + v.getUint16(i + 2)) {
    if (b[i + 1] >= 0xc0 && b[i + 1] <= 0xc2) return b[i + 4] === 8 ? b[i + 9] : 0;
  }
  return 0;
}

/** Inflate a zlib stream and stop with undefined as soon as it is longer than `size`, so a bomb never fills memory. */
function inflateUpTo(data: Uint8Array, size: number): Uint8Array | undefined {
  const out = new Uint8Array(size);
  let at = 0;
  try {
    const z = new Unzlib((chunk) => {
      if (at + chunk.length > size) throw new Error("too long");
      out.set(chunk, at);
      at += chunk.length;
    });
    // Push in small slices: the decoder unpacks a whole push before it calls back, so one big push would fill memory first.
    for (let i = 0; i < data.length || i === 0; i += 16384) z.push(data.subarray(i, i + 16384), i + 16384 >= data.length);
  } catch {
    return undefined;
  }
  return out.subarray(0, at);
}

function unfilter(raw: Uint8Array, width: number, height: number, bpp: number): Uint8Array | undefined {
  const stride = width * bpp;
  const out = new Uint8Array(stride * height);
  for (let y = 0; y < height; y++) {
    const kind = raw[y * (stride + 1)];
    if (kind > 4) return undefined;
    const src = y * (stride + 1) + 1;
    for (let x = 0; x < stride; x++) {
      const left = x >= bpp ? out[y * stride + x - bpp] : 0;
      const up = y > 0 ? out[(y - 1) * stride + x] : 0;
      const upLeft = y > 0 && x >= bpp ? out[(y - 1) * stride + x - bpp] : 0;
      let add = 0;
      if (kind === 1) add = left;
      else if (kind === 2) add = up;
      else if (kind === 3) add = (left + up) >> 1;
      else if (kind === 4) {
        const p = left + up - upLeft;
        const pa = Math.abs(p - left);
        const pb = Math.abs(p - up);
        const pc = Math.abs(p - upLeft);
        add = pa <= pb && pa <= pc ? left : pb <= pc ? up : upLeft;
      }
      out[y * stride + x] = (raw[src + x] + add) & 0xff;
    }
  }
  return out;
}

// ---------- text layout ----------

interface Seg {
  text: string;
  face: Face;
  width: number;
  underline?: boolean;
  strike?: boolean;
  href?: string;
}
type Line = Seg[];
interface Atom {
  text: string;
  face: Face;
  underline?: boolean;
  strike?: boolean;
  href?: string;
}

/** Runs as atoms: text already encoded, with the style a segment needs. A newline is a hard break. */
function atomsOf(pdf: Pdf, runs: TextRun[], bold: boolean): Atom[] {
  return runs.map((r) => {
    const link = r.href !== undefined ? safeHref(r.href) : undefined;
    const external = link !== undefined && /^(https?:|mailto:)/i.test(link) && /^[\x21-\x7e]+$/.test(link) ? link : undefined;
    if (r.href !== undefined && !external) pdf.warnings.add("a link that is not a web or mail address was written as plain text");
    if (r.marks?.includes("code") && (r.marks.includes("bold") || r.marks.includes("italic"))) pdf.warnings.add("bold and italic on code text are written as plain Courier");
    if (r.marks?.includes("superscript") || r.marks?.includes("subscript")) pdf.warnings.add("superscript and subscript are written as plain text");
    return { text: r.text === "\n" ? "\n" : pdf.encode(r.text), face: faceOf(r.marks, bold), underline: r.marks?.includes("underline") || !!external, strike: r.marks?.includes("strike"), href: external };
  });
}

const sameStyle = (s: Seg, a: Atom): boolean => s.face === a.face && !!s.underline === !!a.underline && !!s.strike === !!a.strike && s.href === a.href;

/** Break atoms into lines no wider than `width` points. A word wider than a line splits. */
function wrap(atoms: Atom[], size: number, width: number): Line[] {
  const lines: Line[] = [];
  let line: Line = [];
  let used = 0;
  const push = (text: string, a: Atom, w: number): void => {
    const last = line[line.length - 1];
    if (last && sameStyle(last, a)) {
      last.text += text;
      last.width += w;
    } else line.push({ text, face: a.face, width: w, underline: a.underline, strike: a.strike, href: a.href });
    used += w;
  };
  const wrapAt = (): void => {
    lines.push(line);
    line = [];
    used = 0;
  };
  const w = Math.max(width, size);
  const widthOf = (text: string, face: Face): number => [...text].reduce((n, c) => n + (charWidth(c.charCodeAt(0), face) * size) / 1000, 0);
  // A word is the pieces between spaces, so a change of style inside a word is not a place to break.
  type Piece = { text: string; atom: Atom; w: number };
  let word: Piece[] = [];
  const place = (): void => {
    if (word.length === 0) return;
    const total = word.reduce((n, p) => n + p.w, 0);
    if (used > 0 && used + total > w) wrapAt();
    if (total <= w) for (const p of word) push(p.text, p.atom, p.w);
    else {
      for (const p of word) {
        for (const c of p.text) {
          const cw = widthOf(c, p.atom.face);
          if (used > 0 && used + cw > w) wrapAt();
          push(c, p.atom, cw);
        }
      }
    }
    word = [];
  };
  for (const a of atoms) {
    if (a.text === "\n") {
      place();
      wrapAt();
      continue;
    }
    for (const token of a.text.match(/ +|[^ ]+/g) ?? []) {
      const tw = widthOf(token, a.face);
      if (token[0] !== " ") {
        word.push({ text: token, atom: a, w: tw });
        continue;
      }
      place();
      if ((used > 0 || lines.length === 0) && used + tw <= w) push(token, a, tw);
    }
  }
  place();
  if (line.length || lines.length === 0) lines.push(line);
  return lines;
}

/** Draw a line with its left edge at `x` and its baseline at `y`. */
function drawLine(page: Page, line: Line, size: number, x: number, y: number): void {
  let cx = x;
  for (const seg of line) {
    if (seg.text.trim()) {
      page.ops.push(`BT /${seg.face} ${num(size)} Tf ${seg.href ? "0 0 0.8" : "0 0 0"} rg ${num(cx)} ${num(y)} Td ${lit(seg.text)} Tj ET`);
    }
    if (seg.underline) page.ops.push(`0.5 w ${seg.href ? "0 0 0.8" : "0 0 0"} RG ${num(cx)} ${num(y - size * 0.12)} m ${num(cx + seg.width)} ${num(y - size * 0.12)} l S`);
    if (seg.strike) page.ops.push(`0.5 w 0 0 0 RG ${num(cx)} ${num(y + size * 0.3)} m ${num(cx + seg.width)} ${num(y + size * 0.3)} l S`);
    if (seg.href) page.annots.push(`<</Type /Annot /Subtype /Link /Rect [${num(cx)} ${num(y - size * 0.2)} ${num(cx + seg.width)} ${num(y + size * 0.85)}] /Border [0 0 0] /A <</S /URI /URI ${lit(seg.href)}>>>>`);
    cx += seg.width;
  }
}

// ---------- slides ----------

const fillOp = (shape: { fill?: string; stroke?: string }, dflt?: string): string => {
  const stroke = shape.stroke ?? dflt;
  return `${shape.fill ? `${rgb(shape.fill)} rg ` : ""}${stroke ? `${rgb(stroke)} RG 1 w ` : ""}`;
};
const paint = (shape: { fill?: string; stroke?: string }, dflt?: string): string => ((shape.fill ? 1 : 0) + ((shape.stroke ?? dflt) ? 2 : 0) === 3 ? "B" : shape.fill ? "f" : (shape.stroke ?? dflt) ? "S" : "n");

function writeDeck(deck: Deck, pdf: Pdf): void {
  checkDeck(deck, pdf.caps);
  deck.slides.forEach((slide, index) => {
    const page = pdf.newPage(deck.width, deck.height);
    const top = deck.height;
    for (const shape of slide.shapes) {
      const { x, y, width, height } = shape.frame;
      const py = top - y - height;
      if (shape.type === "shape") {
        if (shape.kind === "line") page.ops.push(`${fillOp({ stroke: shape.stroke }, "#000000")}${num(x)} ${num(top - y)} m ${num(x + width)} ${num(py)} l S`);
        else if (shape.kind === "rectangle") page.ops.push(`${fillOp(shape)}${num(x)} ${num(py)} ${num(width)} ${num(height)} re ${paint(shape)}`);
        else {
          const k = 0.5523;
          const [cx, cy, rx, ry] = [x + width / 2, py + height / 2, width / 2, height / 2];
          page.ops.push(
            `${fillOp(shape)}${num(cx + rx)} ${num(cy)} m ${num(cx + rx)} ${num(cy + k * ry)} ${num(cx + k * rx)} ${num(cy + ry)} ${num(cx)} ${num(cy + ry)} c ${num(cx - k * rx)} ${num(cy + ry)} ${num(cx - rx)} ${num(cy + k * ry)} ${num(cx - rx)} ${num(cy)} c ${num(cx - rx)} ${num(cy - k * ry)} ${num(cx - k * rx)} ${num(cy - ry)} ${num(cx)} ${num(cy - ry)} c ${num(cx + k * rx)} ${num(cy - ry)} ${num(cx + rx)} ${num(cy - k * ry)} ${num(cx + rx)} ${num(cy)} c h ${paint(shape)}`,
          );
        }
      } else if (shape.type === "picture") {
        if (shape.alt) pdf.warnings.add("image descriptions are not written to a PDF");
        const img = pdf.image(shape.data);
        if (img) page.ops.push(`q ${num(width)} 0 0 ${num(height)} ${num(x)} ${num(py)} cm /${img.name} Do Q`);
      } else {
        // The model holds no font size: take the largest of 18 down to 8 points that fits the frame.
        const atoms = shape.paragraphs.map((runs) => atomsOf(pdf, runs, false));
        let size = 18;
        let lines: Line[] = [];
        for (; size >= 8; size--) {
          lines = atoms.flatMap((a) => wrap(a, size, width));
          if (lines.length * size * 1.2 <= height + 0.01) break;
        }
        size = Math.max(size, 8);
        const fit = Math.max(0, Math.floor((height + 0.01) / (size * 1.2)));
        if (lines.length > fit) {
          pdf.warnings.add(`text was cut to fit its box on slide ${index + 1}`);
          lines = lines.slice(0, fit);
        }
        lines.forEach((line, i) => drawLine(page, line, size, x, top - y - size * 1.2 * i - size));
      }
    }
    if (slide.notes) pdf.warnings.add("speaker notes are not written to a PDF");
  });
}

// ---------- documents ----------

const PAGE_W = 612;
const PAGE_H = 792;
const MARGIN = 72;
const CONTENT_H = PAGE_H - 2 * MARGIN;
const BODY = 11;
const HEADING_SIZES = [22, 18, 15, 13, 12, 11];
const CELL_PAD = 4;

interface Item {
  h: number;
  /** Draw with the top left corner at `x`, and `top` points down from the top of the page. */
  draw: (page: Page, x: number, top: number) => void;
}
interface Ctx {
  pdf: Pdf;
  notes: Line[][];
  depth: number;
}

/** Add the items of `from` to `to`. Spreading a long list into `push` overflows the stack. */
const append = (to: Item[], from: Item[]): void => {
  for (const item of from) to.push(item);
};
const gap = (h: number): Item => ({ h, draw: () => {} });

function lineItem(line: Line, size: number, dx: number): Item {
  return {
    h: size * 1.3,
    draw: (page, x, top) => {
      const y = PAGE_H - top - size;
      drawLine(page, line, size, x + dx, y);
    },
  };
}

function imageItem(ctx: Ctx, node: Extract<Inline, { type: "image" }>, width: number, dx: number): Item | undefined {
  if (node.alt) ctx.pdf.warnings.add("image descriptions are not written to a PDF");
  const img = ctx.pdf.image(node.data);
  if (!img) return undefined;
  const [w0, h0] = frameSize(node.width, node.height, sniffImage(node.data)!);
  const scale = Math.min(1, (width - dx) / w0, (CONTENT_H - 2 * CELL_PAD) / h0);
  const [w, h] = [Math.max(1, w0 * scale), Math.max(1, h0 * scale)];
  return { h: h + 4, draw: (page, x, top) => page.ops.push(`q ${num(w)} 0 0 ${num(h)} ${num(x + dx)} ${num(PAGE_H - top - h)} cm /${img.name} Do Q`) };
}

/** Items for inline content. */
function inlineItems(ctx: Ctx, content: Inline[], width: number, dx: number, size: number, bold: boolean): Item[] {
  const items: Item[] = [];
  let runs: TextRun[] = [];
  const flush = (): void => {
    if (runs.length === 0) return;
    for (const line of wrap(atomsOf(ctx.pdf, runs, bold), size, width - dx)) items.push(lineItem(line, size, dx));
    runs = [];
  };
  for (const node of content) {
    if (node.type === "text") runs.push(node);
    else if (node.type === "break") runs.push({ type: "text", text: "\n" });
    else if (node.type === "footnote") {
      ctx.notes.push(wrap(atomsOf(ctx.pdf, [{ type: "text", text: `${ctx.notes.length + 1}. ` }, ...(node.content.map((n) => (n.type === "break" ? { type: "text", text: " " } : n)) as TextRun[])], false), 9, width));
      runs.push({ type: "text", text: `[${ctx.notes.length}]` });
    } else {
      flush();
      const item = imageItem(ctx, node, width, dx);
      if (item) items.push(item);
    }
  }
  flush();
  // A paragraph with no text still takes a line, so an empty paragraph keeps its gap.
  if (items.length === 0 && content.length === 0) items.push(lineItem([], size, dx));
  return items;
}

const hasRich = (block: Block): boolean =>
  block.type === "table"
    ? block.rows.some((r) => r.some((c) => c.some(hasRich)))
    : block.type === "list"
      ? hasRichList(block)
      : block.content.some((n) => n.type === "image" || n.type === "footnote");
const hasRichList = (list: List): boolean => list.items.some((i) => i.content.some((n) => n.type === "image" || n.type === "footnote") || (i.sublist ? hasRichList(i.sublist) : false));

const plainText = (blocks: Block[]): string =>
  blocks
    .map((b) => (b.type === "table" ? b.rows.map((r) => r.map((c) => plainText(c)).join(" | ")).join("\n") : b.type === "list" ? plainList(b) : b.content.map((n) => (n.type === "text" ? n.text : n.type === "break" ? "\n" : "")).join("")))
    .join("\n");
const plainList = (list: List): string => list.items.map((i) => `${i.content.map((n) => (n.type === "text" ? n.text : "")).join("")}${i.sublist ? `\n${plainList(i.sublist)}` : ""}`).join("\n");

/** The most a list can indent: past this the text has no room, so deeper levels share the last indent. */
const maxIndent = (width: number): number => Math.max(0, Math.floor(width / 2 / 18) * 18);

function listItems(ctx: Ctx, list: List, width: number, dx: number): Item[] {
  const out: Item[] = [];
  if (dx > maxIndent(width)) {
    ctx.pdf.warnings.add("a list nested too deep was written at the deepest indent that fits");
    dx = maxIndent(width);
  }
  list.items.forEach((item, i) => {
    const marker = list.ordered ? `${i + 1}.` : "\x95";
    const inner = inlineItems(ctx, item.content, width, dx + 18, BODY, false);
    const drawMarker = (page: Page, x: number, top: number): void => drawLine(page, [{ text: marker, face: "F1", width: 0 }], BODY, x + dx, PAGE_H - top - BODY);
    const lead = inner[0];
    if (lead) inner[0] = { h: lead.h, draw: (page, x, top) => (drawMarker(page, x, top), lead.draw(page, x, top)) };
    else inner.push({ h: BODY * 1.3, draw: drawMarker });
    for (const n of inner) out.push(n);
    if (item.sublist) for (const n of listItems(ctx, item.sublist, width, dx + 18)) out.push(n);
  });
  return out;
}

function tableItems(ctx: Ctx, block: Extract<Block, { type: "table" }>, width: number): Item[] {
  const cols = block.rows.reduce((n, r) => Math.max(n, r.length), 0);
  if (cols === 0) return [];
  const colW = width / cols;
  if (colW - 2 * CELL_PAD < BODY * 2) ctx.pdf.warnings.add("a table has more columns than fit: text in its cells runs past them");
  const out: Item[] = [];
  block.rows.forEach((row, r) => {
    const header = !!block.header && r === 0;
    const cells = row.map((cell) => {
      ctx.depth++;
      const items = flow(ctx, cell, colW - 2 * CELL_PAD, 0, header);
      ctx.depth--;
      return items;
    });
    let h = cells.reduce((m, c) => Math.max(m, c.reduce((n, i) => n + i.h, 0)), 0) + 2 * CELL_PAD;
    if (h > CONTENT_H) {
      ctx.pdf.warnings.add("a table row taller than a page was cut to fit");
      h = CONTENT_H;
    }
    out.push({
      h,
      draw: (page, x, top) => {
        for (let c = 0; c < cells.length; c++) {
          const cx = x + c * colW;
          page.ops.push(`${header ? "0.92 g " : ""}0.5 w 0 0 0 RG ${num(cx)} ${num(PAGE_H - top - h)} ${num(colW)} ${num(h)} re ${header ? "B" : "S"}`);
          let at = top + CELL_PAD;
          for (const item of cells[c]) {
            if (at + item.h > top + h) break;
            item.draw(page, cx + CELL_PAD, at);
            at += item.h;
          }
        }
      },
    });
  });
  return out;
}

function flow(ctx: Ctx, blocks: Block[], width: number, dx: number, bold = false): Item[] {
  const items: Item[] = [];
  for (const block of blocks) {
    if (block.type === "heading") {
      const size = HEADING_SIZES[headingLevel(block.level) - 1];
      items.push(gap(size * 0.4));
      append(items, inlineItems(ctx, block.content, width, dx, size, true));
      items.push(gap(3));
    } else if (block.type === "paragraph") {
      append(items, inlineItems(ctx, block.content, width, dx, BODY, bold));
      items.push(gap(6));
    } else if (block.type === "list") {
      append(items, listItems(ctx, block, width, dx));
      items.push(gap(6));
    }
    else if (ctx.depth > 0) {
      ctx.pdf.warnings.add("a table inside a table was written as lines of text");
      if (hasRich(block)) ctx.pdf.warnings.add("an image or footnote inside a nested table was not written");
      append(items, inlineItems(ctx, [{ type: "text", text: plainText([block]).replace(/\n/g, " / ") }], width, dx, BODY, bold));
    } else {
      append(items, tableItems(ctx, block, width));
      items.push(gap(8));
    }
  }
  return items;
}

function checkBlocks(blocks: unknown, caps: Limits, depth: number, cells: { n: number; work: number }): void {
  checkList(blocks, "the blocks of a document");
  if (depth > Math.floor(caps.maxXmlDepth / 3)) throw new FormatError("INVALID_MODEL", "lists or tables nest too deep");
  // A block, an item, or a node that the model holds in many places is laid out in each place, so each visit counts.
  const tick = (): void => {
    if (++cells.work > caps.maxCells) throw new FormatError("LIMIT_EXCEEDED", `the document has more than ${caps.maxCells} blocks, items, and nodes`);
  };
  const inline = (content: unknown): void => {
    checkList(content, "the content of a block");
    for (const node of content as Inline[]) {
      tick();
      if (typeof node !== "object" || node === null || !["text", "break", "image", "footnote"].includes(node.type)) throw new FormatError("INVALID_MODEL", "an inline node is not text, a break, an image, or a footnote");
      if (node.type === "text") checkRun(node);
      if (node.type === "image") sniffImage(node.data);
      if (node.type === "footnote") {
        checkList(node.content, "the content of a footnote");
        for (const n of node.content as Inline[]) {
          tick();
          if (n?.type !== "text" && n?.type !== "break") throw new FormatError("INVALID_MODEL", "a footnote holds text and line breaks only");
          if (n.type === "text") checkRun(n);
        }
      }
    }
  };
  const list = (l: List, d: number): void => {
    checkList(l?.items, "the items of a list");
    if (d > Math.floor(caps.maxXmlDepth / 3)) throw new FormatError("INVALID_MODEL", "lists or tables nest too deep");
    for (const item of l.items) {
      tick();
      inline(item?.content);
      if (item.sublist) list(item.sublist, d + 1);
    }
  };
  for (const block of blocks as Block[]) {
    tick();
    if (typeof block !== "object" || block === null) throw new FormatError("INVALID_MODEL", "a block must be an object");
    if (block.type === "heading") {
      headingLevel(block.level);
      inline(block.content);
    } else if (block.type === "paragraph") inline(block.content);
    else if (block.type === "list") list(block, depth);
    else if (block.type === "table") {
      checkList(block.rows, "the rows of a table");
      // The table is drawn with a box for each row and column, so the grid counts, not only the cells.
      const cols = block.rows.reduce((m, r) => (Array.isArray(r) ? Math.max(m, r.length) : m), 0);
      if (block.rows.length * cols > caps.maxCells) throw new FormatError("LIMIT_EXCEEDED", `the table has more than ${caps.maxCells} cells`);
      for (const row of block.rows) {
        tick();
        checkList(row, "a table row");
        for (const cell of row) {
          if (++cells.n > caps.maxCells) throw new FormatError("LIMIT_EXCEEDED", `the table has more than ${caps.maxCells} cells`);
          checkBlocks(cell, caps, depth + 1, cells);
        }
      }
    } else throw new FormatError("INVALID_MODEL", "a block is not a heading, a paragraph, a list, or a table");
  }
}

function writeText(doc: TextDoc, pdf: Pdf): void {
  checkBlocks(doc?.blocks, pdf.caps, 0, { n: 0, work: 0 });
  if (doc.title !== undefined && typeof doc.title !== "string") throw new FormatError("INVALID_MODEL", "the title of a document must be a string");
  // The text is charged as it is encoded. Image bytes are not text: they count against the output size.
  if (textLength(doc.blocks) > pdf.caps.maxTextChars + pdf.caps.maxTotalBytes) throw new FormatError("LIMIT_EXCEEDED", "the document is larger than the limits allow");
  const ctx: Ctx = { pdf, notes: [], depth: 0 };
  const items = flow(ctx, doc.blocks, PAGE_W - 2 * MARGIN, 0);
  if (ctx.notes.length) {
    items.push(gap(10));
    for (const note of ctx.notes) for (const line of note) items.push(lineItem(line, 9, 0));
  }
  let page = pdf.newPage(PAGE_W, PAGE_H);
  let top = MARGIN;
  for (const item of items) {
    if (top + item.h > PAGE_H - MARGIN && top > MARGIN) {
      page = pdf.newPage(PAGE_W, PAGE_H);
      top = MARGIN;
    }
    item.draw(page, MARGIN, top);
    top += item.h;
  }
}

/**
 * Write a deck (one page per slide) or a text document (Letter pages) as a PDF. Text uses the standard
 * PDF fonts, so characters outside Latin-1 and Windows punctuation are written as "?" and reported.
 */
export function writePdf(model: Deck | TextDoc, limits?: Partial<Limits>): WriteResult {
  const pdf = new Pdf(withLimits(limits));
  if (typeof model !== "object" || model === null) throw new FormatError("INVALID_MODEL", "a PDF is written from a deck or a text document");
  const isDeck = "slides" in model;
  if (isDeck) writeDeck(model, pdf);
  else writeText(model, pdf);
  return { bytes: pdf.finish(isDeck ? undefined : model.title), warnings: [...pdf.warnings] };
}
