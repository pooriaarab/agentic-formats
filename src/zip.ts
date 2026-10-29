import { Inflate, strFromU8, strToU8, zipSync, type Zippable } from "fflate";
import { FormatError, type Limits } from "./limits.js";

const FIXED_DATE = new Date(2000, 0, 1);

/** A zip entry name that stays inside the package: no `..`, no root, no drive, no backslash, no control character. */
export function isSafeEntryName(name: string): boolean {
  if (!name || name.length > 255 || /[\\\u0000-\u001f]/.test(name) || /^([a-z]:|\/)/i.test(name)) return false;
  return name.replace(/\/$/, "").split("/").every((part) => part !== "" && part !== "." && part !== "..");
}

const EOCD = 0x06054b50;
const CENTRAL = 0x02014b50;
const LOCAL = 0x04034b50;
/** Compressed bytes fed to the inflater at once. Deflate grows at most about 1000 to 1, so a chunk stays small after it expands. */
const CHUNK = 16 * 1024;

interface DirectoryEntry {
  name: string;
  method: number;
  packedSize: number;
  size: number;
  localOffset: number;
}

/** The entries in the central directory. The sizes here are claims, not facts: `inflateEntry` checks them. */
function readDirectory(bytes: Uint8Array, limits: Limits, kind: string): DirectoryEntry[] {
  const notZip = () => new FormatError("INVALID_FILE", `the file is not a valid ${kind} (zip) package`);
  const view = new DataView(bytes.buffer, bytes.byteOffset, bytes.byteLength);
  let end = -1;
  for (let at = bytes.length - 22; at >= Math.max(0, bytes.length - 22 - 0xffff); at--) {
    if (view.getUint32(at, true) === EOCD) {
      end = at;
      break;
    }
  }
  if (end < 0) throw notZip();
  const count = view.getUint16(end + 10, true);
  const size = view.getUint32(end + 12, true);
  let at = view.getUint32(end + 16, true);
  if (count === 0xffff || size === 0xffffffff || at === 0xffffffff) throw notZip();
  if (count > limits.maxEntries) throw new FormatError("LIMIT_EXCEEDED", `the ${kind} file has more than ${limits.maxEntries} entries`);
  if (at + size > end) throw notZip();
  const entries: DirectoryEntry[] = [];
  for (let n = 0; n < count; n++) {
    if (at + 46 > end || view.getUint32(at, true) !== CENTRAL) throw notZip();
    const flags = view.getUint16(at + 8, true);
    const nameLength = view.getUint16(at + 28, true);
    const extraLength = view.getUint16(at + 30, true);
    const commentLength = view.getUint16(at + 32, true);
    if (flags & 1 || at + 46 + nameLength + extraLength + commentLength > end) throw notZip();
    const nameBytes = bytes.subarray(at + 46, at + 46 + nameLength);
    entries.push({
      name: flags & 0x800 ? strFromU8(nameBytes) : strFromU8(nameBytes, true),
      method: view.getUint16(at + 10, true),
      packedSize: view.getUint32(at + 20, true),
      size: view.getUint32(at + 24, true),
      localOffset: view.getUint32(at + 42, true),
    });
    at += 46 + nameLength + extraLength + commentLength;
  }
  return entries;
}

/**
 * Unpack one entry and count the bytes that really come out. The size in the directory is only a
 * claim: output past it, or short of it, rejects the file. The count stops the inflater at the first
 * chunk past the claim, so a small claim over a huge stream costs one chunk, not the whole stream.
 */
function inflateEntry(bytes: Uint8Array, entry: DirectoryEntry, kind: string): Uint8Array {
  const view = new DataView(bytes.buffer, bytes.byteOffset, bytes.byteLength);
  const { localOffset, packedSize, size, name } = entry;
  const bad = () => new FormatError("INVALID_FILE", `the ${kind} file has a broken entry "${name}"`);
  if (localOffset + 30 > bytes.length || view.getUint32(localOffset, true) !== LOCAL) throw bad();
  const dataStart = localOffset + 30 + view.getUint16(localOffset + 26, true) + view.getUint16(localOffset + 28, true);
  if (dataStart + packedSize > bytes.length) throw bad();
  const packed = bytes.subarray(dataStart, dataStart + packedSize);
  if (entry.method === 0) {
    if (packedSize !== size) throw bad();
    return packed.slice();
  }
  if (entry.method !== 8) throw new FormatError("INVALID_FILE", `the ${kind} file uses a zip method that this package does not read`);
  const chunks: Uint8Array[] = [];
  let real = 0;
  try {
    const inflater = new Inflate((chunk) => {
      real += chunk.length;
      if (real > size) throw new FormatError("INVALID_FILE", `the ${kind} file has an entry larger than its directory says`);
      chunks.push(chunk);
    });
    for (let from = 0; ; from += CHUNK) {
      const last = from + CHUNK >= packed.length;
      inflater.push(packed.subarray(from, from + CHUNK), last);
      if (last) break;
    }
  } catch (error) {
    throw error instanceof FormatError ? error : bad();
  }
  if (real !== size) throw new FormatError("INVALID_FILE", `the ${kind} file has an entry smaller than its directory says`);
  const out = new Uint8Array(real);
  let at = 0;
  for (const chunk of chunks) (out.set(chunk, at), (at += chunk.length));
  return out;
}

/**
 * Unpack the entries `wanted` accepts. Every entry in the directory is checked, wanted or not. The
 * caps on the claimed sizes run before any entry is inflated; then each entry is counted as it comes
 * out and must match its claim exactly, so the claims that passed the caps are the real sizes. A file that
 * fails a check is rejected whole. The result has no prototype, so a hostile name such as
 * `constructor` is just an absent key.
 */
export function readZip(bytes: Uint8Array, limits: Limits, kind: string, wanted: (name: string) => boolean): Record<string, Uint8Array> {
  if (bytes.length > limits.maxInputBytes) throw new FormatError("LIMIT_EXCEEDED", `the ${kind} file is larger than ${limits.maxInputBytes} bytes`);
  const seen = new Set<string>();
  let claimed = 0;
  const selected: DirectoryEntry[] = [];
  for (const entry of readDirectory(bytes, limits, kind)) {
    const { name, size, packedSize } = entry;
    if (!isSafeEntryName(name)) throw new FormatError("INVALID_FILE", `the ${kind} file has an unsafe entry name`);
    const key = name.toLowerCase();
    if (seen.has(key)) throw new FormatError("INVALID_FILE", `the ${kind} file lists the entry "${name}" twice`);
    seen.add(key);
    claimed += size;
    if (size > limits.maxEntryBytes) throw new FormatError("LIMIT_EXCEEDED", `the ${kind} file has an entry larger than ${limits.maxEntryBytes} bytes`);
    if (claimed > limits.maxTotalBytes) throw new FormatError("LIMIT_EXCEEDED", `the ${kind} file unpacks to more than ${limits.maxTotalBytes} bytes`);
    if (size > limits.ratioFloorBytes && size / Math.max(packedSize, 1) > limits.maxRatio) {
      throw new FormatError("LIMIT_EXCEEDED", `the ${kind} file has an entry that packs more than ${limits.maxRatio} to 1`);
    }
    if (wanted(name)) selected.push(entry);
  }
  const files = Object.create(null) as Record<string, Uint8Array>;
  for (const entry of selected) files[entry.name] = inflateEntry(bytes, entry, kind);
  return files;
}

export interface ZipEntry {
  name: string;
  data: Uint8Array | string;
  /** Store without compression. ODF needs this for `mimetype`. */
  store?: boolean;
}

/**
 * Pack entries in the order given, with a fixed date, so the same model gives the same bytes.
 * The entry and total size caps apply on the way out too, so a writer cannot make a file that a
 * reader with the same limits would refuse.
 */
export function writeZip(entries: ZipEntry[], limits: Limits): Uint8Array {
  const zippable: Zippable = {};
  let total = 0;
  for (const { name, data, store } of entries) {
    if (!isSafeEntryName(name)) throw new FormatError("INVALID_MODEL", `"${name}" is not a safe entry name`);
    const bytes = typeof data === "string" ? strToU8(data) : data;
    if (bytes.length > limits.maxEntryBytes) throw new FormatError("LIMIT_EXCEEDED", `"${name}" is larger than ${limits.maxEntryBytes} bytes`);
    total += bytes.length;
    if (total > limits.maxTotalBytes) throw new FormatError("LIMIT_EXCEEDED", `the export is larger than ${limits.maxTotalBytes} bytes`);
    zippable[name] = [bytes, { level: store ? 0 : 6, mtime: FIXED_DATE }];
  }
  const out = zipSync(zippable);
  if (out.length > limits.maxInputBytes) throw new FormatError("LIMIT_EXCEEDED", `the export is larger than ${limits.maxInputBytes} bytes`);
  return out;
}
