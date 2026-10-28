import { strToU8, unzipSync, zipSync, type Zippable } from "fflate";
import { FormatError, type Limits } from "./limits.js";

const FIXED_DATE = new Date(2000, 0, 1);

/** A zip entry name that stays inside the package: no `..`, no root, no drive, no backslash, no control character. */
export function isSafeEntryName(name: string): boolean {
  if (!name || name.length > 255 || /[\\\u0000-\u001f]/.test(name) || /^([a-z]:|\/)/i.test(name)) return false;
  return name.replace(/\/$/, "").split("/").every((part) => part !== "" && part !== "." && part !== "..");
}

/**
 * Unpack the entries `wanted` accepts. Every entry in the file is checked, wanted or not, and the
 * checks run on the sizes in the zip directory before any entry is inflated. A file that fails a
 * check is rejected whole.
 */
export function readZip(bytes: Uint8Array, limits: Limits, kind: string, wanted: (name: string) => boolean): Record<string, Uint8Array> {
  if (bytes.length > limits.maxInputBytes) throw new FormatError("LIMIT_EXCEEDED", `the ${kind} file is larger than ${limits.maxInputBytes} bytes`);
  const seen = new Set<string>();
  let total = 0;
  let violation: FormatError | undefined;
  const fail = (code: "INVALID_FILE" | "LIMIT_EXCEEDED", message: string): false => {
    violation ??= new FormatError(code, message);
    return false;
  };
  let files: Record<string, Uint8Array>;
  try {
    files = unzipSync(bytes, {
      filter: ({ name, size, originalSize }) => {
        if (violation) return false;
        if (!isSafeEntryName(name)) return fail("INVALID_FILE", `the ${kind} file has an unsafe entry name`);
        const key = name.toLowerCase();
        if (seen.has(key)) return fail("INVALID_FILE", `the ${kind} file lists the entry "${name}" twice`);
        seen.add(key);
        if (seen.size > limits.maxEntries) return fail("LIMIT_EXCEEDED", `the ${kind} file has more than ${limits.maxEntries} entries`);
        total += originalSize;
        if (originalSize > limits.maxEntryBytes) return fail("LIMIT_EXCEEDED", `the ${kind} file has an entry larger than ${limits.maxEntryBytes} bytes`);
        if (total > limits.maxTotalBytes) return fail("LIMIT_EXCEEDED", `the ${kind} file unpacks to more than ${limits.maxTotalBytes} bytes`);
        if (originalSize > limits.ratioFloorBytes && originalSize / Math.max(size, 1) > limits.maxRatio) {
          return fail("LIMIT_EXCEEDED", `the ${kind} file has an entry that packs more than ${limits.maxRatio} to 1`);
        }
        return wanted(name);
      },
    });
  } catch {
    // fflate stops at the first bad entry; a limit hit inside the filter wins over its error.
    throw violation ?? new FormatError("INVALID_FILE", `the file is not a valid ${kind} (zip) package`);
  }
  if (violation) throw violation;
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
