import { closeSync, existsSync, openSync, readSync, realpathSync, statSync, writeFileSync } from "node:fs";
import { basename, dirname, join, resolve, sep } from "node:path";
import { FormatError, withLimits, type Limits } from "./limits.js";

/** The caller asked for something the tool refuses: a bad path, an overwrite, an unknown extension. */
export class UsageError extends Error {
  constructor(message: string) {
    super(message);
    this.name = "UsageError";
  }
}

/**
 * Resolve a path inside `root`, and refuse one that leaves it, by `..` or by a symbolic link.
 * Without a root, the path is only made absolute.
 */
export function resolveInside(path: string, root?: string): string {
  const absolute = resolve(root ?? ".", path);
  if (!root) return absolute;
  let existing = absolute;
  const rest: string[] = [];
  while (!existsSync(existing) && dirname(existing) !== existing) {
    rest.unshift(basename(existing));
    existing = dirname(existing);
  }
  const real = join(realpathSync(existing), ...rest);
  const realRoot = realpathSync(root);
  if (real !== realRoot && !real.startsWith(realRoot + sep)) throw new UsageError(`"${path}" is outside the allowed folder`);
  return real;
}

/** Read a file after its size is checked, so a huge file is never loaded. */
export function readInput(path: string, limits?: Partial<Limits>): Uint8Array {
  const { maxInputBytes } = withLimits(limits);
  const stats = statSync(path);
  if (!stats.isFile()) throw new UsageError(`"${path}" is not a file`);
  if (stats.size > maxInputBytes) throw new FormatError("LIMIT_EXCEEDED", `the file is larger than ${maxInputBytes} bytes`);
  const fd = openSync(path, "r");
  try {
    const bytes = new Uint8Array(stats.size);
    let at = 0;
    while (at < bytes.length) {
      const n = readSync(fd, bytes, at, bytes.length - at, at);
      if (n === 0) break;
      at += n;
    }
    return bytes.subarray(0, at);
  } finally {
    closeSync(fd);
  }
}

/** Write a file. It refuses to replace one unless `overwrite` is true, and never replaces the input. */
export function writeOutput(path: string, bytes: Uint8Array, overwrite: boolean, input?: string): void {
  if (input !== undefined && existsSync(path) && realpathSync(path) === realpathSync(input)) throw new UsageError("the output is the input file");
  try {
    writeFileSync(path, bytes, { flag: overwrite ? "w" : "wx" });
  } catch (error) {
    if ((error as NodeJS.ErrnoException).code === "EEXIST") throw new UsageError(`"${path}" exists; pass --force to replace it`);
    throw error;
  }
}

/** The text for an error the caller can fix, or undefined for any other error. */
export function describeError(error: unknown): string | undefined {
  if (error instanceof FormatError) return `${error.code}: ${error.message}`;
  if (error instanceof UsageError) return `UsageError: ${error.message}`;
  if ((error as NodeJS.ErrnoException).code === "ENOENT") return "UsageError: file not found";
  return undefined;
}
