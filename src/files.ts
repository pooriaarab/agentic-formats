import { closeSync, constants, existsSync, fstatSync, ftruncateSync, lstatSync, openSync, readSync, realpathSync, statSync, writeSync } from "node:fs";
import { homedir } from "node:os";
import { basename, dirname, join, parse, resolve, sep } from "node:path";
import { FormatError, withLimits, type Limits } from "./limits.js";

/** The caller asked for something the tool refuses: a bad path, an overwrite, an unknown extension. */
export class UsageError extends Error {
  constructor(message: string) {
    super(message);
    this.name = "UsageError";
  }
}

const exists = (path: string): boolean => {
  try {
    lstatSync(path);
    return true;
  } catch {
    return false;
  }
};

/** Which paths a tool may touch: one folder and everything under it, or any path the process can reach. */
export type PathPolicy = { folder: string } | { anyPath: true };

/**
 * The policy the MCP server runs with. With no variable set, the folder is the working folder.
 * `AGENTIC_FORMATS_ROOT` names another folder. `AGENTIC_FORMATS_UNSAFE_ANY_PATH=1` allows any path.
 * An empty variable, a flag other than `1`, both variables, or a root that is not a folder is an
 * error: a typo in a config file never turns the guard off.
 */
export function pathPolicyFromEnv(env: NodeJS.ProcessEnv, cwd: string): PathPolicy {
  const root = env.AGENTIC_FORMATS_ROOT;
  const anyPath = env.AGENTIC_FORMATS_UNSAFE_ANY_PATH;
  if (root !== undefined && root.trim() === "") throw new UsageError("AGENTIC_FORMATS_ROOT is set but empty; set it to a folder, or unset it to use the working folder");
  if (anyPath !== undefined && anyPath !== "1") throw new UsageError('AGENTIC_FORMATS_UNSAFE_ANY_PATH must be "1" to allow any path; unset it otherwise');
  if (anyPath === "1" && root !== undefined) throw new UsageError("set AGENTIC_FORMATS_ROOT or AGENTIC_FORMATS_UNSAFE_ANY_PATH, not both");
  if (anyPath === "1") return { anyPath: true };
  const folder = resolve(cwd, root ?? ".");
  if (!existsSync(folder) || !statSync(folder).isDirectory()) throw new UsageError(`"${folder}" is not a folder`);
  // Some clients start a server in `/` or in the home folder. As a default that is the whole disk or every personal file, so a person must name a folder.
  // Compare real paths: the working folder is always a real path, and a home folder may go through a link.
  const real = realpathSync(folder);
  const realHome = (() => {
    try {
      return realpathSync(homedir());
    } catch {
      return resolve(homedir());
    }
  })();
  if (root === undefined && (real === parse(real).root || real === realHome)) {
    throw new UsageError(`the working folder "${folder}" is the filesystem root or the home folder; set AGENTIC_FORMATS_ROOT to the folder the tools may use`);
  }
  return { folder };
}

/**
 * Resolve a path under a policy. With a folder, a path that leaves it, by `..` or by a symbolic
 * link, is refused.
 */
export function resolvePath(path: string, policy: PathPolicy): string {
  if ("anyPath" in policy) return resolve(path);
  const { folder } = policy;
  const absolute = resolve(folder, path);
  let existing = absolute;
  const rest: string[] = [];
  // lstat sees a link whose target is missing; existsSync does not, and a write would follow that link out of the folder.
  while (!exists(existing) && dirname(existing) !== existing) {
    rest.unshift(basename(existing));
    existing = dirname(existing);
  }
  if (exists(existing) && !existsSync(existing)) throw new UsageError(`"${path}" goes through a link that points nowhere`);
  const real = join(realpathSync(existing), ...rest);
  const realRoot = realpathSync(folder);
  // A root such as `/` already ends in the separator.
  const prefix = realRoot.endsWith(sep) ? realRoot : realRoot + sep;
  if (real !== realRoot && !real.startsWith(prefix)) throw new UsageError(`"${path}" is outside the allowed folder`);
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

/**
 * Write a file. It refuses to replace one unless `overwrite` is true, never replaces the input, does not open
 * a link at the path it is given (the MCP tools resolve a link inside the folder before this call, and refuse one
 * that leaves it), and never rewrites a file that has more than one name.
 */
export function writeOutput(path: string, bytes: Uint8Array, overwrite: boolean, input?: string): void {
  if (input !== undefined && existsSync(path) && realpathSync(path) === realpathSync(input)) throw new UsageError("the output is the input file");
  // A named pipe or a device as the output would block the server at open, or write to a place that is not a file.
  const existing = exists(path) ? lstatSync(path) : undefined;
  if (existing && !existing.isFile() && !existing.isSymbolicLink()) throw new UsageError(`"${path}" is not a plain file`);
  let fd: number;
  try {
    fd = openSync(path, constants.O_WRONLY | constants.O_CREAT | (constants.O_NOFOLLOW ?? 0) | (overwrite ? 0 : constants.O_EXCL), 0o666);
  } catch (error) {
    const code = (error as NodeJS.ErrnoException).code;
    if (code === "EEXIST") throw new UsageError(`"${path}" exists; pass --force to replace it`);
    if (code === "ELOOP") throw new UsageError(`"${path}" is a link; the output must be a plain file`);
    throw error;
  }
  try {
    if (fstatSync(fd).nlink > 1) throw new UsageError(`"${path}" has more than one name (a hard link); it is not replaced`);
    ftruncateSync(fd, 0);
    for (let at = 0; at < bytes.length; ) at += writeSync(fd, bytes, at, bytes.length - at);
  } finally {
    closeSync(fd);
  }
}

/** The text for an error the caller can fix, or undefined for any other error. */
export function describeError(error: unknown): string | undefined {
  // A message can quote a name from the file, which has no length limit.
  const short = (message: string) => (message.length > 400 ? `${message.slice(0, 400)}...` : message);
  if (error instanceof FormatError) return `${error.code}: ${short(error.message)}`;
  if (error instanceof UsageError) return `UsageError: ${short(error.message)}`;
  if ((error as NodeJS.ErrnoException).code === "ENOENT") return "UsageError: file not found";
  return undefined;
}
