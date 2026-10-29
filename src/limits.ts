/** Caps that every reader applies. A file that goes past one is rejected, never trimmed. */
export interface Limits {
  /** Size of the file as given. */
  maxInputBytes: number;
  /** Number of entries in the zip. */
  maxEntries: number;
  /** Size of one entry after it is unpacked. */
  maxEntryBytes: number;
  /** Size of all entries after they are unpacked. */
  maxTotalBytes: number;
  /** Unpacked size divided by packed size, for entries above `ratioFloorBytes`. */
  maxRatio: number;
  /** Entries below this size skip the ratio check. Tiny parts compress very well by nature. */
  ratioFloorBytes: number;
  /** Size of one XML part. */
  maxXmlBytes: number;
  /** Nesting depth of XML elements. */
  maxXmlDepth: number;
  /** Number of elements in one XML part. */
  maxXmlElements: number;
  /** Number of cells in a workbook, or table cells in a text document. Repeat counts count in full. */
  maxCells: number;
  /** Characters of cell or run text, with every copy counted. Copies of one string share memory but not the output. */
  maxTextChars: number;
  /** Length of a style's parent chain. */
  maxStyleDepth: number;
}

export const DEFAULT_LIMITS: Readonly<Limits> = Object.freeze({
  maxInputBytes: 25 * 1024 * 1024,
  maxEntries: 5000,
  maxEntryBytes: 50 * 1024 * 1024,
  maxTotalBytes: 100 * 1024 * 1024,
  maxRatio: 250,
  ratioFloorBytes: 1024 * 1024,
  maxXmlBytes: 32 * 1024 * 1024,
  maxXmlDepth: 100,
  maxXmlElements: 1_000_000,
  maxCells: 1_000_000,
  maxTextChars: 8 * 1024 * 1024,
  maxStyleDepth: 50,
});

export const withLimits = (limits?: Partial<Limits>): Limits => ({ ...DEFAULT_LIMITS, ...limits });

export type FormatErrorCode =
  /** The bytes are not a valid file of the format. */
  | "INVALID_FILE"
  /** The file goes past a cap in `Limits`. */
  | "LIMIT_EXCEEDED"
  /** The model that the caller passed to a writer cannot be written. */
  | "INVALID_MODEL"
  /** The file type or the conversion has no support. */
  | "UNSUPPORTED";

export class FormatError extends Error {
  constructor(
    readonly code: FormatErrorCode,
    message: string,
  ) {
    super(message);
    this.name = "FormatError";
  }
}

/** A running total that throws `LIMIT_EXCEEDED` with `message` once it passes `limit`. */
export function capped(limit: number, message: string): (amount: number) => void {
  let used = 0;
  return (amount) => {
    // A size from a value of the wrong type (a number that is NaN or negative) must not switch the cap off.
    if (!Number.isFinite(amount) || amount < 0) throw new FormatError("INVALID_MODEL", "a size that is not a number of zero or more");
    used += amount;
    if (used > limit) throw new FormatError("LIMIT_EXCEEDED", message);
  };
}
