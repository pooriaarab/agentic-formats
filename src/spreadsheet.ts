import { FormatError, capped, type Limits } from "./limits.js";
import { parseRef, type Cell, type Workbook } from "./model.js";

/** The formats an ODS file holds as data styles. XLSX stores any code as it is. */
export const ODS_FORMAT = /^(#,##0|0)(\.0{1,30})?%?$|^yyyy-mm-dd$/;
export const isDateFormat = (format: string | undefined): boolean => format === "yyyy-mm-dd";

const DAY_MS = 86_400_000;
const EPOCH = Date.UTC(1899, 11, 30);

/** Days since 1899-12-30 as an ISO date, with a time when the serial has a fraction. Valid from serial 61. */
export function serialToIso(serial: number): string | undefined {
  const ms = EPOCH + Math.round(serial * DAY_MS);
  if (!Number.isFinite(ms) || Math.abs(ms - EPOCH) > 3_000_000 * DAY_MS) return undefined;
  const iso = new Date(ms).toISOString();
  return serial % 1 === 0 ? iso.slice(0, 10) : iso.slice(0, 19);
}

export function isoToSerial(iso: string): number | undefined {
  if (!/^\d{4}-\d{2}-\d{2}(T\d{2}:\d{2}:\d{2}(\.\d+)?)?$/.test(iso)) return undefined;
  const ms = Date.parse(`${iso}${iso.length > 10 ? "" : "T00:00:00"}Z`);
  return Number.isNaN(ms) ? undefined : (ms - EPOCH) / DAY_MS;
}

/** A formula longer than this is refused when it is read and when it is written. Formula conversion is not linear in the worst case. */
export const MAX_FORMULA = 8192;

const BAD_SHEET_CHARS = /[[\]:*?/\\\u0000-\u001f]/;

/** Refuse a workbook that no writer could write: bad sheet names, cells off the grid, numbers that are not finite. */
export function checkWorkbook(book: Workbook, limits: Limits): void {
  const bad = (message: string): never => {
    throw new FormatError("INVALID_MODEL", message);
  };
  if (!Array.isArray(book?.sheets)) bad("a workbook needs a list of sheets");
  if (book.sheets.length === 0) bad("a workbook needs at least one sheet");
  const names = new Set<string>();
  let cells = 0;
  const chargeText = capped(limits.maxTextChars, `the workbook has more than ${limits.maxTextChars} characters of cell text`);
  for (const sheet of book.sheets) {
    if (typeof sheet !== "object" || sheet === null || typeof sheet.name !== "string") bad("a sheet needs a name that is a string");
    if (typeof sheet.cells !== "object" || sheet.cells === null) bad(`the sheet "${sheet.name}" needs cells`);
    const { name } = sheet;
    if (!name || name.length > 31 || BAD_SHEET_CHARS.test(name) || name.startsWith("'") || name.endsWith("'")) bad(`"${name}" is not a usable sheet name`);
    if (names.has(name.toLowerCase())) bad(`the sheet name "${name}" is used twice`);
    names.add(name.toLowerCase());
    for (const [ref, cell] of Object.entries(sheet.cells)) {
      if (!parseRef(ref)) bad(`"${ref}" is not a cell reference on the sheet grid`);
      if (typeof cell !== "object" || cell === null) bad(`${sheet.name}!${ref} is not a cell`);
      if (cell.value !== undefined && !["string", "number", "boolean"].includes(typeof cell.value)) bad(`${sheet.name}!${ref} holds a value that is not a string, a number, or a boolean`);
      if (cell.format !== undefined && typeof cell.format !== "string") bad(`${sheet.name}!${ref} holds a number format that is not a string`);
      if (++cells > limits.maxCells) throw new FormatError("LIMIT_EXCEEDED", `the workbook has more than ${limits.maxCells} cells`);
      if (typeof cell.value === "number" && !Number.isFinite(cell.value)) bad(`${sheet.name}!${ref} holds a number that is not finite`);
      if (cell.formula !== undefined && (typeof cell.formula !== "string" || cell.formula.length > MAX_FORMULA)) bad(`${sheet.name}!${ref} holds a formula that is not usable`);
      // A string that many cells point to is written once for each cell, so it counts once for each cell.
      chargeText((typeof cell.value === "string" ? cell.value.length : 0) + (cell.formula?.length ?? 0));
    }
  }
}

/** A cell with something to write. */
export const hasContent = (cell: Cell): boolean => cell.value !== undefined || cell.formula !== undefined || cell.format !== undefined;

/** LibreOffice saves a boolean cell as the formula `TRUE()` or `FALSE()`. That is a value, not a formula. */
export const isBooleanLiteral = (formula: string, value: unknown): boolean => typeof value === "boolean" && formula.toUpperCase() === (value ? "TRUE()" : "FALSE()");
