/*
 * The neutral models. Each format reader fills one of them and each writer takes one, so an app
 * converts once and never touches zip or XML. Lengths are in points (1/72 inch).
 */
import { FormatError } from "./limits.js";

/** What a reader returns: the model, plus every part of the file that had no place in it. */
export interface ReadResult<T> {
  model: T;
  warnings: string[];
}

/** What a writer returns: the file, plus every part of the model that the format could not hold. */
export interface WriteResult {
  bytes: Uint8Array;
  warnings: string[];
}

// ---------- text documents ----------

export type Mark = "bold" | "italic" | "underline" | "strike" | "code" | "superscript" | "subscript";
/** The order in which a model lists the marks of a run. */
/** Refuse a run whose text or link is not a string: a size taken from the wrong type would pass every count. */
export function checkRun(run: TextRun): void {
  if (typeof run?.text !== "string" || (run.href !== undefined && typeof run.href !== "string")) throw new FormatError("INVALID_MODEL", "a text run needs text and a link that are strings");
}

export const MARKS: readonly Mark[] = ["bold", "italic", "underline", "strike", "code", "superscript", "subscript"];

export type ImageMime = "image/png" | "image/jpeg" | "image/gif";
export interface ImageData {
  data: Uint8Array;
  mime: ImageMime;
  alt?: string;
}

export interface TextRun {
  type: "text";
  text: string;
  marks?: Mark[];
  /** A web, mail, in-page, or relative link target. */
  href?: string;
}
export interface LineBreak {
  type: "break";
}
export interface InlineImage extends ImageData {
  type: "image";
  width?: number;
  height?: number;
}
export interface Footnote {
  type: "footnote";
  content: (TextRun | LineBreak)[];
}
export type Inline = TextRun | LineBreak | InlineImage | Footnote;

export interface Heading {
  type: "heading";
  level: 1 | 2 | 3 | 4 | 5 | 6;
  content: Inline[];
}
export interface Paragraph {
  type: "paragraph";
  content: Inline[];
}
export interface ListItem {
  content: Inline[];
  sublist?: List;
}
export interface List {
  type: "list";
  ordered: boolean;
  items: ListItem[];
}
export interface Table {
  type: "table";
  /** The first row is a header row. */
  header?: boolean;
  /** Rows of cells; each cell holds blocks. */
  rows: Block[][][];
}
export type Block = Heading | Paragraph | List | Table;

export interface TextDoc {
  title?: string;
  blocks: Block[];
}

/** Two runs read the same when their marks and their link are equal. Compares in place: a link may be megabytes long. */
const sameStyle = (a: TextRun, b: TextRun): boolean => a.href === b.href && (a.marks?.join() ?? "") === (b.marks?.join() ?? "");

/** The marks of `a` and `b` together, each once, in the canonical order. Seven entries at most, however many styles a file stacks. */
export const mergeMarks = (a: readonly Mark[], b: readonly Mark[]): Mark[] => MARKS.filter((m) => a.includes(m) || b.includes(m));

const inlineLength = (content: Inline[]): number =>
  content.reduce((n, node) => n + (node.type === "text" ? node.text.length : node.type === "footnote" ? inlineLength(node.content) : node.type === "image" ? node.data.length : 1), 0);

/** Characters of text in blocks, and bytes of images, counted as they are laid out: a row that appears twice counts twice. */
export function textLength(blocks: Block[]): number {
  return blocks.reduce((n, block) => {
    if (block.type === "table") return n + block.rows.reduce((m, row) => m + row.reduce((k, cell) => k + textLength(cell), 0), 0);
    if (block.type === "list") {
      const items = (list: List): number => list.items.reduce((m, item) => m + inlineLength(item.content) + (item.sublist ? items(item.sublist) : 0), 0);
      return n + items(block);
    }
    return n + inlineLength(block.content);
  }, 0);
}

/** Merge runs that look the same, drop empty ones, and put marks in the canonical order. */
export function normalizeInline(content: Inline[]): Inline[] {
  const out: Inline[] = [];
  for (const node of content) {
    if (node.type !== "text") {
      out.push(node.type === "footnote" ? { ...node, content: normalizeInline(node.content) as Footnote["content"] } : node);
      continue;
    }
    if (!node.text) continue;
    const marks = MARKS.filter((m) => node.marks?.includes(m));
    const run: TextRun = { type: "text", text: node.text, ...(marks.length ? { marks } : {}), ...(node.href ? { href: node.href } : {}) };
    const prev = out[out.length - 1];
    if (prev?.type === "text" && sameStyle(prev, run)) prev.text += run.text;
    else out.push(run);
  }
  return out;
}

function normalizeList(list: List): List {
  return {
    type: "list",
    ordered: list.ordered,
    items: list.items.map((item) => ({ content: normalizeInline(item.content), ...(item.sublist ? { sublist: normalizeList(item.sublist) } : {}) })),
  };
}

/** A cell with one empty paragraph is an empty cell. */
function normalizeCell(cell: Block[]): Block[] {
  const blocks = cell.map(normalizeBlock);
  return blocks.length === 1 && blocks[0].type === "paragraph" && blocks[0].content.length === 0 ? [] : blocks;
}

export function normalizeBlock(block: Block): Block {
  switch (block.type) {
    case "heading":
    case "paragraph":
      return { ...block, content: normalizeInline(block.content) };
    case "list":
      return normalizeList(block);
    case "table":
      return { type: "table", ...(block.header ? { header: true } : {}), rows: block.rows.map((row) => row.map(normalizeCell)) };
  }
}

/** The form a reader returns. Compare `normalizeTextDoc(a)` with a reader's result to test a round trip. */
export function normalizeTextDoc(doc: TextDoc): TextDoc {
  return { ...(doc.title ? { title: doc.title } : {}), blocks: doc.blocks.map(normalizeBlock) };
}

// ---------- workbooks ----------

export type CellValue = string | number | boolean;
export interface Cell {
  value?: CellValue;
  /** The formula in A1 notation, without the leading `=`. Example: `SUM(A1:A3)`. */
  formula?: string;
  /** A number format code. The codes in `NUMBER_FORMATS` work in every format. */
  format?: string;
}
export interface Sheet {
  name: string;
  /** Cells by A1 reference. Empty cells are absent. */
  cells: Record<string, Cell>;
}
export interface Workbook {
  sheets: Sheet[];
}

/** Number formats that ODS and XLSX both hold. A date cell is a serial number (days since 1899-12-30) with `yyyy-mm-dd`. */
export const NUMBER_FORMATS = ["0", "0.00", "#,##0", "#,##0.00", "0%", "0.00%", "yyyy-mm-dd"] as const;

const MAX_COLUMN = 16384;
const MAX_ROW = 1048576;

export function columnName(column: number): string {
  let name = "";
  for (let n = column; n > 0; n = Math.floor((n - 1) / 26)) name = String.fromCharCode(65 + ((n - 1) % 26)) + name;
  return name;
}

/** Split `B12` into a 1-based column and row. Anything off the sheet grid returns undefined. */
export function parseRef(ref: string): { column: number; row: number } | undefined {
  const found = /^([A-Z]{1,3})([1-9][0-9]{0,6})$/.exec(ref);
  if (!found) return undefined;
  const column = [...found[1]].reduce((n, c) => n * 26 + c.charCodeAt(0) - 64, 0);
  const row = Number(found[2]);
  return column <= MAX_COLUMN && row <= MAX_ROW ? { column, row } : undefined;
}

export const refOf = (column: number, row: number): string => `${columnName(column)}${row}`;

/** The cell references of a sheet in reading order: row by row, left to right. */
export function sortedRefs(sheet: Sheet): string[] {
  const keyed = Object.keys(sheet.cells).map((ref) => ({ ref, at: parseRef(ref) }));
  return keyed
    .filter((k): k is { ref: string; at: { column: number; row: number } } => !!k.at)
    .sort((a, b) => a.at.row - b.at.row || a.at.column - b.at.column)
    .map((k) => k.ref);
}

// ---------- slide decks ----------

export interface Frame {
  x: number;
  y: number;
  width: number;
  height: number;
}
export interface TextBox {
  type: "textbox";
  frame: Frame;
  /** Paragraphs of text runs. */
  paragraphs: TextRun[][];
}
export interface ShapeBox {
  type: "shape";
  kind: "rectangle" | "ellipse" | "line";
  frame: Frame;
  /** `#rrggbb` */
  fill?: string;
  /** `#rrggbb` */
  stroke?: string;
}
export interface PictureBox extends ImageData {
  type: "picture";
  frame: Frame;
}
export type SlideShape = TextBox | ShapeBox | PictureBox;
export interface Slide {
  shapes: SlideShape[];
  /** Speaker notes as plain text, one paragraph per line. */
  notes?: string;
}
export interface Deck {
  width: number;
  height: number;
  slides: Slide[];
}
