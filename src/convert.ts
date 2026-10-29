import { readDocx, writeDocx } from "./docx.js";
import { FormatError, withLimits, type Limits } from "./limits.js";
import type { Block, Deck, Inline, List, TextDoc, Workbook, WriteResult } from "./model.js";
import { readOdp, writeOdp } from "./odp.js";
import { readOds, writeOds } from "./ods.js";
import { readOdt, writeOdt } from "./odt.js";
import { readPptx, writePptx } from "./pptx.js";
import { readXlsx, writeXlsx } from "./xlsx.js";

export type Kind = "text" | "sheet" | "deck";
export type Format = "odt" | "docx" | "ods" | "xlsx" | "odp" | "pptx";

/** A file read into its neutral model. `kind` tells which model it is. */
/** The most warnings, and the longest one, that a file reports to its caller. A hostile file can name its own warnings. */
const MAX_WARNINGS = 100;
const MAX_WARNING_CHARS = 300;

/** Keep the first warnings and cut the long ones, so a hostile file cannot flood a reply with text it chose. */
export function capWarnings(warnings: string[]): string[] {
  const shown = warnings.slice(0, MAX_WARNINGS).map((w) => (w.length > MAX_WARNING_CHARS ? `${w.slice(0, MAX_WARNING_CHARS)}...` : w));
  return warnings.length > MAX_WARNINGS ? [...shown, `${warnings.length - MAX_WARNINGS} more warnings were left out`] : shown;
}

export type Loaded = { kind: "text"; model: TextDoc; warnings: string[] } | { kind: "sheet"; model: Workbook; warnings: string[] } | { kind: "deck"; model: Deck; warnings: string[] };

export const FORMATS: readonly Format[] = ["odt", "docx", "ods", "xlsx", "odp", "pptx"];
const KIND_OF: Record<Format, Kind> = { odt: "text", docx: "text", ods: "sheet", xlsx: "sheet", odp: "deck", pptx: "deck" };

/** The format a file name asks for, by its extension. */
export function formatOfName(name: string): Format | undefined {
  const ext = /\.([a-z]+)$/i.exec(name)?.[1].toLowerCase();
  return FORMATS.find((f) => f === ext);
}

/** Read a file of a known format into its model. */
export function load(bytes: Uint8Array, format: Format, limits?: Partial<Limits>): Loaded {
  switch (format) {
    case "odt":
    case "docx": {
      const { model, warnings } = (format === "odt" ? readOdt : readDocx)(bytes, limits);
      return { kind: "text", model, warnings: capWarnings(warnings) };
    }
    case "ods":
    case "xlsx": {
      const { model, warnings } = (format === "ods" ? readOds : readXlsx)(bytes, limits);
      return { kind: "sheet", model, warnings: capWarnings(warnings) };
    }
    case "odp":
    case "pptx": {
      const { model, warnings } = (format === "odp" ? readOdp : readPptx)(bytes, limits);
      return { kind: "deck", model, warnings: capWarnings(warnings) };
    }
  }
}

function save(loaded: Loaded, format: Format, limits?: Partial<Limits>): WriteResult {
  if (loaded.kind !== KIND_OF[format]) throw new FormatError("UNSUPPORTED", `a ${loaded.kind === "sheet" ? "spreadsheet" : loaded.kind === "deck" ? "slide deck" : "text document"} cannot be written as ${format}`);
  switch (loaded.kind) {
    case "text":
      return (format === "odt" ? writeOdt : writeDocx)(loaded.model, limits);
    case "sheet":
      return (format === "ods" ? writeOds : writeXlsx)(loaded.model, limits);
    case "deck":
      return (format === "odp" ? writeOdp : writePptx)(loaded.model, limits);
  }
}

export interface Conversion extends WriteResult {
  from: Format;
  to: Format;
  /** Warnings from the reader come first, then those from the writer. */
  warnings: string[];
}

/** Convert a file from one format to another of the same kind. */
export function convert(bytes: Uint8Array, from: Format, to: Format, limits?: Partial<Limits>): Conversion {
  const caps = withLimits(limits);
  const loaded = load(bytes, from, caps);
  const written = save(loaded, to, caps);
  return { from, to, bytes: written.bytes, warnings: capWarnings([...loaded.warnings, ...written.warnings]) };
}

const countInline = (content: Inline[], into: Record<string, number>): number => {
  let characters = 0;
  for (const node of content) {
    if (node.type === "text") {
      characters += [...node.text].length;
      if (node.href) into.links++;
    } else if (node.type === "image") into.images++;
    else if (node.type === "footnote") {
      into.footnotes++;
      characters += countInline(node.content, into);
    }
  }
  return characters;
};

/** What a file holds, in numbers. It reads the whole file, so it also proves the file is valid. */
export function inspect(bytes: Uint8Array, format: Format, limits?: Partial<Limits>): Record<string, unknown> {
  const loaded = load(bytes, format, limits);
  const base = { format, kind: loaded.kind, warnings: loaded.warnings };
  // A sheet name comes from the file and has no length limit when it is read.
  const shortName = (name: string) => (name.length > MAX_WARNING_CHARS ? `${name.slice(0, MAX_WARNING_CHARS)}...` : name);
  if (loaded.kind === "sheet") {
    return { ...base, sheets: loaded.model.sheets.map((s) => ({ name: shortName(s.name), cells: Object.keys(s.cells).length, formulas: Object.values(s.cells).filter((c) => c.formula !== undefined).length })) };
  }
  if (loaded.kind === "deck") {
    const shapes = { textbox: 0, shape: 0, picture: 0 };
    for (const slide of loaded.model.slides) for (const shape of slide.shapes) shapes[shape.type]++;
    return { ...base, width: loaded.model.width, height: loaded.model.height, slides: loaded.model.slides.length, shapes, notes: loaded.model.slides.filter((s) => s.notes).length };
  }
  const counts = { links: 0, images: 0, footnotes: 0 };
  const blocks = { heading: 0, paragraph: 0, list: 0, table: 0 };
  let characters = 0;
  const walkList = (list: List): void => {
    for (const item of list.items) {
      characters += countInline(item.content, counts);
      if (item.sublist) walkList(item.sublist);
    }
  };
  const walk = (list: Block[]): void => {
    for (const block of list) {
      blocks[block.type]++;
      if (block.type === "heading" || block.type === "paragraph") characters += countInline(block.content, counts);
      else if (block.type === "list") walkList(block);
      else for (const row of block.rows) for (const cell of row) walk(cell);
    }
  };
  walk(loaded.model.blocks);
  return { ...base, ...(loaded.model.title ? { title: loaded.model.title } : {}), blocks, characters, ...counts };
}

/** The longest model, in characters of JSON, that a tool hands to an agent. */
export const MAX_MODEL_JSON = 8 * 1024 * 1024;

/**
 * The model as JSON for an agent to read. Image bytes become their size. A model past `max`
 * characters is refused while it is written, before the whole text exists.
 */
export function modelJson(model: unknown, max = MAX_MODEL_JSON): string {
  let size = 0;
  return JSON.stringify(model, (key, value: unknown) => {
    // Image bytes are replaced by their size, so count the replacement and never turn the bytes into a string.
    if (value instanceof Uint8Array) return { bytes: value.length };
    size += key.length + (typeof value === "string" ? JSON.stringify(value).length : typeof value === "object" ? 2 : String(value).length);
    if (size > max) throw new FormatError("LIMIT_EXCEEDED", `the model is larger than ${max} characters as JSON; use inspect_file for counts`);
    return value;
  });
}
