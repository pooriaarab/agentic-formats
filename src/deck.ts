import { FormatError, capped, type Limits } from "./limits.js";
import { checkRun, type Deck, type Frame, type TextRun } from "./model.js";

export const HEX_COLOR = /^#[0-9a-fA-F]{6}$/;
/** The size of a deck, a slide, or a shape never goes past 22 inches times 10. */
export const MAX_POINTS = 15_840;
const bad = (message: string): never => {
  throw new FormatError("INVALID_MODEL", message);
};

const checkFrame = (frame: Frame, where: string): void => {
  if (typeof frame !== "object" || frame === null) bad(`${where} needs a frame`);
  for (const n of [frame.x, frame.y, frame.width, frame.height]) if (!Number.isFinite(n) || Math.abs(n) > MAX_POINTS) bad(`${where} has a position or size that is not usable`);
  if (frame.width < 0 || frame.height < 0) bad(`${where} has a negative size`);
};

/** Refuse a deck that no writer could write: sizes that are not usable, colors that are not `#rrggbb`. */
/**
 * The most characters the XML of a text can need, found without making it: the text, a tag of `perLine`
 * characters for each line, and five more for each character that XML escapes.
 */
export function textBound(text: string, perLine: number): number {
  let n = text.length + perLine;
  for (let i = 0; i < text.length; i++) {
    const c = text[i];
    if (c === "\n") n += perLine;
    else if (c === "&" || c === "<" || c === ">" || c === '"' || c === "\r") n += 5;
  }
  return n;
}

const SHAPE_TYPES = ["textbox", "picture", "shape"];

export function checkDeck(deck: Deck, limits: Limits): void {
  if (typeof deck !== "object" || deck === null || !Array.isArray(deck.slides)) bad("a deck needs a list of slides");
  checkFrame({ x: 0, y: 0, width: deck.width, height: deck.height }, "the deck");
  if (deck.width <= 0 || deck.height <= 0) bad("the deck needs a size above zero");
  if (deck.slides.length === 0) bad("a deck needs at least one slide");
  if (deck.slides.length > limits.maxEntries / 4) bad(`a deck holds at most ${Math.floor(limits.maxEntries / 4)} slides`);
  // A run or a note that the model holds in many places is written in each place, so each use counts.
  const chargeText = capped(limits.maxTextChars, `the deck has more than ${limits.maxTextChars} characters of text`);
  deck.slides.forEach((slide, i) => {
    if (typeof slide !== "object" || slide === null) bad(`slide ${i + 1} is not an object`);
    if (!Array.isArray(slide.shapes)) bad(`slide ${i + 1} needs a list of shapes`);
    if (slide.notes !== undefined && typeof slide.notes !== "string") bad(`slide ${i + 1} has notes that are not a string`);
    chargeText(slide.notes?.length ?? 0);
    slide.shapes.forEach((shape, j) => {
      if (typeof shape !== "object" || shape === null) bad(`slide ${i + 1}, shape ${j + 1} is not an object`);
      const where = `slide ${i + 1}, shape ${j + 1}`;
      if (!SHAPE_TYPES.includes(shape.type)) bad(`${where} is not a kind of shape`);
      checkFrame(shape.frame, where);
      // A link target or a description that many runs or pictures share is written for each of them.
      if (shape.type === "textbox") {
        if (!Array.isArray(shape.paragraphs) || shape.paragraphs.some((runs) => !Array.isArray(runs))) bad(`${where} needs paragraphs that are lists of runs`);
        for (const runs of shape.paragraphs) {
          for (const run of runs) {
            checkRun(run);
            chargeText(run.text.length + (run.href?.length ?? 0));
          }
        }
      }
      if (shape.type === "picture") {
        if (shape.alt !== undefined && typeof shape.alt !== "string") bad(`${where} has a description that is not a string`);
        chargeText(shape.alt?.length ?? 0);
      }
      if (shape.type === "shape") for (const color of [shape.fill, shape.stroke]) if (color !== undefined && !HEX_COLOR.test(color)) bad(`${where} has a color that is not #rrggbb`);
    });
  });
}

/** Plain text of a run list, for a fallback and for tests. */
export const textOfRuns = (runs: TextRun[]): string => runs.map((r) => r.text).join("");

/** Round a length to 0.1 point. */
export const tenth = (n: number): number => Math.round(n * 10) / 10;
