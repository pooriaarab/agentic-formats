import type { TextDoc } from "../model.js";

/** A 1x1 PNG. */
export const PNG = Uint8Array.from(atob("iVBORw0KGgoAAAANSUhEUgAAAAEAAAABCAYAAAAfFcSJAAAADUlEQVR42mP8z8BQDwAEhQGAhKmMIQAAAABJRU5ErkJggg=="), (c) => c.charCodeAt(0));

/** One document with every block and mark in the model. Each format's test writes it, reads it back, and sends it through LibreOffice. */
export const textFixture: TextDoc = {
  title: "Quarterly notes & <plans>",
  blocks: [
    { type: "heading", level: 1, content: [{ type: "text", text: "Plan" }] },
    { type: "heading", level: 3, content: [{ type: "text", text: "Details" }] },
    {
      type: "paragraph",
      content: [
        { type: "text", text: "Plain, " },
        { type: "text", text: "bold", marks: ["bold"] },
        { type: "text", text: " " },
        { type: "text", text: "italic and underline", marks: ["italic", "underline"] },
        { type: "text", text: " " },
        { type: "text", text: "struck", marks: ["strike"] },
        { type: "text", text: " " },
        { type: "text", text: "code()", marks: ["code"] },
        { type: "text", text: " x" },
        { type: "text", text: "2", marks: ["superscript"] },
        { type: "text", text: " H" },
        { type: "text", text: "2", marks: ["subscript"] },
        { type: "text", text: "O, " },
        { type: "text", text: "a link", href: "https://example.org/a?b=1&c=2" },
        { type: "text", text: " and " },
        { type: "text", text: "bold link", marks: ["bold"], href: "mailto:team@example.org" },
        { type: "text", text: "." },
      ],
    },
    { type: "paragraph", content: [{ type: "text", text: "  lead, a    gap,\ttab" }, { type: "break" }, { type: "text", text: "after break  " }] },
    { type: "paragraph", content: [{ type: "text", text: "Entities: 5 < 6 & \"quoted\" 'single' ünï ✓ 𝒳" }] },
    {
      type: "list",
      ordered: false,
      items: [
        { content: [{ type: "text", text: "one" }] },
        { content: [{ type: "text", text: "two" }], sublist: { type: "list", ordered: true, items: [{ content: [{ type: "text", text: "two-a" }] }, { content: [{ type: "text", text: "two-b" }], sublist: { type: "list", ordered: false, items: [{ content: [{ type: "text", text: "deep" }] }] } }] } },
      ],
    },
    { type: "paragraph", content: [{ type: "text", text: "Between lists." }] },
    { type: "list", ordered: true, items: [{ content: [{ type: "text", text: "first" }] }, { content: [{ type: "text", text: "second", marks: ["bold"] }] }] },
    {
      type: "table",
      header: true,
      rows: [
        [[{ type: "paragraph", content: [{ type: "text", text: "Name" }] }], [{ type: "paragraph", content: [{ type: "text", text: "Qty" }] }]],
        [[{ type: "paragraph", content: [{ type: "text", text: "Apples" }] }], []],
        [[{ type: "paragraph", content: [{ type: "text", text: "Two" }] }, { type: "paragraph", content: [{ type: "text", text: "paragraphs" }] }], [{ type: "paragraph", content: [{ type: "text", text: "7", marks: ["bold"] }] }]],
      ],
    },
    { type: "paragraph", content: [{ type: "text", text: "Image: " }, { type: "image", data: PNG, mime: "image/png", alt: "a red dot", width: 72, height: 36 }, { type: "text", text: " ends." }] },
    {
      type: "paragraph",
      content: [{ type: "text", text: "Claim" }, { type: "footnote", content: [{ type: "text", text: "See " }, { type: "text", text: "source", marks: ["italic"] }, { type: "break" }, { type: "text", text: "page 4." }] }, { type: "text", text: " holds." }],
    },
  ],
};
