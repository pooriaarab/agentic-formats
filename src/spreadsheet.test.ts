/*
 * ODS and XLSX, end to end: the real writers, the real readers, an independent XML parser, and
 * LibreOffice. The ways it can fail are written first; each has a case below.
 *
 *  1 a value is lost or changed: markup characters, edge spaces, newlines, text that looks like a
 *    number or a formula, negatives, booleans, a float with full precision
 *  2 a formula is mangled: separators, absolute refs, a sheet name with a space or a quote, a
 *    function name with digits, a string literal that holds refs, commas, or semicolons
 *  3 a number format is lost or wrong, or a date moves
 *  4 a cell lands in the wrong place: far cells, gaps, row order
 *  5 sheets: order, odd names, an empty sheet
 *  6 the package is malformed: a part does not parse, `mimetype` is not first and stored, a
 *    part is missing from the content types, a relationship points nowhere
 *  7 the same model gives different bytes
 *  8 a writer takes a model it cannot write: no sheets, bad names, bad refs, numbers that are not
 *    finite, too many cells; a format ODS cannot hold goes unreported
 *  9 a hostile file is read: a DOCTYPE, the wrong type, a huge repeat count, a cell off the grid,
 *    a shared string index past the table, a missing relationship, too many cells
 * 11 one big string fills many cells: a shared string used by many cells, a repeated cell, a model
 *    whose cells share one string; a writer builds its whole output before it checks the size
 * 12 a long formula hangs the formula conversion; a few bytes of <text:s> expand to millions of
 *    spaces; two sheet entries name one worksheet part; many distinct formats make a writer slow or
 *    its warnings long
 * 10 LibreOffice cannot open our file, or its re-save does not read back to the same model, in the
 *    same format or after a trip through the other format; LibreOffice does not compute a formula
 *    we wrote without a cached result
 */
import { isDeepStrictEqual } from "node:util";
import { DOMParser } from "@xmldom/xmldom";
import { strFromU8, unzipSync } from "fflate";
import { afterAll, describe, expect, test } from "vitest";
import { fromOpenFormula, toOpenFormula } from "./formula.js";
import { FormatError, withLimits } from "./limits.js";
import type { Workbook } from "./model.js";
import { readOds, writeOds } from "./ods.js";
import { isoToSerial, serialToIso } from "./spreadsheet.js";
import { LibreOffice, hasLibreOffice, libreOfficeRequired, sha256, writeArtifact, type RoundTripRecord } from "./testing/libreoffice.js";
import { readXlsx, writeXlsx } from "./xlsx.js";
import { writeZip } from "./zip.js";

const fixture: Workbook = {
  sheets: [
    {
      name: "Sales Q1",
      cells: {
        A1: { value: "Item" },
        B1: { value: "Qty" },
        C1: { value: "Price" },
        D1: { value: "Total" },
        A2: { value: "Apples & <pears> \"q\" 'a'" },
        B2: { value: 3 },
        C2: { value: 1.25, format: "0.00" },
        D2: { value: 3.75, formula: "B2*C2", format: "#,##0.00" },
        E2: { value: 45000, format: "yyyy-mm-dd" },
        F2: { value: "hello-x", formula: "'Other Sheet'!A1&\"-x\"" },
        G2: { value: 0.48, formula: "ROUND(LOG10($B$2),2)" },
        A3: { value: "  padded  text  " },
        B3: { value: -7 },
        C3: { value: 2.5 },
        D3: { value: -17.5, formula: "B3*C3" },
        E3: { value: 0.125, format: "0.00%" },
        F3: { value: 5, formula: "'O''Brien'!A1+1" },
        A4: { value: "multi\nline ünï ✓" },
        B4: { value: true },
        C4: { value: 0.25, format: "0%" },
        D4: { value: "yes", formula: 'IF(B4,"yes","no")' },
        E4: { value: "A1,B2;Apples & <pears> \"q\" 'a'", formula: 'CONCATENATE("A1,B2;",A2)' },
        A5: { value: "Total" },
        D5: { value: -13.75, formula: "SUM(D2:D3)" },
        A6: { value: "007" },
        B6: { value: "=SUM(1,2)" },
        C6: { value: "TRUE" },
        H1: { value: 42, format: "0" },
        H2: { value: 1234567, format: "#,##0" },
        H3: { value: 0.1234, format: "0.00" },
        Z100: { value: "far" },
      },
    },
    { name: "Other Sheet", cells: { A1: { value: "hello" }, B2: { value: 2 } } },
    { name: "O'Brien", cells: { A1: { value: 4, formula: "'Other Sheet'!B2*2" } } },
    { name: "Empty", cells: {} },
  ],
};

const writers = { ods: writeOds, xlsx: writeXlsx } as const;
const readers = { ods: readOds, xlsx: readXlsx } as const;
const formats = ["ods", "xlsx"] as const;

const wellFormed = (xml: string): void => {
  const seen: string[] = [];
  new DOMParser({ onError: (level, message) => seen.push(`${level}: ${message}`) }).parseFromString(xml, "text/xml");
  expect(seen).toEqual([]);
};
const patch = (bytes: Uint8Array, name: string, change: (xml: string) => string): Uint8Array => {
  const files = unzipSync(bytes);
  const entries = Object.entries(files).map(([n, data]) => ({ name: n, data: n === name ? change(strFromU8(data)) : data, store: n === "mimetype" }));
  return writeZip(entries, withLimits());
};
const sheet = (name: string, cells: Workbook["sheets"][0]["cells"]): Workbook => ({ sheets: [{ name, cells }] });

describe.each(formats)("%s: write then read", (format) => {
  const write = writers[format];
  const read = readers[format];

  test("1-5 the workbook reads back unchanged, with no warnings", () => {
    const written = write(fixture);
    const back = read(written.bytes);
    expect(back.model).toEqual(fixture);
    expect(back.warnings).toEqual([]);
    expect(written.warnings).toEqual([]);
  });

  test("1 a float with full precision, edge spaces, and an empty string survive", () => {
    const book = sheet("S", { A1: { value: 0.1 + 0.2 }, A2: { value: " x " }, A3: { value: "" }, A4: { value: 1e-7 }, A5: { value: 123456789012345680000 }, A6: { value: "\ttab\t" } });
    expect(read(write(book).bytes).model).toEqual(book);
  });

  test("4 cells far apart keep their places, whatever the order they were added", () => {
    const book = sheet("S", { XFD1048576: { value: "last" }, C3: { value: 3 }, A1: { value: 1 }, AA2: { value: 2 }, B1048575: { value: 4 } });
    expect(read(write(book).bytes).model).toEqual(book);
  });

  test("7 the same model gives the same bytes", () => {
    expect(write(fixture).bytes).toEqual(write(fixture).bytes);
  });

  test("8 a writer refuses a model it cannot write", () => {
    const refuse = (book: Workbook, limits = {}) => expect(() => write(book, limits)).toThrow(FormatError);
    refuse({ sheets: [] });
    refuse({ sheets: [{ name: "A", cells: {} }, { name: "a", cells: {} }] });
    for (const name of ["", "x".repeat(32), "a/b", "a:b", "[a]", "a?", "a*", "a\\b", "'a", "a'", "a\nb", "a\tb"]) refuse({ sheets: [{ name, cells: {} }] });
    for (const ref of ["A0", "a1", "AAAA1", "1A", "XFE1", "A1048577", "A01"]) refuse(sheet("S", { [ref]: { value: 1 } }));
    for (const value of [Number.NaN, Infinity, -Infinity]) refuse(sheet("S", { A1: { value } }));
    refuse(sheet("S", { A1: { value: 1 }, A2: { value: 2 } }), { maxCells: 1 });
    for (const cell of [{ value: 1, format: 123 }, { value: {} }, { value: [1] }, null, 7]) refuse(sheet("S", { A1: cell as never }));
    expect(() => write(sheet("S", { A1: { value: 1, format: "constructor" } }))).not.toThrow();
    for (const book of [{ sheets: "x" }, {}, { sheets: [null] }, { sheets: [{ name: ["a"], cells: {} }] }, { sheets: [{ name: 5, cells: {} }] }, { sheets: [{ name: "S", cells: null }] }, { sheets: [{ name: "S" }] }]) refuse(book as never);
  });

  test("11 a model whose cells share one big string is refused before any output is built", () => {
    const big = "x".repeat(50_000);
    const cells = { A1: { value: big }, A2: { value: big }, A3: { value: big } };
    const caps = { maxTextChars: 100_000 };
    expect(() => write(sheet("S", cells), caps)).toThrow(/more than 100000 characters of cell text/);
    expect(() => write(sheet("S", { A1: cells.A1, A2: cells.A2 }), caps)).not.toThrow();
    expect(() => write(sheet("S", { A1: { value: 1, formula: "A".repeat(8000) }, A2: { value: 1, formula: "A".repeat(8000) } }), { maxTextChars: 15_000 })).toThrow(/characters of cell text/);
  });

  test("11 the output size is capped while it is written, not after", () => {
    const cells = Object.fromEntries(Array.from({ length: 100 }, (_, i) => [`A${i + 1}`, { value: "abcdefghij" }]));
    expect(() => write(sheet("S", cells), { maxEntryBytes: 1500, maxTotalBytes: 1500 })).toThrow(/while it is written/);
  });

  test("6 every XML part parses with an independent parser", () => {
    const files = unzipSync(write(fixture).bytes);
    const xml = Object.keys(files).filter((n) => /\.(xml|rels)$/.test(n));
    expect(xml.length).toBeGreaterThanOrEqual(4);
    for (const name of xml) wellFormed(strFromU8(files[name]));
  });
});

describe("ods package", () => {
  test("6 `mimetype` is first and stored, and the manifest carries `manifest:version`", () => {
    const { bytes } = writeOds(fixture);
    expect(strFromU8(bytes.subarray(30, 38))).toBe("mimetype");
    expect(new DataView(bytes.buffer, bytes.byteOffset).getUint16(8, true)).toBe(0);
    expect(strFromU8(unzipSync(bytes)["META-INF/manifest.xml"])).toMatch(/<manifest:manifest [^>]*manifest:version="1\.3"/);
  });

  test("8 a number format with more than 30 decimals is reported and the value is kept", () => {
    const at = (zeros: number) => writeOds(sheet("S", { A1: { value: 1, format: `0.${"0".repeat(zeros)}` } }));
    expect(at(30).warnings).toEqual([]);
    expect(at(31).warnings).toEqual([`the number format "0.${"0".repeat(31)}" is not written in ODS`]);
  });

  test("8 a format that ODS cannot hold is reported and the value is kept", () => {
    const book = sheet("S", { A1: { value: 1, format: "[$€-407] #,##0.00" } });
    const { bytes, warnings } = writeOds(book);
    expect(warnings).toEqual(['the number format "[$€-407] #,##0.00" is not written in ODS']);
    expect(readOds(bytes).model).toEqual(sheet("S", { A1: { value: 1 } }));
  });

  test("3 dates are real dates in the file", () => {
    const content = strFromU8(unzipSync(writeOds(fixture).bytes)["content.xml"]);
    expect(content).toContain('office:value-type="date" office:date-value="2023-03-15"');
    expect(content).toContain('table:formula="of:=[$\'Other Sheet\'.A1]&amp;&quot;-x&quot;"');
    expect(content).toContain("CONCATENATE(&quot;A1,B2;&quot;;[.A2])");
  });
});

describe("xlsx package", () => {
  test("6 every part is in the content types, and every relationship target exists", () => {
    const files = unzipSync(writeXlsx(fixture).bytes);
    const types = strFromU8(files["[Content_Types].xml"]);
    for (const part of Object.keys(files).filter((n) => n.endsWith(".xml") && n !== "[Content_Types].xml")) expect(types).toContain(`PartName="/${part}"`);
    expect(Object.keys(files)[0]).toBe("[Content_Types].xml");
    const rels = strFromU8(files["xl/_rels/workbook.xml.rels"]);
    for (const target of [...rels.matchAll(/Target="([^"]+)"/g)].map((m) => m[1])) expect(files[`xl/${target}`]).toBeDefined();
  });

  test("3 numbers use built-in format ids where they exist, and custom ones for the rest", () => {
    const styles = strFromU8(unzipSync(writeXlsx(fixture).bytes)["xl/styles.xml"]);
    expect(styles).toContain('<numFmt numFmtId="164" formatCode="yyyy-mm-dd"/>');
    for (const id of [1, 2, 3, 4, 9, 10]) expect(styles).toContain(`<xf numFmtId="${id}"`);
  });

  test("3 a format code that names a key of Object.prototype writes a valid styles part", () => {
    const styles = strFromU8(unzipSync(writeXlsx(sheet("S", { A1: { value: 1, format: "constructor" } })).bytes)["xl/styles.xml"]);
    expect(styles).toContain('formatCode="constructor"');
    expect(styles).not.toContain("native code");
    wellFormed(styles);
  });

  test("3 any format code is kept as written", () => {
    const book = sheet("S", { A1: { value: 5, format: '#,##0.0" kg"' } });
    expect(readXlsx(writeXlsx(book).bytes).model).toEqual(book);
  });
});

describe("formulas", () => {
  const pairs: [string, string][] = [
    ["SUM(A1:B2)", "SUM([.A1:.B2])"],
    ["SUM(A1,B2,3)", "SUM([.A1];[.B2];3)"],
    ["$A$1+B$2+$C3", "[.$A$1]+[.B$2]+[.$C3]"],
    ["Sheet2!A1:B2", "[$Sheet2.A1:.B2]"],
    ["'My Sheet'!A1*2", "[$'My Sheet'.A1]*2"],
    ["LOG10(A1)+ATAN2(1,2)", "LOG10([.A1])+ATAN2(1;2)"],
    ['CONCATENATE("A1,B2;","x""y")', 'CONCATENATE("A1,B2;";"x""y")'],
    ["IF(A1>0,\"[.A1]\",\"\")", "IF([.A1]>0;\"[.A1]\";\"\")"],
  ];
  test.each(pairs)("2 %s", (a1, open) => {
    expect(toOpenFormula(a1)).toBe(open);
    expect(fromOpenFormula(open)).toBe(a1);
  });

  test("2 an OpenFormula reference outside the A1 forms is refused, not guessed", () => {
    expect(fromOpenFormula("SUM([.A:.A])")).toBeUndefined();
    expect(fromOpenFormula("[$'file.ods'#$Sheet1.A1]")).toBeUndefined();
  });

  test("3 serial numbers and ISO dates convert both ways", () => {
    expect(serialToIso(45000)).toBe("2023-03-15");
    expect(serialToIso(45000.5)).toBe("2023-03-15T12:00:00");
    expect(isoToSerial("2023-03-15")).toBe(45000);
    expect(isoToSerial("2023-03-15T12:00:00")).toBe(45000.5);
    expect(isoToSerial("March 15")).toBeUndefined();
  });
});

describe("hostile files", () => {
  const ods = (body: string, mime = "application/vnd.oasis.opendocument.spreadsheet") =>
    writeZip(
      [
        { name: "mimetype", data: mime, store: true },
        {
          name: "content.xml",
          data: `<?xml version="1.0"?><office:document-content xmlns:office="urn:oasis:names:tc:opendocument:xmlns:office:1.0" xmlns:table="urn:oasis:names:tc:opendocument:xmlns:table:1.0" xmlns:text="urn:oasis:names:tc:opendocument:xmlns:text:1.0"><office:body><office:spreadsheet>${body}</office:spreadsheet></office:body></office:document-content>`,
        },
      ],
      withLimits(),
    );
  const table = (rows: string) => `<table:table table:name="S">${rows}</table:table>`;
  const one = writeXlsx(sheet("S", { A1: { value: "x" } })).bytes;

  test("9 a DOCTYPE with an entity is refused", () => {
    const withDoctype = patch(writeOds(sheet("S", { A1: { value: "x" } })).bytes, "content.xml", (x) => x.replace("?>", `?><!DOCTYPE x [<!ENTITY e SYSTEM "file:///etc/passwd">]>`));
    expect(() => readOds(withDoctype)).toThrow(/DOCTYPE/);
    expect(() => readXlsx(patch(one, "xl/worksheets/sheet1.xml", (x) => x.replace("?>", `?><!DOCTYPE x [<!ENTITY e "x">]>`)))).toThrow(/DOCTYPE/);
  });

  test("9 the wrong type of file is refused", () => {
    expect(() => readOds(ods(table(""), "application/vnd.oasis.opendocument.text"))).toThrow(/not an ODS/);
    expect(() => readOds(one)).toThrow(FormatError);
    expect(() => readXlsx(writeOds(fixture).bytes)).toThrow(FormatError);
    expect(() => readXlsx(patch(one, "[Content_Types].xml", (x) => x.replace("spreadsheetml.sheet.main", "wordprocessingml.document.main")))).toThrow(/not an XLSX/);
    expect(() => readXlsx(new TextEncoder().encode("PK nope"))).toThrow(FormatError);
  });

  test("9 a huge repeat count on a cell or a row is refused before it grows", () => {
    const cell = (extra: string) => `<table:table-row ${extra}><table:table-cell office:value-type="float" office:value="1" ${extra}/></table:table-row>`;
    expect(() => readOds(ods(table(cell('table:number-columns-repeated="1000000000"'))))).toThrow(/more than/);
    expect(() => readOds(ods(table(cell('table:number-rows-repeated="1000000000"'))))).toThrow(/more than/);
  });

  test("9 a repeat count on an empty cell or row is fine, as LibreOffice writes it", () => {
    const body = table(`<table:table-row><table:table-cell office:value-type="float" office:value="1"/><table:table-cell table:number-columns-repeated="1000000"/></table:table-row><table:table-row table:number-rows-repeated="1048575"><table:table-cell table:number-columns-repeated="16384"/></table:table-row>`);
    expect(readOds(ods(body)).model).toEqual(sheet("S", { A1: { value: 1 } }));
  });

  test("9 a cell off the grid, a shared string past the table, and too many cells", () => {
    expect(() => readXlsx(patch(one, "xl/worksheets/sheet1.xml", (x) => x.replace('r="A1"', 'r="XFE1"')))).toThrow(/sheet grid/);
    expect(() => readXlsx(patch(one, "xl/worksheets/sheet1.xml", (x) => x.replace('r="A1" t="inlineStr"><is><t xml:space="preserve">x</t></is>', 'r="A1" t="s"><v>9</v>')))).not.toThrow();
    const missing = readXlsx(patch(one, "xl/worksheets/sheet1.xml", (x) => x.replace('r="A1" t="inlineStr"><is><t xml:space="preserve">x</t></is>', 'r="A1" t="s"><v>9</v>')));
    expect(missing.warnings).toEqual(["a cell points to a missing shared string"]);
    expect(missing.model).toEqual(sheet("S", {}));
    const many = sheet("S", { A1: { value: 1 }, A2: { value: 2 }, A3: { value: 3 } });
    expect(() => readXlsx(writeXlsx(many).bytes, { maxCells: 2 })).toThrow(/more than 2 cells/);
    expect(() => readOds(writeOds(many).bytes, { maxCells: 2 })).toThrow(/more than 2 cells/);
  });

  test("11 one shared string used by many cells is refused, one use is fine", () => {
    const big = "x".repeat(50_000);
    const bytes = (uses: number) => {
      const files = unzipSync(one);
      const rels = strFromU8(files["xl/_rels/workbook.xml.rels"]).replace("</Relationships>", `<Relationship Id="rId9" Type="http://schemas.openxmlformats.org/officeDocument/2006/relationships/sharedStrings" Target="sharedStrings.xml"/></Relationships>`);
      const sheetXml = strFromU8(files["xl/worksheets/sheet1.xml"]).replace(/<sheetData>.*<\/sheetData>/, `<sheetData>${Array.from({ length: uses }, (_, i) => `<row r="${i + 1}"><c r="A${i + 1}" t="s"><v>0</v></c></row>`).join("")}</sheetData>`);
      const shared = `<?xml version="1.0"?><sst xmlns="http://schemas.openxmlformats.org/spreadsheetml/2006/main"><si><t>${big}</t></si></sst>`;
      return writeZip(
        Object.entries(files)
          .map(([name, data]) => ({ name, data: name === "xl/_rels/workbook.xml.rels" ? rels : name === "xl/worksheets/sheet1.xml" ? sheetXml : data }))
          .concat([{ name: "xl/sharedStrings.xml", data: shared }]),
        withLimits(),
      );
    };
    const caps = { maxTextChars: 100_000 };
    expect(readXlsx(bytes(2), caps).model.sheets[0].cells.A2).toEqual({ value: big });
    expect(() => readXlsx(bytes(3), caps)).toThrow(/more than 100000 characters of cell text/);
    expect(() => readXlsx(bytes(300), caps)).toThrow(FormatError);
  });

  test("11 a repeated cell or row counts its text in every copy", () => {
    const big = "y".repeat(50_000);
    const caps = { maxTextChars: 100_000 };
    const cell = `<table:table-cell office:value-type="string"><text:p>${big}</text:p></table:table-cell>`;
    expect(() => readOds(ods(table(`<table:table-row>${cell}</table:table-row>`)), caps)).not.toThrow();
    expect(() => readOds(ods(table(`<table:table-row>${cell.replace("<table:table-cell ", '<table:table-cell table:number-columns-repeated="3" ')}</table:table-row>`)), caps)).toThrow(/characters of cell text/);
    expect(() => readOds(ods(table(`<table:table-row table:number-rows-repeated="3">${cell}</table:table-row>`)), caps)).toThrow(/characters of cell text/);
    expect(() => readOds(ods(table(`<table:table-row table:number-rows-repeated="1000"><table:table-cell table:number-columns-repeated="1000" office:value-type="string" office:string-value="${big}"/></table:table-row>`)), caps)).toThrow(/characters of cell text/);
  });

  test("12 a formula past 8192 characters is refused fast, in both formats", () => {
    const started = Date.now();
    const cell = (formula: string) => `<table:table-row><table:table-cell office:value-type="float" office:value="1" table:formula="of:=${formula}"/></table:table-row>`;
    expect(() => readOds(ods(table(cell("[".repeat(1_000_000)))))).toThrow(/formula is longer than 8192/);
    expect(readOds(ods(table(cell("SUM([.A1];[.A2])")))).model.sheets[0].cells.A1).toMatchObject({ formula: "SUM(A1,A2)" });
    expect(() => readXlsx(patch(one, "xl/worksheets/sheet1.xml", (x) => x.replace(/<c r="A1"[^>]*>.*?<\/c>/, `<c r="A1"><f>${"A".repeat(9000)}</f></c>`)))).toThrow(/formula is longer than 8192/);
    expect(Date.now() - started).toBeLessThan(3000);
  });

  test("12 many formulas of near the cap that are built to make a regex backtrack are read in linear time", { timeout: 60_000 }, () => {
    const cell = (formula: string) => `<table:table-row><table:table-cell office:value-type="float" office:value="1" table:formula="of:=${formula.replaceAll('"', "&quot;")}"/></table:table-row>`;
    // Random letters keep the part from packing past the ratio cap.
    const letters = (n: number) => Array.from({ length: n }, () => String.fromCharCode(97 + Math.floor(Math.random() * 26)));
    const rows = (make: () => string) => table(Array.from({ length: 200 }, () => cell(make())).join(""));
    const unclosed = () => `[x];${letters(4050).map((c) => `[${c}`).join("")}`;
    const quotes = () => `"${letters(2700).map((c) => `${c}""`).join("")}`;
    for (const make of [unclosed, quotes]) {
      const started = Date.now();
      readOds(ods(rows(make)));
      expect(Date.now() - started).toBeLessThan(20_000);
    }
    expect(fromOpenFormula('IF([.A1]>0;"a;[.B2]";[.C3])')).toBe('IF(A1>0,"a;[.B2]",C3)');
    expect(fromOpenFormula("SUM([$'My Sheet'.A1:.B2])")).toBe("SUM('My Sheet'!A1:B2)");
    expect(fromOpenFormula("[.A1:.B2:.C3]")).toBeUndefined();
    expect(fromOpenFormula("[[.A1]")).toBe("[A1");
  });

  test("12 one cell of eight million newlines is charged before its XML is made", () => {
    const started = Date.now();
    expect(() => writeOds(sheet("S", { A1: { value: "\n".repeat(8_000_000) } }))).toThrow(/while it is written/);
    expect(Date.now() - started).toBeLessThan(3000);
    // A quote becomes six characters in XLSX, so five million of them are 30 MB.
    expect(() => writeXlsx(sheet("S", { A1: { value: '"'.repeat(5_000_000) } }), { maxEntryBytes: 20_000_000 })).toThrow(/while it is written/);
    expect(readOds(writeOds(sheet("S", { A1: { value: "a\nb\tc & d" } })).bytes).model.sheets[0].cells.A1).toMatchObject({ value: "a\nb\tc & d" });
  });

  test("12 many sheets that each repeat a column are refused fast, and a column style still reaches its cells", { timeout: 60_000 }, () => {
    const wide = `<table:table table:name="W"><table:table-column table:number-columns-repeated="16384"/></table:table>`;
    const started = Date.now();
    expect(() => readOds(ods(wide.repeat(1300)))).toThrow(/more than 1250 sheets/);
    expect(Date.now() - started).toBeLessThan(3000);
    expect(() => writeOds({ sheets: Array.from({ length: 1300 }, (_, i) => ({ name: `S${i}`, cells: {} })) })).toThrow(/at most 1250 sheets/);
    const styled = ods(
      `<style:style xmlns:style="urn:oasis:names:tc:opendocument:xmlns:style:1.0" style:name="pct" style:family="table-cell" style:data-style-name="N1"/>` +
        `<number:percentage-style xmlns:number="urn:oasis:names:tc:opendocument:xmlns:datastyle:1.0" xmlns:style="urn:oasis:names:tc:opendocument:xmlns:style:1.0" style:name="N1"><number:number number:decimal-places="0"/><number:text>%</number:text></number:percentage-style>` +
        `<table:table table:name="S"><table:table-column table:number-columns-repeated="2"/><table:table-column table:default-cell-style-name="pct"/><table:table-row><table:table-cell/><table:table-cell/><table:table-cell office:value-type="float" office:value="0.5"/></table:table-row></table:table>`,
    );
    expect(readOds(styled).model.sheets[0].cells.C1).toMatchObject({ value: 0.5, format: "0%" });
  });

  test("12 one number format that every cell points to counts for each cell", () => {
    const code = "0.00_" + "x".repeat(50_000);
    const xlsx = (cells: number) => {
      const base = writeXlsx(sheet("S", { A1: { value: 1, format: code } })).bytes;
      const sheetXml = (x: string) => x.replace(/<sheetData>.*<\/sheetData>/, `<sheetData>${Array.from({ length: cells }, (_, i) => `<row r="${i + 1}"><c r="A${i + 1}" s="1"/></row>`).join("")}</sheetData>`);
      return patch(base, "xl/worksheets/sheet1.xml", sheetXml);
    };
    const caps = { maxTextChars: 120_000 };
    expect(Object.keys(readXlsx(xlsx(2), caps).model.sheets[0].cells)).toHaveLength(2);
    expect(() => readXlsx(xlsx(3), caps)).toThrow(/characters of cell text/);
    expect(() => writeXlsx(sheet("S", Object.fromEntries(Array.from({ length: 3 }, (_, i) => [`A${i + 1}`, { value: 1, format: code }]))), caps)).toThrow(/characters of cell text/);
  });

  test("12 a cell style chain past the depth cap is refused, and a loop in the parents ends the walk", () => {
    const style = (name: string, parent?: string) => `<style:style xmlns:style="urn:oasis:names:tc:opendocument:xmlns:style:1.0" style:name="${name}" style:family="table-cell"${parent ? ` style:parent-style-name="${parent}"` : ""}/>`;
    const withCell = (styles: string, used: string) => ods(`${styles}${table(`<table:table-row><table:table-cell table:style-name="${used}" office:value-type="float" office:value="1"/></table:table-row>`)}`);
    const chain = Array.from({ length: 60 }, (_, i) => style(`s${i}`, i ? `s${i - 1}` : undefined)).join("");
    expect(() => readOds(withCell(chain, "s59"))).toThrow(/a style has more than 50 parents/);
    expect(readOds(withCell(chain, "s10")).model.sheets[0].cells.A1).toMatchObject({ value: 1 });
    expect(readOds(withCell(style("a", "b") + style("b", "a"), "a")).model.sheets[0].cells.A1).toMatchObject({ value: 1 });
  });

  test("12 text:s runs in a cell count against the text cap before the string is made", () => {
    const cell = (n: number) => `<table:table-row><table:table-cell office:value-type="string"><text:p>a${'<text:s text:c="1000"/>'.repeat(n)}b</text:p></table:table-cell></table:table-row>`;
    const caps = { maxTextChars: 10_000 };
    expect(readOds(ods(table(cell(9))), caps).model.sheets[0].cells.A1).toMatchObject({ value: `a${" ".repeat(9000)}b` });
    expect(() => readOds(ods(table(cell(11))), caps)).toThrow(/spaces in <text:s> elements/);
    const started = Date.now();
    expect(() => readOds(ods(table(cell(200_000))))).toThrow(FormatError);
    expect(Date.now() - started).toBeLessThan(5000);
    expect(readOds(ods(table(cell(0).replace("ab</text:p>", 'a<text:s text:c="5000"/>b</text:p>'))), {}).warnings).toEqual(["a run of more than 1000 spaces was shortened"]);
  });

  test("12 two sheet entries that name one worksheet part read it once, and the repeat warns", () => {
    const twice = patch(one, "xl/workbook.xml", (x) => x.replace("</sheets>", `<sheet name="T" sheetId="2" r:id="rId1"/></sheets>`));
    const read = readXlsx(twice);
    expect(read.model.sheets.map((s) => s.name)).toEqual(["S"]);
    expect(read.warnings).toEqual(['the sheet "T" shares its worksheet part with an earlier sheet: not imported']);
  });

  test("13 content the XLSX reader drops is named in a warning", () => {
    const sheetXml = (x: string, ...changes: [string, string][]) => changes.reduce((acc, [a, b]) => acc.replace(a, b), x);
    const warnings = (bytes: Uint8Array) => readXlsx(bytes).warnings;
    const withSheet = (add: string) => patch(one, "xl/worksheets/sheet1.xml", (x) => x.replace("</worksheet>", `${add}</worksheet>`));
    expect(warnings(withSheet('<drawing r:id="rId9"/>'))).toEqual(["charts, images, or comments were not read"]);
    expect(warnings(withSheet('<conditionalFormatting sqref="A1"/>'))).toEqual(["conditional formats or data validation rules were not read"]);
    expect(warnings(patch(one, "xl/workbook.xml", (x) => x.replace("</sheets>", '</sheets><definedNames><definedName name="n">S!A1</definedName></definedNames>')))).toEqual(["named ranges, print areas, or filters were not read"]);
    expect(warnings(patch(one, "xl/workbook.xml", (x) => sheetXml(x, ["<sheet ", '<sheet state="hidden" ']))).slice(0, 1)).toEqual(["a hidden sheet was read as visible"]);
    expect(warnings(patch(one, "xl/worksheets/sheet1.xml", (x) => x.replace('r="A1" t="inlineStr"><is><t xml:space="preserve">x</t></is>', 'r="A1"><v>abc</v>')))).toEqual(["a cell value that is not a number or a date was not read"]);
    expect(warnings(patch(one, "xl/worksheets/sheet1.xml", (x) => x.replace('r="A1"', 'r="A1" s="99"')))).toEqual(["a cell points to a missing style"]);
    expect(warnings(patch(one, "xl/worksheets/sheet1.xml", (x) => x.replace('r="A1" t="inlineStr"><is><t xml:space="preserve">x</t></is>', 'r="A1"><f t="array" ref="A1">1+1</f><v>2</v>')))).toEqual(["an array formula or data table was read as a plain formula"]);
  });

  test("13 one unknown format id used by many styles counts once past the cap", () => {
    const xfs = [...Array.from({ length: 50 }, (_, i) => 100 + i), ...Array.from({ length: 1000 }, () => 999)].map((id) => `<xf numFmtId="${id}"/>`).join("");
    const read = readXlsx(patch(one, "xl/styles.xml", (x) => x.replace("</cellXfs>", `${xfs}</cellXfs>`)));
    expect(read.warnings).toHaveLength(51);
    expect(read.warnings[50]).toBe("1 more warnings");
  });

  test("13 rich text in an inline string is named in a warning", () => {
    const read = readXlsx(patch(one, "xl/worksheets/sheet1.xml", (x) => x.replace('<is><t xml:space="preserve">x</t></is>', "<is><r><rPr><b/></rPr><t>x</t></r></is>")));
    expect(read.model).toEqual(sheet("S", { A1: { value: "x" } }));
    expect(read.warnings).toEqual(["text styles inside a cell were not read"]);
  });

  test("13 names past the cap never hide a fixed warning, and a repeated fixed warning is one line", () => {
    const rels = Array.from({ length: 60 }, (_, i) => `<Relationship Id="rIdX${i}" Type="http://schemas.openxmlformats.org/officeDocument/2006/relationships/chartsheet" Target="c${i}.xml"/>`).join("");
    const entries = Array.from({ length: 60 }, (_, i) => `<sheet name="c${i}" sheetId="${i + 10}" r:id="rIdX${i}"/>`).join("");
    const cells = Array.from({ length: 1000 }, (_, i) => `<c r="B${i + 1}" s="99"><v>1</v></c>`).join("");
    const bytes = patch(patch(patch(one, "xl/workbook.xml", (x) => x.replace("</sheets>", `${entries}</sheets>`)), "xl/_rels/workbook.xml.rels", (x) => x.replace("</Relationships>", `${rels}</Relationships>`)), "xl/worksheets/sheet1.xml", (x) => x.replace("</row>", `${cells}</row>`).replace("</sheetData>", "</sheetData><mergeCells/>"));
    const { warnings } = readXlsx(bytes);
    expect(warnings).toContain("a cell points to a missing style");
    expect(warnings).toContain("merged cells were read as single cells");
    expect(warnings.filter((w) => w === "a cell points to a missing style")).toHaveLength(1);
    expect(warnings).toContain("10 more warnings");
  });

  test("13 rich text runs in a shared string are named in a warning", () => {
    const bytes = patch(patch(one, "xl/worksheets/sheet1.xml", (x) => x.replace('r="A1" t="inlineStr"><is><t xml:space="preserve">x</t></is>', 'r="A1" t="s"><v>0</v>')), "xl/_rels/workbook.xml.rels", (x) => x.replace("</Relationships>", '<Relationship Id="rId8" Type="http://schemas.openxmlformats.org/officeDocument/2006/relationships/sharedStrings" Target="sharedStrings.xml"/></Relationships>'));
    const withStrings = writeZip([...Object.entries(unzipSync(bytes)).map(([name, data]) => ({ name, data, store: name === "mimetype" })), { name: "xl/sharedStrings.xml", data: '<sst xmlns="http://schemas.openxmlformats.org/spreadsheetml/2006/main"><si><r><rPr><b/></rPr><t>bold</t></r></si></sst>', store: false }], withLimits());
    const read = readXlsx(withStrings);
    expect(read.model).toEqual(sheet("S", { A1: { value: "bold" } }));
    expect(read.warnings).toEqual(["text styles inside a cell were not read"]);
  });

  test("13 many warnings are capped, with one line that counts the rest, and a long sheet name is cut", () => {
    const entries = Array.from({ length: 300 }, (_, i) => `<sheet name="${i}${"n".repeat(500)}" sheetId="${i + 10}" r:id="rIdX${i}"/>`).join("");
    const read = readXlsx(patch(patch(one, "xl/workbook.xml", (x) => x.replace("</sheets>", `${entries}</sheets>`)), "xl/_rels/workbook.xml.rels", (x) => x.replace("</Relationships>", `${Array.from({ length: 300 }, (_, i) => `<Relationship Id="rIdX${i}" Type="http://schemas.openxmlformats.org/officeDocument/2006/relationships/chartsheet" Target="c${i}.xml"/>`).join("")}</Relationships>`)));
    expect(read.warnings).toHaveLength(51);
    expect(read.warnings[50]).toBe("250 more warnings");
    expect(read.warnings.slice(0, 50).every((w) => w.startsWith("the sheet "))).toBe(true);
    expect(Math.max(...read.warnings.map((w) => w.length))).toBeLessThan(120);
  });

  test("12 many distinct number formats write fast, and the ODS warnings stay short", () => {
    const cells = Object.fromEntries(Array.from({ length: 20_000 }, (_, i) => [`A${i + 1}`, { value: 1, format: `0.00_${i}` }]));
    const started = Date.now();
    expect(readXlsx(writeXlsx(sheet("S", cells)).bytes).model.sheets[0].cells.A20000).toMatchObject({ format: "0.00_19999" });
    const { warnings } = writeOds(sheet("S", cells));
    expect(Date.now() - started).toBeLessThan(5000);
    expect(warnings.length).toBeLessThanOrEqual(101);
    expect(warnings).toContain("more number formats are not written in ODS");
  });

  test("9 a relationship that leaves the package, or points nowhere, is refused or skipped", () => {
    const escape = patch(one, "xl/_rels/workbook.xml.rels", (x) => x.replace('Target="worksheets/sheet1.xml"', 'Target="../../etc/passwd"'));
    expect(() => readXlsx(escape)).toThrow(/no sheet/);
    const unknown = patch(one, "xl/_rels/workbook.xml.rels", (x) => x.replace('Id="rId1"', 'Id="rId9"'));
    expect(() => readXlsx(unknown)).toThrow(/no sheet/);
  });

  test("9 an external link and an external formula reference are text, never fetched", () => {
    const book = sheet("S", { A1: { value: 1, formula: "'[1]Other'!A1" }, A2: { value: "x", formula: 'HYPERLINK("http://evil.example/a","x")' } });
    const back = readXlsx(writeXlsx(book).bytes);
    expect(back.model).toEqual(book);
  });
});

describe.skipIf(!hasLibreOffice && !libreOfficeRequired)("LibreOffice round trip", () => {
  const lo = hasLibreOffice ? new LibreOffice() : undefined;
  const records: RoundTripRecord[] = [];
  afterAll(() => {
    lo?.dispose();
    writeArtifact("spreadsheet-roundtrip", records);
  });

  test.each(formats)("10 LibreOffice opens our %s, prints it, and its re-save reads back to the same model", { timeout: 300_000 }, (format) => {
    expect(hasLibreOffice).toBe(true);
    const { bytes } = writers[format](fixture);
    const pdf = lo!.convert(bytes, format, "pdf");
    expect(strFromU8(pdf.subarray(0, 5))).toBe("%PDF-");
    const text = strFromU8(lo!.convert(bytes, format, "csv"));
    expect(text).toContain("far");
    const resaved = lo!.convert(bytes, format, format);
    const back = readers[format](resaved);
    records.push({ format, case: "fixture", written: sha256(bytes), resaved: sha256(resaved), pdfBytes: pdf.length, equal: isDeepStrictEqual(back.model, fixture) });
    expect(back.model).toEqual(fixture);
  });

  test.each([
    ["xlsx", "ods"],
    ["ods", "xlsx"],
  ] as const)("10 %s through LibreOffice to %s reads back to the same model", { timeout: 300_000 }, (from, to) => {
    const { bytes } = writers[from](fixture);
    const converted = lo!.convert(bytes, from, to);
    const back = readers[to](converted);
    records.push({ format: `${from}->${to}`, case: "fixture", written: sha256(bytes), resaved: sha256(converted), pdfBytes: 0, equal: isDeepStrictEqual(back.model, fixture) });
    expect(back.model).toEqual(fixture);
  });

  test.each(formats)("10 LibreOffice computes a %s formula that has no cached result", { timeout: 300_000 }, (format) => {
    const book = sheet("S", { A1: { value: 20 }, A2: { value: 22 }, A3: { formula: "SUM(A1:A2)" } });
    const csv = strFromU8(lo!.convert(writers[format](book).bytes, format, "csv"));
    expect(csv.trim().split("\n")).toEqual(["20", "22", "42"]);
  });
});
