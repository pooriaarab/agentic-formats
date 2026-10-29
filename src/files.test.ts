import { linkSync, mkdirSync, mkdtempSync, readFileSync, readdirSync, rmSync, statSync, symlinkSync, writeFileSync, chmodSync } from "node:fs";
import { tmpdir } from "node:os";
import { join } from "node:path";
import { afterAll, describe, expect, test } from "vitest";
import { UsageError, writeOutput } from "./files.js";

const dir = mkdtempSync(join(tmpdir(), "files-"));
afterAll(() => rmSync(dir, { recursive: true, force: true }));
const bytes = (s: string) => new TextEncoder().encode(s);

describe("writeOutput is atomic", () => {
  test("a replace leaves no temporary file and keeps the mode", () => {
    const d = join(dir, "a");
    mkdirSync(d);
    writeFileSync(join(d, "out.bin"), "old");
    chmodSync(join(d, "out.bin"), 0o640);
    writeOutput(join(d, "out.bin"), bytes("new"), true);
    expect(readFileSync(join(d, "out.bin"), "utf8")).toBe("new");
    expect(statSync(join(d, "out.bin")).mode & 0o777).toBe(0o640);
    expect(readdirSync(d)).toEqual(["out.bin"]);
  });

  test("a failed write leaves the old file whole and no temporary file", () => {
    const d = join(dir, "b");
    mkdirSync(d);
    writeFileSync(join(d, "out.bin"), "old");
    expect(() => writeOutput(join(d, "out.bin"), { length: 3 } as unknown as Uint8Array, true)).toThrow();
    expect(readFileSync(join(d, "out.bin"), "utf8")).toBe("old");
    expect(readdirSync(d)).toEqual(["out.bin"]);
  });

  test("a second name for the file is not written through", () => {
    const d = join(dir, "c");
    mkdirSync(d);
    writeFileSync(join(d, "victim"), "original");
    linkSync(join(d, "victim"), join(d, "out.bin"));
    writeOutput(join(d, "out.bin"), bytes("new"), true);
    expect(readFileSync(join(d, "victim"), "utf8")).toBe("original");
    expect(readFileSync(join(d, "out.bin"), "utf8")).toBe("new");
  });

  test("a link refused up front, and a link swapped in later is replaced and never followed", () => {
    const d = join(dir, "d");
    mkdirSync(d);
    writeFileSync(join(d, "victim"), "original");
    symlinkSync(join(d, "victim"), join(d, "out.bin"));
    for (const overwrite of [true, false]) expect(() => writeOutput(join(d, "out.bin"), bytes("x"), overwrite)).toThrow(/is a link/);
    expect(readFileSync(join(d, "victim"), "utf8")).toBe("original");
  });

  test("without overwrite, the name taken by another writer fails and keeps that file", () => {
    const d = join(dir, "e");
    mkdirSync(d);
    writeOutput(join(d, "out.bin"), bytes("first"), false);
    expect(() => writeOutput(join(d, "out.bin"), bytes("second"), false)).toThrow(UsageError);
    expect(readFileSync(join(d, "out.bin"), "utf8")).toBe("first");
    expect(readdirSync(d)).toEqual(["out.bin"]);
  });
});
