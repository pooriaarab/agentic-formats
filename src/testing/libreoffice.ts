import { spawnSync } from "node:child_process";
import { createHash } from "node:crypto";
import { mkdirSync, mkdtempSync, readdirSync, readFileSync, rmSync, writeFileSync } from "node:fs";
import { tmpdir } from "node:os";
import { join } from "node:path";

/** Test kit: drive the real LibreOffice, and keep a repeatable record of each round trip. */

export const hasLibreOffice = spawnSync("soffice", ["--version"], { encoding: "utf8" }).status === 0;
/** CI sets this, so a missing LibreOffice fails the run instead of skipping it. */
export const libreOfficeRequired = process.env.REQUIRE_LIBREOFFICE === "1";

export const sha256 = (bytes: Uint8Array): string => createHash("sha256").update(bytes).digest("hex");

export class LibreOffice {
  private readonly dir = mkdtempSync(join(tmpdir(), "formats-lo-"));
  private count = 0;

  /** Convert a file with LibreOffice. `to` is a filter name such as `odt`, `pdf`, or `txt:Text`. */
  convert(bytes: Uint8Array, ext: string, to: string): Uint8Array {
    const work = join(this.dir, `job${++this.count}`);
    mkdirSync(join(work, "out"), { recursive: true });
    const input = join(work, `input.${ext}`);
    writeFileSync(input, bytes);
    const run = spawnSync(
      "soffice",
      ["--headless", "--norestore", `-env:UserInstallation=file://${join(this.dir, "profile")}`, "--convert-to", to, "--outdir", join(work, "out"), input],
      { encoding: "utf8", timeout: 170_000, env: { ...process.env, HOME: this.dir } },
    );
    const produced = readdirSync(join(work, "out"));
    if (run.status !== 0 || produced.length !== 1) throw new Error(`LibreOffice could not convert ${ext} to ${to}: ${run.stderr || run.stdout}`);
    const out = readFileSync(join(work, "out", produced[0]));
    rmSync(work, { recursive: true, force: true });
    return new Uint8Array(out);
  }

  dispose(): void {
    rmSync(this.dir, { recursive: true, force: true });
  }
}

export interface RoundTripRecord {
  format: string;
  case: string;
  /** SHA-256 of the file our writer made. */
  written: string;
  /** SHA-256 of the file LibreOffice saved. */
  resaved: string;
  pdfBytes: number;
  equal: boolean;
}

/** Write the records to `artifacts/<name>.json`. The same code and LibreOffice give the same file. */
export function writeArtifact(name: string, records: RoundTripRecord[]): void {
  mkdirSync("artifacts", { recursive: true });
  writeFileSync(join("artifacts", `${name}.json`), `${JSON.stringify({ libreOffice: spawnSync("soffice", ["--version"], { encoding: "utf8" }).stdout.trim(), records }, null, 2)}\n`);
}
