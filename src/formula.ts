/*
 * Formulas are text. The model holds A1 notation with commas: `SUM(A1:B2,Sheet2!C3)`.
 * XLSX stores that text as it is. ODS stores OpenFormula: `of:=SUM([.A1:.B2];[$Sheet2.C3])`.
 * These two functions change only references and argument separators, and never touch a string literal.
 */

const STRING = `"(?:[^"]|"")*"`;
const SHEET = `(?:'(?:[^']|'')+'|[A-Za-z_][A-Za-z0-9_.]*)`;
const CELL = `\\$?[A-Z]{1,3}\\$?[0-9]+`;
// A reference is not part of a longer name, and a name before `(` is a function such as LOG10.
const A1_REF = `(?<![A-Za-z0-9_.$'])(?:(${SHEET})!)?(${CELL})(?::(${CELL}))?(?![A-Za-z0-9_(])`;
const A1 = new RegExp(`${STRING}|${A1_REF}|,`, "g");
const OF = new RegExp(`${STRING}|\\[([^\\]]*)\\]|;`, "g");
const OF_PART = new RegExp(`^(?:\\$?(${SHEET}))?\\.(${CELL})$`);

const unquote = (name: string): string => (name.startsWith("'") ? name.slice(1, -1).replaceAll("''", "'") : name);
const quote = (name: string): string => (/^[A-Za-z_][A-Za-z0-9_.]*$/.test(name) ? name : `'${name.replaceAll("'", "''")}'`);

/** A1 text to OpenFormula text, without the `of:=` prefix. */
export function toOpenFormula(formula: string): string {
  return formula.replace(A1, (match, sheet?: string, from?: string, to?: string) => {
    if (match === ",") return ";";
    if (from === undefined) return match;
    const at = sheet ? `$${sheet}` : "";
    return `[${at}.${from}${to ? `:.${to}` : ""}]`;
  });
}

/** OpenFormula text, without the `of:=` prefix, to A1 text. Returns undefined for a reference it cannot map. */
export function fromOpenFormula(formula: string): string | undefined {
  let failed = false;
  const text = formula.replace(OF, (match, ref?: string) => {
    if (match === ";") return ",";
    if (ref === undefined) return match;
    const parts = ref.split(":");
    const mapped = parts.map((p) => OF_PART.exec(p));
    if (parts.length > 2 || mapped.some((m) => !m)) {
      failed = true;
      return match;
    }
    const [first, second] = mapped as RegExpExecArray[];
    const sheet = first[1] ? `${quote(unquote(first[1]))}!` : "";
    return `${sheet}${first[2]}${second ? `:${second[2]}` : ""}`;
  });
  return failed ? undefined : text;
}
