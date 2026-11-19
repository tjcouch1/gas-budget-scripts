/**
 * Compares the text of two cell formulas, e.g. a row the web app expects with the row the sheet
 * now holds. Pure (no GAS or DOM globals).
 */

type Token = { kind: "r1c1" | "a1" | "number" | "other"; text: string };

const R1C1_REF = /^R(?:\[(-?\d+)\]|(\d+))?C(?:\[(-?\d+)\]|(\d+))?/;
const A1_REF = /^\$?([A-Z]{1,3})\$?(\d+)/;
const NUMBER = /^(?:\d+(?:\.\d*)?|\.\d+)/;

/** Tokens of a formula: case, spaces, `$` and the spelling of numbers and `C[0]` do not matter. */
function tokenize(formula: string): Token[] {
  const src = formula
    .trim()
    .replace(/^=/, "")
    .replace(/\s+/g, "")
    .toUpperCase();
  const tokens: Token[] = [];
  let pos = 0;
  while (pos < src.length) {
    const rest = src.slice(pos);
    let m = R1C1_REF.exec(rest);
    if (m && m[0].length > 1) {
      const row = m[1] !== undefined ? `[${m[1]}]` : (m[2] ?? "[0]");
      const col = m[3] !== undefined ? `[${m[3]}]` : (m[4] ?? "[0]");
      tokens.push({ kind: "r1c1", text: `R${row}C${col}` });
      pos += m[0].length;
      continue;
    }
    m = A1_REF.exec(rest);
    if (m) {
      tokens.push({ kind: "a1", text: `${m[1]}${Number(m[2])}` });
      pos += m[0].length;
      continue;
    }
    m = NUMBER.exec(rest);
    if (m) {
      tokens.push({ kind: "number", text: String(Number(m[0])) });
      pos += m[0].length;
      continue;
    }
    tokens.push({ kind: "other", text: src[pos] });
    pos++;
  }
  return tokens;
}

/**
 * `"same"` when both formulas have the same text (ignoring case, spaces, `$`, `.5` vs `0.5`,
 * `R[1]C` vs `R[1]C[0]`); `"different"` when they differ; `"refNotationDiffers"` when they are
 * the same except that references are in A1 notation on one side and R1C1 on the other at the
 * same places (`=45-C12` vs `=45-R[1]C[0]`): converting needs the cell's position, so the caller
 * must decide by the values. Real Sheets `getFormulas` returns A1 and `getFormulasR1C1` R1C1.
 */
export function compareFormulas(
  a: string,
  b: string,
): "same" | "different" | "refNotationDiffers" {
  const ta = tokenize(a);
  const tb = tokenize(b);
  if (ta.length !== tb.length) return "different";
  let notationDiffers = false;
  for (let i = 0; i < ta.length; i++) {
    const x = ta[i];
    const y = tb[i];
    if (x.kind === y.kind && x.text === y.text) continue;
    const refs = new Set([x.kind, y.kind]);
    if (refs.has("a1") && refs.has("r1c1")) {
      notationDiffers = true;
      continue;
    }
    return "different";
  }
  return notationDiffers ? "refNotationDiffers" : "same";
}
