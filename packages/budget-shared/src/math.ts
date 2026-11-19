/**
 * Inline math parser and cost-formula builder. No `eval` / `Function`.
 * Cost formulas written to the sheet must only come from `formatSplitCostFormula`
 * (AGENTS.md section 1.4).
 */

class ParseError extends Error {}

/**
 * Evaluates `+ - * /`, parentheses, unary minus/plus, decimals (including `.5`)
 * and whitespace. Returns null on any invalid input, division by zero, or a
 * non-finite result.
 */
export function evaluateMathExpression(expr: string): number | null {
  if (typeof expr !== "string") return null;
  const src = expr;
  let pos = 0;

  const fail = (): never => {
    throw new ParseError("invalid expression");
  };
  const skipWs = () => {
    while (pos < src.length && /\s/.test(src[pos])) pos++;
  };

  function parseNumber(): number {
    const match = /^(?:\d+(?:\.\d+)?|\.\d+)/.exec(src.slice(pos));
    if (!match) return fail();
    pos += match[0].length;
    return Number(match[0]);
  }

  function parseFactor(): number {
    skipWs();
    const ch = src[pos];
    if (ch === "-") {
      pos++;
      return -parseFactor();
    }
    if (ch === "+") {
      pos++;
      return parseFactor();
    }
    if (ch === "(") {
      pos++;
      const value = parseExpression();
      skipWs();
      if (src[pos] !== ")") return fail();
      pos++;
      return value;
    }
    return parseNumber();
  }

  function parseTerm(): number {
    let value = parseFactor();
    for (;;) {
      skipWs();
      const op = src[pos];
      if (op !== "*" && op !== "/") return value;
      pos++;
      const rhs = parseFactor();
      if (op === "*") {
        value *= rhs;
      } else {
        if (rhs === 0) return fail();
        value /= rhs;
      }
    }
  }

  function parseExpression(): number {
    let value = parseTerm();
    for (;;) {
      skipWs();
      const op = src[pos];
      if (op !== "+" && op !== "-") return value;
      pos++;
      const rhs = parseTerm();
      value = op === "+" ? value + rhs : value - rhs;
    }
  }

  try {
    const value = parseExpression();
    skipWs();
    if (pos !== src.length) return null;
    return Number.isFinite(value) ? value : null;
  } catch (e) {
    if (e instanceof ParseError) return null;
    throw e;
  }
}

const NUMERIC_LITERAL = /^-?(?:\d+(?:\.\d+)?|\.\d+)$/;

/**
 * Builds the value to write in a Cost cell.
 *
 * Trims, strips one leading `=`, validates with `evaluateMathExpression` (throws
 * on invalid input). With `applyTax`: `=${taxMultiplier}*(${expr})`. Without:
 * the number for a single numeric literal (including negative), else `=${expr}`
 * with no outer parentheses added.
 */
export function formatSplitCostFormula(
  expr: string,
  applyTax: boolean,
  taxMultiplier: number,
): string | number {
  let cleaned = String(expr).trim();
  if (cleaned.startsWith("=")) cleaned = cleaned.slice(1).trim();
  if (evaluateMathExpression(cleaned) === null) {
    throw new Error(`Invalid cost expression: ${expr}`);
  }
  if (applyTax) return `=${taxMultiplier}*(${cleaned})`;
  if (NUMERIC_LITERAL.test(cleaned)) return Number(cleaned);
  return `=${cleaned}`;
}

/** A plain amount: optional leading "-", digits, and at most one "." (e.g. `12`, `-3.5`, `.25`, `7.`). */
const PLAIN_AMOUNT = /^-?(?:\d+\.?\d*|\.\d+)$/;

/** True when `text` (trimmed) is a plain amount: no letters, operators, "$" or thousands commas. */
export function isPlainAmount(text: string): boolean {
  return typeof text === "string" && PLAIN_AMOUNT.test(text.trim());
}

/**
 * The plain amount exactly as typed (trimmed, a trailing "." dropped), ready to send as a cost
 * expression, or null when it is not a plain amount. `"123288.0288"` stays `"123288.0288"`.
 */
export function normalizePlainAmount(text: string): string | null {
  if (!isPlainAmount(text)) return null;
  return text.trim().replace(/\.$/, "");
}

/** One subtracted child reference of a split parent formula: R1C1 `-R[n]C[0]` or A1 `-C12`. */
const SPLIT_REFS = String.raw`((?:\s*-\s*(?:R\[[1-9]\d*\]C(?:\[0\])?|\$?[A-Za-z]{1,3}\$?\d+))+)`;
const SPLIT_PARENT_PARTS = new RegExp(
  String.raw`^=\s*(.*?)` + SPLIT_REFS + String.raw`\s*$`,
);

/**
 * Splits a split parent's cost formula (`=45-R[1]C[0]-R[2]C[0]`, or the A1 form `=45-C12-C13`
 * that `getFormulas` returns) into its total expression (`45`) and the subtracted child
 * references. `total` is the evaluated expression, or null when it is not plain math.
 * Returns null when the formula does not end in child references.
 */
export function splitParentFormulaParts(
  formula: string | null | undefined,
): { totalExpression: string; total: number | null; refs: string } | null {
  if (typeof formula !== "string") return null;
  const m = SPLIT_PARENT_PARTS.exec(formula.trim());
  if (!m || m[1].trim() === "") return null;
  const totalExpression = m[1].trim();
  return {
    totalExpression,
    total: evaluateMathExpression(totalExpression),
    refs: m[2],
  };
}

/**
 * A split parent's cost formula with its total (the leading value before the child references)
 * replaced by the plain amount `total`. Throws when `total` is not a plain amount or the formula
 * is not a split parent formula (AGENTS.md section 1.4: only validated input becomes a formula).
 */
export function buildSplitParentFormula(
  formula: string,
  total: string,
): string {
  const amount = normalizePlainAmount(total);
  if (amount === null || evaluateMathExpression(amount) === null)
    throw new Error(`Invalid split total: ${total}`);
  const parts = splitParentFormulaParts(formula);
  if (!parts) throw new Error(`Not a split parent formula: ${formula}`);
  return `=${amount}${parts.refs.replace(/\s+/g, "")}`;
}
