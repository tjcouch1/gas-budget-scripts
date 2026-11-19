/**
 * Inline math parser and cost-formula builder. No `eval` / `Function`.
 * Cost formulas written to the sheet must only come from `formatSplitCostFormula`
 * (AGENTS.md section 1.4).
 */

class ParseError extends Error {}

/**
 * Evaluates `+ - * /`, parentheses, one unary minus/plus (not right after another sign:
 * `--5`, `1+-2` are invalid; `2*-3` is fine), decimals (including `.5`)
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

  /** `signAllowed`: false right after a "+" or "-", so doubled signs (`--5`, `1+-2`) fail. */
  function parseFactor(signAllowed: boolean): number {
    skipWs();
    const ch = src[pos];
    if (ch === "-" || ch === "+") {
      if (!signAllowed) return fail();
      pos++;
      const value = parseFactor(false);
      return ch === "-" ? -value : value;
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

  function parseTerm(signAllowed: boolean): number {
    let value = parseFactor(signAllowed);
    for (;;) {
      skipWs();
      const op = src[pos];
      if (op !== "*" && op !== "/") return value;
      pos++;
      const rhs = parseFactor(true);
      if (op === "*") {
        value *= rhs;
      } else {
        if (rhs === 0) return fail();
        value /= rhs;
      }
    }
  }

  function parseExpression(): number {
    let value = parseTerm(true);
    for (;;) {
      skipWs();
      const op = src[pos];
      if (op !== "+" && op !== "-") return value;
      pos++;
      const rhs = parseTerm(false);
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

/**
 * An amount field's input, detected automatically: a plain numeric literal (`isPlainAmount`) is
 * kept exactly as typed (a trailing "." dropped) and is written as a number; anything else must
 * be math the parser accepts and is written as a formula (`formatSplitCostFormula`). One leading
 * "=" is ignored. `expression` is what to send as a cost expression.
 */
export type AmountInput =
  | { ok: true; plain: boolean; expression: string; value: number }
  | { ok: false; reason: "empty" | "invalid" };

export function parseAmountInput(text: string): AmountInput {
  let expression = String(text).trim();
  if (expression.startsWith("=")) expression = expression.slice(1).trim();
  if (expression === "") return { ok: false, reason: "empty" };
  const plain = normalizePlainAmount(expression);
  if (plain !== null) expression = plain;
  const value = evaluateMathExpression(expression);
  if (value === null) return { ok: false, reason: "invalid" };
  return { ok: true, plain: plain !== null, expression, value };
}

/** One subtracted child reference of a split parent formula: R1C1 `-R[n]C[0]` or A1 `-C12`. */
const SPLIT_REFS = String.raw`((?:\s*-\s*(?:R\[[1-9]\d*\]C(?:\[0\])?|\$?[A-Za-z]{1,3}\$?\d+))+)`;
const SPLIT_PARENT_PARTS = new RegExp(
  String.raw`^=\s*(.*?)` + SPLIT_REFS + String.raw`\s*$`,
);

/** `expr` without one pair of parentheses around all of it (`(40+5)` -> `40+5`; not `(1)*(2)`). */
function stripOuterParens(expr: string): string {
  const text = expr.trim();
  if (!text.startsWith("(") || !text.endsWith(")")) return text;
  let depth = 0;
  for (let i = 0; i < text.length; i++) {
    if (text[i] === "(") depth++;
    else if (text[i] === ")") depth--;
    // The first "(" closes before the end: it does not wrap everything
    if (depth === 0 && i < text.length - 1) return text;
  }
  return text.slice(1, -1).trim();
}

/**
 * Splits a split parent's cost formula (`=45-R[1]C[0]-R[2]C[0]`, the A1 form `=45-C12-C13` that
 * `getFormulas` returns, or `=(40+5)-R[1]C-...`) into its total expression and the subtracted
 * child references. The native split appends `-R[n]C[0]` to an existing formula, so the total is
 * everything before the trailing references (`=10+5-R[1]C[0]`: `10+5`). `total` is its value and
 * `totalText` the text to show and edit (without the parentheses `buildSplitParentFormula` adds);
 * both are null when the total is not math the parser accepts (show it read-only). Returns null
 * when the formula does not end in child references.
 */
export function splitParentFormulaParts(formula: string | null | undefined): {
  totalExpression: string;
  totalText: string | null;
  total: number | null;
  refs: string;
} | null {
  if (typeof formula !== "string") return null;
  const m = SPLIT_PARENT_PARTS.exec(formula.trim());
  if (!m || m[1].trim() === "") return null;
  const totalExpression = m[1].trim();
  const total = evaluateMathExpression(totalExpression);
  return {
    totalExpression,
    totalText: total === null ? null : stripOuterParens(totalExpression),
    total,
    refs: m[2],
  };
}

/**
 * A split parent's cost formula with its total (everything before the child references)
 * replaced by `total`, detected like `parseAmountInput`: a plain amount is written as typed
 * (`=45-R[1]C[0]`), math in parentheses (`=(40+5)-R[1]C[0]`). Throws when `total` is invalid or
 * the formula is not a split parent formula (AGENTS.md section 1.4: only validated input becomes
 * a formula).
 */
export function buildSplitParentFormula(
  formula: string,
  total: string,
): string {
  const amount = parseAmountInput(total);
  if (!amount.ok) throw new Error(`Invalid split total: ${total}`);
  const parts = splitParentFormulaParts(formula);
  if (!parts) throw new Error(`Not a split parent formula: ${formula}`);
  const head = amount.plain
    ? amount.expression
    : `(${stripOuterParens(amount.expression)})`;
  return `=${head}${parts.refs.replace(/\s+/g, "")}`;
}
