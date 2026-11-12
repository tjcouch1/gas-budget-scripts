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
