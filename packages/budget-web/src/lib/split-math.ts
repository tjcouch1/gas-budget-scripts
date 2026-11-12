import {
  evaluateMathExpression,
  isUnsafeFreeText,
  type SplitItemInput,
} from "@budget/shared";

export interface SplitDraftItem {
  /** Stable key for React lists. */
  key: string;
  itemNote: string;
  expression: string;
  preTax: boolean;
  envelope: string;
}

export interface EvaluatedSplitItem {
  /** Tax-applied amount, or null when the expression is invalid. */
  value: number | null;
  noteUnsafe: boolean;
}

const roundCents = (n: number) => Math.round(n * 100) / 100;

/** The amount a child cell will hold: the expression, times TaxMultiplier when Pre-Tax is on. */
export function evaluateSplitItem(
  item: Pick<SplitDraftItem, "expression" | "preTax" | "itemNote">,
  taxMultiplier: number,
): EvaluatedSplitItem {
  let expr = item.expression.trim();
  if (expr.startsWith("=")) expr = expr.slice(1).trim();
  const base = evaluateMathExpression(expr);
  return {
    value: base === null ? null : item.preTax ? base * taxMultiplier : base,
    noteUnsafe: isUnsafeFreeText(item.itemNote),
  };
}

export interface SplitSummary {
  items: EvaluatedSplitItem[];
  /** Parent value minus every child value; invalid items count as zero. */
  remaining: number;
  allValid: boolean;
  /** True when the split may be submitted. */
  canSubmit: boolean;
}

/** Live "Remaining to allocate" and validity for the split editor. */
export function summarizeSplit(
  parentValue: number | null,
  items: SplitDraftItem[],
  taxMultiplier: number,
): SplitSummary {
  const evaluated = items.map((i) => evaluateSplitItem(i, taxMultiplier));
  const allValid =
    evaluated.length > 0 &&
    evaluated.every((e) => e.value !== null && !e.noteUnsafe);
  const sum = evaluated.reduce((s, e) => s + (e.value ?? 0), 0);
  const remaining = roundCents((parentValue ?? 0) - sum);
  return {
    items: evaluated,
    remaining,
    allValid,
    // A negative remainder is almost certainly a typo; block it.
    canSubmit: allValid && remaining >= -0.005,
  };
}

export function toSplitInputs(items: SplitDraftItem[]): SplitItemInput[] {
  return items.map((i) => ({
    itemNote: i.itemNote.trim(),
    expression: i.expression.trim(),
    preTax: i.preTax,
    envelope: i.envelope,
  }));
}
