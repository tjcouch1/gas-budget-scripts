import {
  isUnsafeFreeText,
  parseAmountInput,
  type SplitItemInput,
} from "@budget/shared";
import { AMOUNT_MESSAGE } from "./row-draft.ts";

export interface SplitDraftItem {
  /** Stable key for React lists. */
  key: string;
  itemNote: string;
  expression: string;
  preTax: boolean;
  envelope: string;
  /** "" copies the original row's Paid with. */
  paidWith: string;
  taxDeductible: boolean;
  /** The new row's Name cell note; blank: none. */
  note: string;
}

let keyCounter = 0;

/**
 * A blank new split item; "Add tax" starts checked (most receipts list pre-tax prices). Paid with
 * and Tax deductible start from `inherit` (the original's staged values).
 */
export const newSplitItem = (
  inherit: { paidWith?: string; taxDeductible?: boolean } = {},
): SplitDraftItem => ({
  key: `item-${keyCounter++}`,
  itemNote: "",
  expression: "",
  preTax: true,
  envelope: "",
  paidWith: inherit.paidWith ?? "",
  taxDeductible: inherit.taxDeductible ?? false,
  note: "",
});

export interface EvaluatedSplitItem {
  /** Tax-applied amount, or null when the expression is blank or invalid. */
  value: number | null;
  /** Inline message for an invalid amount; null when it is valid or still blank. */
  error: string | null;
  noteUnsafe: boolean;
}

const roundCents = (n: number) => Math.round(n * 100) / 100;

/** The amount a child cell will hold: the expression, times TaxMultiplier when Pre-Tax is on. */
export function evaluateSplitItem(
  item: Pick<SplitDraftItem, "expression" | "preTax" | "itemNote">,
  taxMultiplier: number,
): EvaluatedSplitItem {
  const amount = parseAmountInput(item.expression);
  const base = amount.ok ? amount.value : null;
  return {
    value: base === null ? null : item.preTax ? base * taxMultiplier : base,
    error: !amount.ok && amount.reason === "invalid" ? AMOUNT_MESSAGE : null,
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

/** The request items; each amount as `parseAmountInput` sends it (a plain number as typed). */
export function toSplitInputs(items: SplitDraftItem[]): SplitItemInput[] {
  return items.map((i) => {
    const amount = parseAmountInput(i.expression);
    return {
      itemNote: i.itemNote.trim(),
      expression: amount.ok ? amount.expression : i.expression.trim(),
      preTax: i.preTax,
      envelope: i.envelope,
      ...(i.paidWith ? { paidWith: i.paidWith } : {}),
      ...(i.taxDeductible ? { taxDeductible: "Yes" as const } : {}),
      ...(i.note.trim() ? { note: i.note } : {}),
    };
  });
}
