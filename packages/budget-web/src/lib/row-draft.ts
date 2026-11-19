import {
  evaluateMathExpression,
  isUnsafeFreeText,
  normalizePlainAmount,
  type TransactionChanges,
  type TransactionRowDto,
} from "@budget/shared";
import { joinNameParts, splitNameParts } from "./name-parts.ts";

/** What a card's name field edits: the whole name, only the item note after " | ", or nothing. */
export type NameMode = "name" | "itemNote" | "none";
/**
 * What a card's cost field edits: a plain cost (a plain number only, no tax toggle), a split
 * child's math expression with its "Add tax" toggle, or nothing (a split parent's remainder).
 */
export type CostMode = "plain" | "child" | "none";

export interface RowDraftOptions {
  nameMode: NameMode;
  costMode: CostMode;
  taxMultiplier: number;
}

/** The editable fields of an existing row, as staged in the drawer. */
export interface RowDraft {
  /** The whole name, or the item note (see `NameMode`). */
  name: string;
  /** Cost text as typed: math, optionally with a leading "=". */
  cost: string;
  preTax: boolean;
  envelope: string;
  paidWith: string;
  taxDeductible: boolean;
  /** The Name cell's note. */
  note: string;
}

const MULTIPLIER_FORMULA = /^=\s*(-?\d+(?:\.\d+)?)\s*\*\s*\((.*)\)\s*$/;

/**
 * The expression and "Add tax" state behind a split child's cost cell:
 * `=<TaxMultiplier>*(<expr>)` is `<expr>` with tax; any other formula is shown
 * without its "="; a number is shown as is (AGENTS.md section 3.3).
 */
export function parseChildCost(
  row: Pick<TransactionRowDto, "cost" | "costFormula">,
  taxMultiplier: number,
): { expression: string; preTax: boolean } {
  const formula = row.costFormula?.trim();
  if (formula) {
    const m = MULTIPLIER_FORMULA.exec(formula);
    if (m && Number(m[1]) === taxMultiplier)
      return { expression: m[2].trim(), preTax: true };
    return { expression: formula.replace(/^=/, "").trim(), preTax: false };
  }
  return {
    expression: row.cost === null ? "" : String(row.cost),
    preTax: false,
  };
}

export function initialRowDraft(
  row: TransactionRowDto,
  { nameMode, costMode, taxMultiplier }: RowDraftOptions,
): RowDraft {
  const child =
    costMode === "child" ? parseChildCost(row, taxMultiplier) : null;
  return {
    name: nameMode === "itemNote" ? splitNameParts(row.name).note : row.name,
    cost: child
      ? child.expression
      : (row.costFormula ?? (row.cost === null ? "" : String(row.cost))),
    preTax: child?.preTax ?? false,
    envelope: row.envelope,
    paidWith: row.paidWith,
    taxDeductible: row.taxDeductible.trim().toLowerCase() === "yes",
    note: row.note,
  };
}

export interface EvaluatedRowDraft {
  /** Only the fields that differ from the row, as one update. */
  changes: TransactionChanges;
  /** The row's amount for the split remainder: the edited value, or the row's current cost. */
  value: number | null;
  costInvalid: boolean;
  /** Why the cost is invalid, for an inline error; null when it is valid. */
  costError: string | null;
  nameUnsafe: boolean;
}

export const ENTER_VALUE_MESSAGE = "Enter a value";
export const PLAIN_NUMBER_MESSAGE = "Enter a number, e.g. 12.34";

/**
 * Checks a plain amount field: blank is "Enter a value", anything but an optional "-", digits and
 * at most one "." is invalid. `amount` is the number as typed, to send as the cost expression.
 */
export function checkPlainAmount(
  text: string,
): { amount: string; error: null } | { amount: null; error: string } {
  if (text.trim() === "") return { amount: null, error: ENTER_VALUE_MESSAGE };
  const amount = normalizePlainAmount(text);
  return amount === null
    ? { amount: null, error: PLAIN_NUMBER_MESSAGE }
    : { amount, error: null };
}

export function evaluateRowDraft(
  row: TransactionRowDto,
  initial: RowDraft,
  draft: RowDraft,
  { nameMode, costMode, taxMultiplier }: RowDraftOptions,
): EvaluatedRowDraft {
  const changes: TransactionChanges = {};
  if (nameMode !== "none" && draft.name !== initial.name)
    changes.name =
      nameMode === "itemNote"
        ? joinNameParts(splitNameParts(row.name).prefix, draft.name)
        : draft.name;

  const expr = draft.cost.trim().replace(/^=/, "").trim();
  let value = row.cost;
  let costError: string | null = null;
  if (costMode === "plain" && draft.cost.trim() !== initial.cost.trim()) {
    // A plain amount: exactly the number as typed, sent and stored as that number.
    const checked = checkPlainAmount(draft.cost);
    costError = checked.error;
    value = checked.amount === null ? null : Number(checked.amount);
    if (checked.amount !== null) changes.costExpression = checked.amount;
  }
  // A child's expression: clearing the box keeps the current cost.
  const costChanged =
    costMode === "child" &&
    expr !== "" &&
    (draft.cost.trim() !== initial.cost.trim() ||
      draft.preTax !== initial.preTax);
  if (costChanged) {
    const base = evaluateMathExpression(expr);
    if (base === null) costError = "Enter an amount, e.g. 4.23*2+3";
    value = base === null ? null : draft.preTax ? base * taxMultiplier : base;
    changes.costExpression = expr;
    if (draft.preTax) changes.costPreTax = true;
  }

  if (draft.envelope !== initial.envelope) changes.envelope = draft.envelope;
  if (draft.paidWith !== initial.paidWith) changes.paidWith = draft.paidWith;
  if (draft.taxDeductible !== initial.taxDeductible)
    changes.taxDeductible = draft.taxDeductible ? "Yes" : "";
  if (draft.note !== initial.note) changes.note = draft.note;

  return {
    changes,
    value,
    costInvalid: costError !== null,
    costError,
    nameUnsafe: nameMode !== "none" && isUnsafeFreeText(draft.name),
  };
}
