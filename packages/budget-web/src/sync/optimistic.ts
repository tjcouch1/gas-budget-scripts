import {
  buildSplitChildName,
  buildSplitParentName,
  evaluateMathExpression,
  formatSplitCostFormula,
  type TransactionChanges,
  type TransactionRowDto,
} from "@budget/shared";
import { formatMoney } from "../lib/format.ts";
import type { ViewRow } from "../lib/rows.ts";
import type { QueuedMutation } from "./types.ts";

/** Applies an update's changes to a row the way the sheet will show them. */
export function applyChangesToRow(
  row: TransactionRowDto,
  changes: TransactionChanges,
  taxMultiplier: number,
): TransactionRowDto {
  const next = { ...row };
  if (changes.name !== undefined) next.name = changes.name;
  if (changes.envelope !== undefined) next.envelope = changes.envelope;
  if (changes.paidWith !== undefined) next.paidWith = changes.paidWith;
  if (changes.taxDeductible !== undefined)
    next.taxDeductible = changes.taxDeductible;
  if (changes.costExpression !== undefined) {
    try {
      const formula = formatSplitCostFormula(
        changes.costExpression,
        changes.costPreTax === true,
        taxMultiplier,
      );
      const value = evaluateMathExpression(
        typeof formula === "number" ? String(formula) : formula.slice(1),
      );
      next.cost = value;
      next.costFormula = typeof formula === "string" ? formula : null;
      next.costDisplay = formatMoney(value);
    } catch {
      // Invalid expressions never reach the queue; ignore defensively.
    }
  }
  return next;
}

/** The parent row after a split, as shown optimistically. */
export function applySplitToParent(
  row: TransactionRowDto,
  splitTotal: number,
  parent: { itemNote?: string; envelope?: string } = {},
): TransactionRowDto {
  const next = {
    ...row,
    name:
      parent.itemNote === undefined
        ? buildSplitParentName(row.name)
        : buildSplitChildName(row.name, parent.itemNote),
  };
  if (parent.envelope !== undefined) next.envelope = parent.envelope;
  if (next.cost !== null) {
    next.cost = Math.round((next.cost - splitTotal) * 100) / 100;
    next.costDisplay = formatMoney(next.cost);
  }
  return next;
}

/** The row as it would look if `mutation` were applied ("Discarded Local Change"). */
export function proposedRow(
  mutation: QueuedMutation,
  base: TransactionRowDto,
  taxMultiplier: number,
): TransactionRowDto {
  if (mutation.action === "updateTransaction") {
    return applyChangesToRow(base, mutation.payload.changes, taxMultiplier);
  }
  return applySplitToParent(base, 0, {
    itemNote: mutation.payload.parentItemNote,
    envelope: mutation.payload.parentEnvelope,
  });
}

/** What `buildViewRows` needs of a sheet: a pay-period state or one tab of the all-sheets view. */
export interface RowSource {
  sheetName: string;
  taxMultiplier: number;
  transactions: (TransactionRowDto | null)[];
}

/**
 * Server rows with queued mutations applied optimistically. Queued splits mark
 * the parent and add synthetic placeholder children (without shifting real
 * row indexes, which later queued edits still refer to).
 */
export function buildViewRows(
  state: RowSource | undefined,
  queue: QueuedMutation[],
  /** True when rows of several sheets are shown together: keys and group ids get the sheet name. */
  scoped = false,
): ViewRow[] {
  if (!state) return [];
  const mutationsByRow = new Map<number, QueuedMutation[]>();
  for (const m of queue) {
    if (m.sheetName !== state.sheetName) continue;
    const list = mutationsByRow.get(m.rowIndex) ?? [];
    list.push(m);
    mutationsByRow.set(m.rowIndex, list);
  }
  const rows: ViewRow[] = [];
  const prefix = scoped ? `${state.sheetName}#` : "";
  for (const original of state.transactions) {
    if (!original) continue;
    const t: TransactionRowDto = scoped
      ? {
          ...original,
          splitGroupId:
            original.splitGroupId === null
              ? null
              : `${prefix}${original.splitGroupId}`,
        }
      : original;
    const pendingGroupId = t.splitGroupId ?? `${prefix}pending:${t.rowIndex}`;
    let row: ViewRow = {
      ...t,
      key: `${prefix}row:${t.rowIndex}`,
      sheetName: state.sheetName,
    };
    const synthetic: ViewRow[] = [];
    for (const m of mutationsByRow.get(t.rowIndex) ?? []) {
      row.pending = true;
      if (m.action === "updateTransaction") {
        row = {
          ...applyChangesToRow(row, m.payload.changes, state.taxMultiplier),
          key: row.key,
          pending: true,
          splitPending: row.splitPending,
        };
      } else {
        let total = 0;
        m.payload.splits.forEach((s, i) => {
          const expr = s.expression.trim().replace(/^=/, "");
          const base = evaluateMathExpression(expr) ?? 0;
          const value = s.preTax ? base * state.taxMultiplier : base;
          total += value;
          synthetic.push({
            ...t,
            key: `${prefix}split:${m.id}:${i}`,
            sheetName: state.sheetName,
            rowIndex: -1 - i,
            name: buildSplitChildName(t.name, s.itemNote),
            cost: value,
            costFormula: null,
            costDisplay: formatMoney(value),
            envelope: s.envelope,
            paidWith: s.paidWith ?? t.paidWith,
            taxDeductible: "",
            note: "",
            isError: false,
            isNote: false,
            splitGroupId: pendingGroupId,
            pending: true,
            synthetic: true,
          });
        });
        row = {
          ...applySplitToParent(row, total, {
            itemNote: m.payload.parentItemNote,
            envelope: m.payload.parentEnvelope,
          }),
          key: row.key,
          pending: true,
          splitPending: true,
          splitGroupId: pendingGroupId,
        };
      }
    }
    rows.push(row, ...synthetic);
  }
  return rows;
}
