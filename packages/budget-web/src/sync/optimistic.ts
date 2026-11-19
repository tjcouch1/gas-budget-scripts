import {
  buildSplitChildName,
  buildSplitParentName,
  evaluateMathExpression,
  formatSplitCostFormula,
  getSplitGroupKey,
  normalizePlainAmount,
  splitParentFormulaParts,
  type TransactionChanges,
  type TransactionRowDto,
} from "@budget/shared";
import type { ViewRow } from "../lib/rows.ts";
import type {
  AddMutation,
  GroupMutation,
  QueuedMutation,
  RowMutation,
  UpdateMutation,
  SplitMutation,
} from "./types.ts";

const roundCents = (n: number) => Math.round(n * 100) / 100;
/** Drops floating point noise (e.g. 39.599999999999994 -> 39.6) without rounding to cents. */
const clean = (n: number) => Number(n.toPrecision(12));

/**
 * Applies an update's changes to a row the way the sheet will show them. Costs are kept as plain
 * numbers; `costDisplay` is cleared (the sheet's display text comes back with the next pull), so
 * the UI formats `cost` for display.
 */
export function applyChangesToRow(
  row: TransactionRowDto,
  changes: TransactionChanges,
  taxMultiplier: number,
): TransactionRowDto {
  const next = { ...row };
  if (changes.name !== undefined) next.name = changes.name;
  if (changes.note !== undefined) next.note = changes.note;
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
      next.costDisplay = "";
    } catch {
      // Invalid expressions never reach the queue; ignore defensively.
    }
  }
  if (changes.splitTotal !== undefined) {
    // The parent's remainder moves by the change of its total (`=<total>-<children>`).
    const parts = splitParentFormulaParts(row.costFormula);
    const total = normalizePlainAmount(changes.splitTotal);
    if (parts && parts.total !== null && total !== null && next.cost !== null) {
      next.cost = clean(next.cost + Number(total) - parts.total);
      next.costFormula = `=${total}${parts.refs}`;
      next.costDisplay = "";
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
    next.cost = roundCents(next.cost - splitTotal);
    next.costDisplay = "";
  }
  return next;
}

/** A split row's name with the group's new title before "|" (a parent without "|" gets " | "). */
function withTitle(name: string, title: string, isParent: boolean): string {
  const bar = name.indexOf("|");
  if (bar < 0) return isParent ? buildSplitParentName(title) : name;
  return `${title} |${name.slice(bar + 1)}`;
}

/**
 * Row `index` of a split group save applied to `row`: its own changes, the parent's new total,
 * and the new title. The parent's remainder does not follow its children here (see
 * `proposedGroupRows` and `applyGroupEdits`).
 */
function applyGroupRow(
  row: TransactionRowDto,
  m: GroupMutation,
  index: number,
  taxMultiplier: number,
): TransactionRowDto {
  const changes: TransactionChanges = {
    ...m.payload.rows[index]?.changes,
    ...(index === 0 && m.payload.splitTotal !== undefined
      ? { splitTotal: m.payload.splitTotal }
      : {}),
  };
  const next = applyChangesToRow(row, changes, taxMultiplier);
  if (m.payload.title !== undefined)
    next.name = withTitle(next.name, m.payload.title, index === 0);
  return next;
}

/** The value of each new split item, with tax where asked. */
function newItemValues(m: GroupMutation, taxMultiplier: number): number[] {
  return (m.payload.newItems ?? []).map((s) => {
    const base = evaluateMathExpression(s.expression.trim().replace(/^=/, ""));
    return (base ?? 0) * (s.preTax ? taxMultiplier : 1);
  });
}

/**
 * Every row of a split group save as it would look on the sheet ("Discarded Local Change"), in
 * the order of `payload.rows`: the parent's remainder moves by its children's cost changes and
 * the new items, like its `=<total>-R[1]C[0]-...` formula.
 */
export function proposedGroupRows(
  m: GroupMutation,
  taxMultiplier: number,
): TransactionRowDto[] {
  const rows = m.payload.rows.map((r, i) =>
    applyGroupRow(r.baseRow, m, i, taxMultiplier),
  );
  let delta = 0;
  m.payload.rows.forEach((r, i) => {
    const before = r.baseRow.cost;
    const after = rows[i].cost;
    if (i > 0 && before !== null && after !== null) delta += after - before;
  });
  delta += newItemValues(m, taxMultiplier).reduce((sum, v) => sum + v, 0);
  const head = rows[0];
  if (head && head.cost !== null && delta !== 0)
    rows[0] = { ...head, cost: clean(head.cost - delta), costDisplay: "" };
  return rows;
}

/** The row as it would look if `mutation` were applied ("Discarded Local Change"). */
export function proposedRow(
  mutation: RowMutation,
  base: TransactionRowDto,
  taxMultiplier: number,
): TransactionRowDto {
  if (mutation.action === "updateTransaction") {
    return applyChangesToRow(base, mutation.payload.changes, taxMultiplier);
  }
  if (mutation.action === "updateSplitGroup") {
    const proposed = proposedGroupRows(mutation, taxMultiplier);
    return proposed.find((r) => r.rowIndex === base.rowIndex) ?? base;
  }
  // A delete has no proposed row (the conflict dialog shows it as deleted).
  if (mutation.action === "deleteTransaction") return base;
  return applySplitToParent(base, 0, {
    itemNote: mutation.payload.parentItemNote,
    envelope: mutation.payload.parentEnvelope,
  });
}

/** One row's share of a queued split group save, in `buildViewRows`. */
interface GroupRowEntry {
  action: "groupRow";
  m: GroupMutation;
  /** The row's position in `m.payload.rows`. */
  index: number;
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
 * row indexes, which later queued edits still refer to). Rows of a queued
 * delete are hidden.
 */
export function buildViewRows(
  state: RowSource | undefined,
  queue: QueuedMutation[],
  /** True when rows of several sheets are shown together: keys and group ids get the sheet name. */
  scoped = false,
): ViewRow[] {
  if (!state) return [];
  const mutationsByRow = new Map<
    number,
    (UpdateMutation | SplitMutation | GroupRowEntry)[]
  >();
  /** Split group saves with new items: the parent's rowIndex and the items' total. */
  const newItemTotals = new Map<number, number>();
  const deleting = new Set<number>();
  const adds: AddMutation[] = [];
  for (const m of queue) {
    if (m.sheetName !== state.sheetName) continue;
    if (m.action === "addTransaction") {
      adds.push(m);
      continue;
    }
    if (m.action === "deleteTransaction") {
      for (const r of m.payload.rows) deleting.add(r.rowIndex);
      continue;
    }
    if (m.action === "updateSplitGroup") {
      m.payload.rows.forEach((r, index) => {
        const list = mutationsByRow.get(r.rowIndex) ?? [];
        list.push({ action: "groupRow", m, index });
        mutationsByRow.set(r.rowIndex, list);
      });
      continue;
    }
    const list = mutationsByRow.get(m.rowIndex) ?? [];
    list.push(m);
    mutationsByRow.set(m.rowIndex, list);
  }
  const rows: ViewRow[] = [];
  const prefix = scoped ? `${state.sheetName}#` : "";
  for (const original of state.transactions) {
    if (!original || deleting.has(original.rowIndex)) continue;
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
      } else if (m.action === "groupRow") {
        row = {
          ...applyGroupRow(row, m.m, m.index, state.taxMultiplier),
          key: row.key,
          pending: true,
          splitPending: row.splitPending,
        };
        // New items go after the group's last row, named like the (renamed) parent.
        const items = m.m.payload.newItems ?? [];
        if (m.index === m.m.payload.rows.length - 1 && items.length > 0) {
          const head = m.m.payload.rows[0].rowIndex;
          const parentName =
            rows.find((r) => !r.synthetic && r.rowIndex === head)?.name ??
            row.name;
          const values = newItemValues(m.m, state.taxMultiplier);
          newItemTotals.set(
            head,
            (newItemTotals.get(head) ?? 0) +
              values.reduce((sum, v) => sum + v, 0),
          );
          items.forEach((s, i) =>
            synthetic.push({
              ...t,
              key: `${prefix}group:${m.m.id}:${i}`,
              sheetName: state.sheetName,
              rowIndex: -1 - i,
              name: buildSplitChildName(parentName, s.itemNote),
              cost: values[i],
              costFormula: null,
              costDisplay: "",
              envelope: s.envelope,
              paidWith: s.paidWith ?? t.paidWith,
              taxDeductible: s.taxDeductible ?? "",
              note: s.note ?? "",
              isError: false,
              isNote: false,
              splitGroupId: row.splitGroupId ?? pendingGroupId,
              pending: true,
              synthetic: true,
            }),
          );
        }
      } else if (m.action === "splitTransaction") {
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
            costDisplay: "",
            envelope: s.envelope,
            paidWith: s.paidWith ?? t.paidWith,
            taxDeductible: s.taxDeductible ?? "",
            note: s.note ?? "",
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
  // The parent's remainder drops by the queued new items (like `applySplitToParent`).
  for (const [head, total] of newItemTotals) {
    const index = rows.findIndex((r) => !r.synthetic && r.rowIndex === head);
    const parent = rows[index];
    if (parent && parent.cost !== null)
      rows[index] = {
        ...parent,
        cost: roundCents(parent.cost - total),
        costDisplay: "",
      };
  }
  applyGroupEdits(rows, state.transactions);
  // Queued new rows go at the end, where the server will write them.
  adds.forEach((m, i) => rows.push(pendingAddRow(m, i, state, prefix)));
  return rows;
}

/**
 * Keeps each split group consistent with queued edits of its rows, like the sheet's formulas do:
 * the parent's remainder (`=<total>-R[1]C[0]-...`) drops by however much its children's costs
 * went up, and a renamed parent's new prefix (before "|") is shown on every child. Changes `rows`
 * in place.
 */
function applyGroupEdits(
  rows: ViewRow[],
  server: (TransactionRowDto | null)[],
): void {
  const groups = new Map<string, ViewRow[]>();
  for (const row of rows) {
    if (row.splitGroupId === null) continue;
    const list = groups.get(row.splitGroupId) ?? [];
    list.push(row);
    groups.set(row.splitGroupId, list);
  }
  for (const list of groups.values()) {
    const head = list[0];
    const serverHead = server[head.rowIndex];
    if (list.length < 2 || head.synthetic || !serverHead) continue;
    const children = list.slice(1);
    let delta = 0;
    for (const child of children) {
      const before = child.synthetic ? null : server[child.rowIndex];
      if (before && before.cost !== null && child.cost !== null)
        delta += child.cost - before.cost;
    }
    const index = rows.indexOf(head);
    if (delta !== 0 && head.cost !== null)
      rows[index] = {
        ...rows[index],
        cost: clean(head.cost - delta),
        costDisplay: "",
        pending: true,
      };
    const prefix = getSplitGroupKey(head.name);
    if (prefix !== getSplitGroupKey(serverHead.name))
      for (const child of children) {
        const bar = child.name.indexOf("|");
        if (bar < 0 || getSplitGroupKey(child.name) === prefix) continue;
        rows[rows.indexOf(child)] = {
          ...child,
          name: `${prefix} |${child.name.slice(bar + 1)}`,
        };
      }
  }
}

/** The placeholder row of a queued `addTransaction`, as the sheet will show it. */
function pendingAddRow(
  m: AddMutation,
  i: number,
  state: RowSource,
  prefix: string,
): ViewRow {
  const { fields } = m.payload;
  const blank: TransactionRowDto = {
    rowIndex: -1000 - i,
    date: fields.date,
    name: "",
    cost: null,
    costFormula: null,
    costDisplay: "",
    envelope: "",
    paidWith: "",
    taxDeductible: "",
    note: "",
    isError: false,
    isNote: false,
    splitGroupId: null,
    fingerprint: "",
  };
  return {
    ...applyChangesToRow(
      blank,
      {
        name: fields.name,
        costExpression: fields.costExpression,
        envelope: fields.envelope ?? "",
        paidWith: fields.paidWith ?? "",
        taxDeductible: fields.taxDeductible ?? "",
        note: fields.note ?? "",
      },
      state.taxMultiplier,
    ),
    key: `${prefix}add:${m.id}`,
    sheetName: state.sheetName,
    pending: true,
    synthetic: true,
  };
}

const sameText = (a: string, b: string) => a.trimEnd() === b.trimEnd();

/**
 * True when two rows hold the same values in every field the user can edit (name, cost, envelope,
 * paid with, tax deductible, note). Costs compare by value when both are numbers, so `12.5` and
 * `=12.5`, or a formula in R1C1 and in A1 notation, match; otherwise by their text.
 */
export function sameEditableValues(
  a: TransactionRowDto,
  b: TransactionRowDto,
): boolean {
  const costText = (r: TransactionRowDto) =>
    (r.costFormula ?? r.costDisplay).replace(/\s+/g, "").toUpperCase();
  const sameCost =
    a.cost !== null && b.cost !== null
      ? Math.abs(a.cost - b.cost) <= 1e-9 * Math.max(1, Math.abs(a.cost))
      : a.cost === b.cost && costText(a) === costText(b);
  return (
    sameCost &&
    sameText(a.name, b.name) &&
    sameText(a.envelope, b.envelope) &&
    sameText(a.paidWith, b.paidWith) &&
    sameText(a.taxDeductible, b.taxDeductible) &&
    a.note === b.note
  );
}
