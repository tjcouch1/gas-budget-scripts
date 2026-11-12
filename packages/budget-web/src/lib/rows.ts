import type {
  EnvelopeRowDto,
  PayPeriodStateDto,
  TransactionRowDto,
} from "@budget/shared";

/** A row as displayed: a server row, possibly with queued (optimistic) edits applied. */
export interface ViewRow extends TransactionRowDto {
  /** True when a queued mutation affects this row. */
  pending?: boolean;
  /** True for placeholder children of a queued split (no real row yet). */
  synthetic?: boolean;
  /** True when this row has a queued split. */
  splitPending?: boolean;
  /** The tab the row belongs to. Set by `buildViewRows`. */
  sheetName?: string;
  /** Key for React lists. */
  key: string;
}

export type FilterId = "uncategorized" | "all" | "issues";

export const FILTERS: { id: FilterId; label: string }[] = [
  { id: "uncategorized", label: "Uncategorized" },
  { id: "all", label: "All" },
  { id: "issues", label: "Errors/Notes" },
];

export function matchesFilter(row: ViewRow, filter: FilterId): boolean {
  switch (filter) {
    case "all":
      return true;
    case "uncategorized":
      return row.envelope.trim() === "";
    case "issues":
      return row.isError || row.isNote;
  }
}

export function matchesSearch(row: ViewRow, query: string): boolean {
  const q = query.trim().toLowerCase();
  if (!q) return true;
  return [row.name, row.envelope, row.paidWith, row.date, row.costDisplay].some(
    (field) => field.toLowerCase().includes(q),
  );
}

export interface RowGroup {
  id: string;
  rows: ViewRow[];
}

/**
 * Filters rows and nests split groups. A group is kept whole (parent plus
 * children, in order) when any member matches, so the split context is never lost.
 */
export function groupRows(
  rows: ViewRow[],
  filter: FilterId,
  query: string,
): RowGroup[] {
  const groups: RowGroup[] = [];
  for (const row of rows) {
    const last = groups[groups.length - 1];
    if (row.splitGroupId !== null && last?.id === `g:${row.splitGroupId}`) {
      last.rows.push(row);
    } else {
      groups.push({
        id:
          row.splitGroupId !== null ? `g:${row.splitGroupId}` : `r:${row.key}`,
        rows: [row],
      });
    }
  }
  return groups.filter((g) =>
    g.rows.some((r) => matchesFilter(r, filter) && matchesSearch(r, query)),
  );
}

/** Non-null rows of a state, in sheet order. */
export function stateRows(state: PayPeriodStateDto | undefined): ViewRow[] {
  if (!state) return [];
  return state.transactions
    .filter((t): t is TransactionRowDto => t !== null)
    .map((t) => ({ ...t, key: `row:${t.rowIndex}` }));
}

/** A row that may be split: not a child in the middle of a group. */
export function isSplitChild(row: ViewRow, all: ViewRow[]): boolean {
  if (row.splitGroupId === null) return false;
  const first = all.find((r) => r.splitGroupId === row.splitGroupId);
  return first !== undefined && first.rowIndex !== row.rowIndex;
}

export function isSplitParent(row: ViewRow, all: ViewRow[]): boolean {
  if (row.splitGroupId === null) return false;
  const first = all.find((r) => r.splitGroupId === row.splitGroupId);
  return (
    first?.rowIndex === row.rowIndex &&
    all.some(
      (r) => r.splitGroupId === row.splitGroupId && r.rowIndex !== row.rowIndex,
    )
  );
}

/**
 * Envelope choices ordered by how often they are used in the current sheet
 * (most used first), then sheet order, so the likely choice is one tap away.
 */
export function orderEnvelopeOptions(
  options: string[],
  rows: TransactionRowDto[],
): string[] {
  const counts = new Map<string, number>();
  for (const r of rows) {
    if (r.envelope) counts.set(r.envelope, (counts.get(r.envelope) ?? 0) + 1);
  }
  return options
    .map((name, i) => ({ name, i, n: counts.get(name) ?? 0 }))
    .sort((a, b) => b.n - a.n || a.i - b.i)
    .map((o) => o.name);
}

/** `Remain` is column 4 of the envelope table, i.e. index 3 of `display`/`values`. */
export const ENVELOPE_COLUMNS = [
  "Available Cumul",
  "Size",
  "Spent",
  "Remain",
  "Remain Cumul",
  "Size YTD",
  "Spent YTD",
  "Spent Avg YTD",
  "Size Projected",
  "Spent Projected",
] as const;
export const CORE_COLUMN_INDEXES = [1, 2, 3, 4] as const;
export const REMAIN_INDEX = 3;

export function envelopeRemain(
  envelopes: EnvelopeRowDto[] | null,
  name: string,
): { display: string; value: number | null } | null {
  const row = envelopes?.find((e) => e.name === name);
  if (!row) return null;
  return {
    display: row.display[REMAIN_INDEX] ?? "",
    value: row.values[REMAIN_INDEX] ?? null,
  };
}
