import { useEffect, useMemo, useState } from "react";
import type { PayPeriodStateDto } from "@budget/shared";
import { formatMoney, formatShortDate } from "../lib/format.ts";
import {
  FILTERS,
  groupRows,
  isSplitChild,
  isSplitParent,
  matchesFilter,
  matchesSearch,
  orderEnvelopeOptions,
  type FilterId,
  type ViewRow,
} from "../lib/rows.ts";
import type { SyncEngine } from "../sync/engine.ts";
import { buildViewRows } from "../sync/optimistic.ts";
import type { EngineSnapshot } from "../sync/types.ts";
import { TransactionSheet } from "./TransactionSheet.tsx";
import { inputClass } from "./ui.tsx";

function RowButton({
  row,
  nested,
  dim,
  onOpen,
}: {
  row: ViewRow;
  nested: boolean;
  dim: boolean;
  onOpen: (row: ViewRow) => void;
}) {
  const uncategorized = row.envelope.trim() === "";
  const tone = row.isError
    ? "border-l-4 border-l-red-500"
    : row.isNote
      ? "border-l-4 border-l-pink-400"
      : "";
  const content = (
    <>
      <span className="flex min-w-0 flex-1 flex-col text-left">
        <span className="truncate font-medium">{row.name || "(no name)"}</span>
        <span className="truncate text-xs text-slate-500 dark:text-slate-400">
          {formatShortDate(row.date)}
          {row.paidWith ? ` · ${row.paidWith}` : ""}
          {row.taxDeductible.trim() ? " · Tax deductible" : ""}
          {row.pending ? " · syncing…" : ""}
        </span>
      </span>
      <span className="flex shrink-0 flex-col items-end">
        <span className="font-medium tabular-nums">
          {row.costDisplay || formatMoney(row.cost)}
        </span>
        <span
          className={`max-w-32 truncate text-xs ${
            uncategorized ? "font-semibold text-amber-600" : "text-slate-500"
          }`}
        >
          {uncategorized ? "Uncategorized" : row.envelope}
        </span>
      </span>
    </>
  );
  const cls = `flex min-h-14 w-full items-center gap-3 bg-white px-3 py-2 dark:bg-slate-900 ${tone} ${
    nested ? "pl-8" : ""
  } ${dim ? "opacity-50" : ""}`;
  if (row.synthetic) {
    return <div className={cls}>{content}</div>;
  }
  return (
    <button
      type="button"
      className={`${cls} hover:bg-teal-50 dark:hover:bg-slate-800`}
      onClick={() => onOpen(row)}
    >
      {content}
    </button>
  );
}

type Scope = "current" | "all";

export function ExpensesView({
  engine,
  sheetName,
  state,
  rows: currentRows,
  snapshot,
}: {
  engine: SyncEngine;
  sheetName: string;
  state: PayPeriodStateDto;
  /** Rows of the selected pay period, with queued edits applied. */
  rows: ViewRow[];
  snapshot: Pick<EngineSnapshot, "all" | "loadingAll" | "queue">;
}) {
  const [filter, setFilter] = useState<FilterId>("uncategorized");
  const [query, setQuery] = useState("");
  const [openKey, setOpenKey] = useState<string | null>(null);
  const [scope, setScope] = useState<Scope>("current");
  const { all, loadingAll, queue } = snapshot;

  // All pay periods read every tab, so they load on demand only (cached rows show meanwhile).
  useEffect(() => {
    if (scope === "all") void engine.loadAll();
  }, [scope, engine]);

  const allRows = useMemo(
    () =>
      scope === "all" && all
        ? all.sheets.flatMap((sheet) =>
            buildViewRows(
              {
                sheetName: sheet.name,
                taxMultiplier: all.taxMultiplier,
                transactions: sheet.transactions,
              },
              queue,
              true,
            ),
          )
        : [],
    [scope, all, queue],
  );
  const rows = scope === "all" ? allRows : currentRows;

  const groups = useMemo(
    () => groupRows(rows, filter, query),
    [rows, filter, query],
  );
  const counts = useMemo(() => {
    const result = {} as Record<FilterId, number>;
    for (const f of FILTERS)
      result[f.id] = rows.filter(
        (r) => !r.synthetic && matchesFilter(r, f.id),
      ).length;
    return result;
  }, [rows]);
  const openRow = rows.find((r) => r.key === openKey && !r.synthetic) ?? null;
  // The tab the open row lives in: edits and splits go to it with its own fingerprint.
  const openSheetName = openRow?.sheetName ?? sheetName;
  const openAllSheet =
    scope === "all" && openRow
      ? all?.sheets.find((x) => x.name === openSheetName)
      : undefined;
  const serverRow =
    openRow &&
    (openAllSheet ? openAllSheet.transactions : state.transactions).find(
      (t) => t?.rowIndex === openRow.rowIndex,
    );
  const rawEnvelopeOptions = openAllSheet
    ? openAllSheet.envelopeOptions
    : state.envelopeOptions;
  const options = useMemo(
    () =>
      rawEnvelopeOptions
        ? orderEnvelopeOptions(
            rawEnvelopeOptions,
            rows.filter((r) => !r.synthetic && r.sheetName === openSheetName),
          )
        : null,
    [rawEnvelopeOptions, rows, openSheetName],
  );

  return (
    <div>
      <div className="sticky top-0 z-10 space-y-2 bg-slate-50 pb-2 pt-1 dark:bg-slate-950">
        <div className="flex gap-2">
          <select
            aria-label="Transactions shown"
            className="min-h-11 w-36 shrink-0 rounded-lg border border-slate-300 bg-white px-2 text-base dark:border-slate-600 dark:bg-slate-900"
            value={scope}
            onChange={(e) => setScope(e.target.value as Scope)}
          >
            <option value="current">This pay period</option>
            <option value="all">All pay periods</option>
          </select>
          <input
            className={`${inputClass} min-w-0 flex-1`}
            type="search"
            aria-label="Search"
            placeholder={
              scope === "all"
                ? "Search all pay periods"
                : "Search name, envelope, paid with, date"
            }
            value={query}
            onChange={(e) => setQuery(e.target.value)}
          />
        </div>
        <div
          className="flex gap-2 overflow-x-auto"
          role="tablist"
          aria-label="Filter"
        >
          {FILTERS.map((f) => (
            <button
              key={f.id}
              type="button"
              role="tab"
              aria-selected={filter === f.id}
              onClick={() => setFilter(f.id)}
              className={`min-h-11 shrink-0 rounded-full px-4 text-sm ring-1 ${
                filter === f.id
                  ? "bg-teal-700 text-white ring-teal-700"
                  : "bg-white ring-slate-300 dark:bg-slate-800 dark:ring-slate-600"
              }`}
            >
              {f.label} <span className="opacity-70">{counts[f.id]}</span>
            </button>
          ))}
        </div>
      </div>

      {scope === "all" && !all ? (
        <p role="status" className="py-10 text-center text-slate-500">
          {loadingAll
            ? "Loading all pay periods…"
            : "All pay periods could not be loaded yet. Check your connection and try again."}
        </p>
      ) : groups.length === 0 ? (
        <p className="py-10 text-center text-slate-500">
          {rows.length === 0
            ? "No transactions in this pay period yet."
            : filter === "uncategorized" && !query
              ? "Everything is categorized."
              : "Nothing matches."}
        </p>
      ) : (
        <ul className="space-y-3">
          {groups.map((g, gi) => (
            <li key={g.id} className="space-y-1">
              {scope === "all" &&
                g.rows[0].sheetName !== groups[gi - 1]?.rows[0].sheetName && (
                  <h3 className="px-1 pt-2 text-sm font-semibold text-slate-600 dark:text-slate-300">
                    {g.rows[0].sheetName}
                  </h3>
                )}
              <div className="divide-y divide-slate-100 overflow-hidden rounded-xl shadow-sm ring-1 ring-slate-200 dark:divide-slate-800 dark:ring-slate-700">
                {g.rows.map((r, i) => (
                  <RowButton
                    key={r.key}
                    row={r}
                    nested={g.rows.length > 1 && i > 0}
                    dim={!(matchesFilter(r, filter) && matchesSearch(r, query))}
                    onOpen={(row) => setOpenKey(row.key)}
                  />
                ))}
              </div>
            </li>
          ))}
        </ul>
      )}

      {openRow && serverRow && (
        <TransactionSheet
          key={openRow.key}
          row={openRow}
          isSplitParent={isSplitParent(openRow, rows)}
          isSplitChild={isSplitChild(openRow, rows)}
          taxMultiplier={
            all && openAllSheet ? all.taxMultiplier : state.taxMultiplier
          }
          envelopeOptions={options}
          envelopes={openSheetName === state.sheetName ? state.envelopes : null}
          paidWithOptions={
            openAllSheet ? openAllSheet.paidWithOptions : state.paidWithOptions
          }
          onChange={(changes) =>
            engine.enqueueUpdate(openSheetName, serverRow, changes)
          }
          onSplit={(splits) =>
            engine.enqueueSplit(openSheetName, serverRow, splits)
          }
          onClose={() => setOpenKey(null)}
        />
      )}
    </div>
  );
}
