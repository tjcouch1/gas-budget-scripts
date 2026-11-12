import { useMemo, useState } from "react";
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

export function ExpensesView({
  engine,
  sheetName,
  state,
  rows,
}: {
  engine: SyncEngine;
  sheetName: string;
  state: PayPeriodStateDto;
  rows: ViewRow[];
}) {
  const [filter, setFilter] = useState<FilterId>("uncategorized");
  const [query, setQuery] = useState("");
  const [openKey, setOpenKey] = useState<string | null>(null);

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
  const options = useMemo(
    () =>
      state.envelopeOptions
        ? orderEnvelopeOptions(
            state.envelopeOptions,
            rows.filter((r) => !r.synthetic),
          )
        : null,
    [state.envelopeOptions, rows],
  );

  const openRow = rows.find((r) => r.key === openKey && !r.synthetic) ?? null;
  const serverRow =
    openRow && state.transactions.find((t) => t?.rowIndex === openRow.rowIndex);

  return (
    <div>
      <div className="sticky top-0 z-10 space-y-2 bg-slate-50 pb-2 pt-1 dark:bg-slate-950">
        <input
          className={inputClass}
          type="search"
          aria-label="Search"
          placeholder="Search name, envelope, paid with, date"
          value={query}
          onChange={(e) => setQuery(e.target.value)}
        />
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

      {groups.length === 0 ? (
        <p className="py-10 text-center text-slate-500">
          {rows.length === 0
            ? "No transactions in this pay period yet."
            : filter === "uncategorized" && !query
              ? "Everything is categorized."
              : "Nothing matches."}
        </p>
      ) : (
        <ul className="space-y-3">
          {groups.map((g) => (
            <li
              key={g.id}
              className="divide-y divide-slate-100 overflow-hidden rounded-xl shadow-sm ring-1 ring-slate-200 dark:divide-slate-800 dark:ring-slate-700"
            >
              {g.rows.map((r, i) => (
                <RowButton
                  key={r.key}
                  row={r}
                  nested={g.rows.length > 1 && i > 0}
                  dim={!(matchesFilter(r, filter) && matchesSearch(r, query))}
                  onOpen={(row) => setOpenKey(row.key)}
                />
              ))}
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
          taxMultiplier={state.taxMultiplier}
          envelopeOptions={options}
          envelopes={state.envelopes}
          paidWithOptions={state.paidWithOptions}
          onChange={(changes) =>
            engine.enqueueUpdate(sheetName, serverRow, changes)
          }
          onSplit={(splits) =>
            engine.enqueueSplit(sheetName, serverRow, splits)
          }
          onClose={() => setOpenKey(null)}
        />
      )}
    </div>
  );
}
