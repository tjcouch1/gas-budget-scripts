import { useEffect, useMemo, useRef, useState } from "react";
import type { PayPeriodStateDto } from "@budget/shared";
import { formatMoney, formatShortDate } from "../lib/format.ts";
import {
  FILTERS,
  groupRows,
  matchesFilter,
  matchesSearch,
  orderEnvelopeOptions,
  type FilterId,
  type ViewRow,
} from "../lib/rows.ts";
import type { SyncEngine } from "../sync/engine.ts";
import {
  getRecentEnvelopes,
  recordRecentEnvelope,
} from "../lib/preferences.ts";
import { buildViewRows } from "../sync/optimistic.ts";
import type { EngineSnapshot } from "../sync/types.ts";
import { NewExpenseSheet } from "./NewExpenseSheet.tsx";
import { TransactionSheet } from "./TransactionSheet.tsx";
import { Button, inputClass } from "./ui.tsx";

type SaveStatus = "conflict" | "offline" | "saving" | "saved";

const STATUS_LABEL: Record<SaveStatus, string> = {
  conflict: "Conflict",
  offline: "Queued – offline",
  saving: "Saving…",
  saved: "Saved ✓",
};

const STATUS_COLOR: Record<SaveStatus, string> = {
  conflict: "text-red-600",
  offline: "text-amber-600",
  saving: "text-sky-600",
  saved: "text-emerald-600",
};

/** Where a row's changes stand: not yet at the sheet, in conflict, or just confirmed. */
function rowStatus(
  row: ViewRow,
  s: Pick<EngineSnapshot, "conflicts" | "offline" | "recentlySaved">,
): SaveStatus | null {
  if (row.synthetic)
    return row.pending ? (s.offline ? "offline" : "saving") : null;
  const same = (m: { sheetName: string; rowIndex: number }) =>
    m.sheetName === row.sheetName && m.rowIndex === row.rowIndex;
  if (s.conflicts.some((c) => same(c.mutation))) return "conflict";
  if (row.pending) return s.offline ? "offline" : "saving";
  if (s.recentlySaved.some(same)) return "saved";
  return null;
}

function RowButton({
  row,
  dim,
  status,
  highlighted,
  envelopeChoices,
  onCategorize,
  onOpen,
  onShowInAll,
}: {
  row: ViewRow;
  dim: boolean;
  status: SaveStatus | null;
  highlighted: boolean;
  /** Envelopes an uncategorized row can be put in from the list; null hides the dropdown. */
  envelopeChoices: string[] | null;
  onCategorize: (row: ViewRow, envelope: string) => void;
  onOpen: (row: ViewRow) => void;
  /** Present in the Uncategorized and Errors/Notes filters: jump to the row in All. */
  onShowInAll: ((row: ViewRow) => void) | null;
}) {
  const uncategorized = row.envelope.trim() === "";
  const tone = row.isError
    ? "border-l-4 border-l-red-500"
    : row.isNote
      ? "border-l-4 border-l-pink-400"
      : row.note
        ? // A Name-cell note without a special background: a bar in the text color
          "border-l-4 border-l-slate-700 dark:border-l-slate-200"
        : "";
  const nameBlock = (
    <span className="flex min-w-0 flex-1 flex-col text-left">
      <span className="truncate font-medium">{row.name || "(no name)"}</span>
      <span className="truncate text-xs text-slate-500 dark:text-slate-400">
        {formatShortDate(row.date)}
        {row.paidWith ? ` · ${row.paidWith}` : ""}
        {row.taxDeductible.trim() ? " · Tax deductible" : ""}
      </span>
      {status && (
        <span
          data-testid="row-status"
          data-row-status={status}
          className={`truncate text-xs font-medium ${STATUS_COLOR[status]}`}
        >
          {STATUS_LABEL[status]}
        </span>
      )}
    </span>
  );
  const costBlock = (
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
  );
  const cls = `flex min-h-14 w-full items-center bg-white dark:bg-slate-900 ${tone} ${
    dim ? "opacity-50" : ""
  } ${highlighted ? "ring-2 ring-inset ring-teal-500" : ""}`;
  if (row.synthetic) {
    return (
      <div className={`${cls} gap-3 px-3 py-2`} data-row-key={row.key}>
        {nameBlock}
        {costBlock}
      </div>
    );
  }
  const open =
    "flex items-center gap-3 self-stretch py-2 hover:bg-teal-50 dark:hover:bg-slate-800";
  return (
    <div className={cls} data-row-key={row.key}>
      {uncategorized && envelopeChoices ? (
        <>
          <button
            type="button"
            className={`${open} min-w-0 flex-1 pl-3`}
            onClick={() => onOpen(row)}
          >
            {nameBlock}
          </button>
          <select
            aria-label={`Envelope for ${row.name || "(no name)"}`}
            className="mx-2 max-w-32 shrink-0 rounded-lg border border-slate-300 bg-white px-1 py-1 text-xs dark:border-slate-600 dark:bg-slate-800"
            value=""
            onChange={(e) => {
              if (e.target.value) onCategorize(row, e.target.value);
            }}
          >
            <option value="">Envelope…</option>
            {envelopeChoices.map((name) => (
              <option key={name} value={name}>
                {name}
              </option>
            ))}
          </select>
          {/* A tap target, not a second button: the name button opens the row for keyboards */}
          <div
            className={`${open} cursor-pointer pr-3`}
            onClick={() => onOpen(row)}
          >
            {costBlock}
          </div>
        </>
      ) : (
        <button
          type="button"
          className={`${open} min-w-0 flex-1 px-3`}
          onClick={() => onOpen(row)}
        >
          {nameBlock}
          {costBlock}
        </button>
      )}
      {onShowInAll && (
        <button
          type="button"
          aria-label="Show in All"
          title="Show in All"
          className="mr-2 flex size-9 shrink-0 items-center justify-center rounded-full text-slate-500 ring-1 ring-slate-300 hover:bg-teal-50 dark:text-slate-300 dark:ring-slate-600 dark:hover:bg-slate-800"
          onClick={() => onShowInAll(row)}
        >
          <svg
            aria-hidden="true"
            viewBox="0 0 24 24"
            className="size-4"
            fill="none"
            stroke="currentColor"
            strokeWidth="2"
            strokeLinecap="round"
            strokeLinejoin="round"
          >
            <path d="M5 12h14M13 6l6 6-6 6" />
          </svg>
        </button>
      )}
    </div>
  );
}

export function ExpensesView({
  engine,
  sheetName,
  state,
  rows: currentRows,
  showAll,
  allSheetNames = null,
  snapshot,
}: {
  engine: SyncEngine;
  sheetName: string;
  /** The selected pay period; undefined while it is loading. */
  state: PayPeriodStateDto | undefined;
  /** Rows of the selected pay period, with queued edits applied. */
  rows: ViewRow[];
  /** True when the pay-period dropdown is on "All". */
  showAll: boolean;
  /** Pay periods the multi-period view shows (a date range); null for every pay period. */
  allSheetNames?: string[] | null;
  snapshot: Pick<
    EngineSnapshot,
    | "all"
    | "loadingAll"
    | "loadAllFailed"
    | "queue"
    | "conflicts"
    | "offline"
    | "recentlySaved"
  >;
}) {
  const [filter, setFilter] = useState<FilterId>("uncategorized");
  const [query, setQuery] = useState("");
  const [openKey, setOpenKey] = useState<string | null>(null);
  const [adding, setAdding] = useState(false);
  const [jumpKey, setJumpKey] = useState<string | null>(null);
  const [highlightKey, setHighlightKey] = useState<string | null>(null);
  const listRef = useRef<HTMLUListElement>(null);
  const { all, loadingAll, loadAllFailed, queue } = snapshot;

  // Multi-period views read several tabs, so they load on demand only (cached rows show
  // meanwhile), in small batches whose rows appear as they arrive.
  const sheetNamesKey = allSheetNames ? JSON.stringify(allSheetNames) : "";
  useEffect(() => {
    if (showAll)
      void engine.loadAll(
        sheetNamesKey ? (JSON.parse(sheetNamesKey) as string[]) : undefined,
      );
  }, [showAll, sheetNamesKey, engine]);

  const allRows = useMemo(
    () =>
      showAll && all
        ? all.sheets
            .filter(
              (sheet) => !allSheetNames || allSheetNames.includes(sheet.name),
            )
            .flatMap((sheet) =>
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
    [showAll, all, queue, allSheetNames],
  );
  const rows = showAll ? allRows : currentRows;

  // Rows whose save is in flight or just confirmed stay listed so the user sees the result.
  const { conflicts, offline, recentlySaved } = snapshot;
  const statuses = useMemo(() => {
    const map = new Map<string, SaveStatus>();
    for (const r of rows) {
      const status = rowStatus(r, { conflicts, offline, recentlySaved });
      if (status) map.set(r.key, status);
    }
    return map;
  }, [rows, conflicts, offline, recentlySaved]);
  const groups = useMemo(
    () => groupRows(rows, filter, query, (r) => statuses.has(r.key)),
    [rows, filter, query, statuses],
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
    showAll && openRow
      ? all?.sheets.find((x) => x.name === openSheetName)
      : undefined;
  // The whole split group opens together, in sheet order (queued split placeholders are not editable).
  const openGroup = !openRow
    ? []
    : openRow.splitGroupId === null
      ? [openRow]
      : rows.filter(
          (r) =>
            !r.synthetic &&
            r.splitGroupId === openRow.splitGroupId &&
            (r.sheetName ?? sheetName) === openSheetName,
        );
  const serverRows = openAllSheet
    ? openAllSheet.transactions
    : (state?.transactions ?? []);
  /** The row as last pulled: edits and splits carry its fingerprint. */
  const serverRowOf = (row: ViewRow) =>
    serverRows.find((t) => t?.rowIndex === row.rowIndex) ?? null;
  const canOpen =
    openGroup.length > 0 && openGroup.every((r) => serverRowOf(r) !== null);
  const rawEnvelopeOptions = openAllSheet
    ? openAllSheet.envelopeOptions
    : (state?.envelopeOptions ?? null);
  const options = useMemo(
    () =>
      rawEnvelopeOptions
        ? orderEnvelopeOptions(
            rawEnvelopeOptions,
            rows.filter((r) => !r.synthetic && r.sheetName === openSheetName),
            getRecentEnvelopes(),
          )
        : null,
    // The recent order is re-read whenever a sheet opens or closes
    [rawEnvelopeOptions, rows, openSheetName, openKey, adding],
  );

  /** The All-view tab a listed row lives in (null outside All). */
  const sheetOf = (row: ViewRow) =>
    showAll && all
      ? (all.sheets.find((x) => x.name === row.sheetName) ?? null)
      : null;
  const inlineChoices = (row: ViewRow): string[] | null =>
    sheetOf(row)?.envelopeOptions ?? state?.envelopeOptions ?? null;
  const categorize = (row: ViewRow, envelope: string) => {
    const sheet = sheetOf(row);
    const server = (sheet?.transactions ?? state?.transactions ?? []).find(
      (t) => t?.rowIndex === row.rowIndex,
    );
    if (!server) return;
    recordRecentEnvelope(envelope);
    engine.enqueueUpdate(sheet?.name ?? state?.sheetName ?? sheetName, server, {
      envelope,
    });
  };

  // "Show in All": switch the filter, then scroll to the row once the All list is rendered.
  useEffect(() => {
    if (!jumpKey || filter !== "all") return;
    const el = [
      ...(listRef.current?.querySelectorAll<HTMLElement>("[data-row-key]") ??
        []),
    ].find((x) => x.dataset.rowKey === jumpKey);
    el?.scrollIntoView?.({ block: "center" });
    setHighlightKey(jumpKey);
    setJumpKey(null);
    const timer = setTimeout(() => setHighlightKey(null), 2000);
    return () => clearTimeout(timer);
  }, [jumpKey, filter, groups]);

  return (
    <div>
      <div className="sticky top-0 z-10 space-y-2 bg-slate-50 pb-2 pt-1 dark:bg-slate-950">
        <div className="flex gap-2">
          <input
            className={`${inputClass} min-w-0 flex-1`}
            type="search"
            aria-label="Search"
            placeholder={
              showAll
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

      {showAll && all && (loadingAll || loadAllFailed) && (
        <p role="status" className="pb-2 text-center text-sm text-slate-500">
          {loadingAll
            ? "Loading more pay periods…"
            : "Some pay periods could not be loaded yet. Tap Sync to try again."}
        </p>
      )}

      {showAll && !all ? (
        <p role="status" className="py-10 text-center text-slate-500">
          {loadingAll
            ? "Loading pay periods…"
            : "These pay periods could not be loaded yet. Check your connection and tap Sync to try again."}
        </p>
      ) : !showAll && !state ? (
        <p role="status" className="py-10 text-center text-slate-500">
          Loading this pay period…
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
        <ul ref={listRef} className="space-y-3">
          {groups.map((g, gi) => (
            <li key={g.id} className="space-y-1">
              {showAll &&
                g.rows[0].sheetName !== groups[gi - 1]?.rows[0].sheetName && (
                  <h3 className="px-1 pt-2 text-sm font-semibold text-slate-600 dark:text-slate-300">
                    {g.rows[0].sheetName}
                  </h3>
                )}
              <div className="divide-y divide-slate-100 overflow-hidden rounded-xl shadow-sm ring-1 ring-slate-200 dark:divide-slate-800 dark:ring-slate-700">
                {g.rows.map((r) => (
                  <RowButton
                    key={r.key}
                    row={r}
                    dim={
                      !statuses.has(r.key) &&
                      !(matchesFilter(r, filter) && matchesSearch(r, query))
                    }
                    status={statuses.get(r.key) ?? null}
                    highlighted={highlightKey === r.key}
                    envelopeChoices={inlineChoices(r)}
                    onCategorize={categorize}
                    onShowInAll={
                      filter === "all"
                        ? null
                        : (row) => {
                            setQuery("");
                            setFilter("all");
                            setJumpKey(row.key);
                          }
                    }
                    onOpen={(row) => setOpenKey(row.key)}
                  />
                ))}
              </div>
            </li>
          ))}
        </ul>
      )}

      {/* At the end of the list (not sticky): new rows go after the last row of the selected tab. */}
      {!showAll && state && (
        <Button
          variant="secondary"
          className="mt-4 w-full"
          onClick={() => setAdding(true)}
        >
          New expense
        </Button>
      )}

      {adding && state && !showAll && (
        <NewExpenseSheet
          period={state.sheets.find((x) => x.name === state.sheetName) ?? null}
          envelopeOptions={options}
          envelopeLayoutOptions={state.envelopeOptions}
          envelopes={state.envelopes}
          paidWithOptions={state.paidWithOptions}
          onSave={(fields) => engine.enqueueAdd(state.sheetName, fields)}
          onClose={() => setAdding(false)}
        />
      )}

      {openRow && canOpen && (
        <TransactionSheet
          key={openRow.key}
          rows={openGroup}
          focusRowIndex={openRow.rowIndex}
          taxMultiplier={
            all && openAllSheet
              ? all.taxMultiplier
              : (state?.taxMultiplier ?? 1)
          }
          envelopeOptions={options}
          envelopeLayoutOptions={rawEnvelopeOptions}
          envelopes={
            state && openSheetName === state.sheetName ? state.envelopes : null
          }
          paidWithOptions={
            openAllSheet
              ? openAllSheet.paidWithOptions
              : (state?.paidWithOptions ?? null)
          }
          onUpdate={(row, changes) => {
            const server = serverRowOf(row);
            if (server) engine.enqueueUpdate(openSheetName, server, changes);
          }}
          onSplit={(row, splits, parent) => {
            const server = serverRowOf(row);
            if (server)
              engine.enqueueSplit(openSheetName, server, splits, parent);
          }}
          onDelete={(group) => {
            const servers = group.map(serverRowOf);
            if (servers.every((r) => r !== null))
              engine.enqueueDelete(openSheetName, servers);
          }}
          onClose={() => setOpenKey(null)}
        />
      )}
    </div>
  );
}
