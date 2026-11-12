import { useEffect, useMemo, useState } from "react";
import { ConflictDialog } from "./components/ConflictDialog.tsx";
import { EnvelopesView } from "./components/EnvelopesView.tsx";
import { ExpensesView } from "./components/ExpensesView.tsx";
import { SettingsDialog } from "./components/SettingsDialog.tsx";
import { StatusPill } from "./components/StatusPill.tsx";
import { Button } from "./components/ui.tsx";
import { useEngineSnapshot, useHashTab, useOnlineEvents } from "./lib/hooks.ts";
import type { Services } from "./services.ts";
import { buildViewRows } from "./sync/optimistic.ts";

export function App({
  services,
  updateAvailable,
  onUpdate,
}: {
  services: Services;
  updateAvailable?: boolean;
  onUpdate?: () => void;
}) {
  const { engine, vault } = services;
  const snapshot = useEngineSnapshot(engine);
  const [tab, setTab] = useHashTab();
  const [settingsOpen, setSettingsOpen] = useState(false);
  const [configured, setConfigured] = useState<boolean | null>(null);
  useOnlineEvents(engine);

  useEffect(() => {
    let cancelled = false;
    void vault.hasCredentials().then(async (has) => {
      if (cancelled) return;
      setConfigured(has);
      if (!has) setSettingsOpen(true);
    });
    void engine.init();
    return () => {
      cancelled = true;
    };
  }, [engine, vault]);

  const state = snapshot.currentSheet
    ? snapshot.states[snapshot.currentSheet]
    : undefined;
  const rows = useMemo(
    () => buildViewRows(state, snapshot.queue),
    [state, snapshot.queue],
  );
  const conflict = snapshot.conflicts[0];

  return (
    <div className="mx-auto min-h-screen max-w-2xl px-4 pb-24 pt-3">
      <header className="mb-3 flex flex-wrap items-center gap-2">
        <h1 className="mr-auto text-xl font-bold text-teal-800 dark:text-teal-300">
          Fast Budget
        </h1>
        <StatusPill snapshot={snapshot} />
        <Button
          aria-label="Sync receipts"
          onClick={() => void engine.syncReceipts()}
          disabled={snapshot.syncingReceipts || snapshot.offline}
        >
          {snapshot.syncingReceipts ? "Importing…" : "Import receipts"}
        </Button>
        <Button
          aria-label="Settings"
          variant="ghost"
          onClick={() => setSettingsOpen(true)}
        >
          ⚙ Settings
        </Button>
      </header>

      {updateAvailable && (
        <div className="mb-3 flex items-center justify-between gap-2 rounded-xl bg-sky-100 p-3 text-sm text-sky-900">
          <span>A new version is available.</span>
          <Button onClick={onUpdate}>Update</Button>
        </div>
      )}

      {snapshot.notice && (
        <div
          role={snapshot.notice.kind === "error" ? "alert" : "status"}
          className={`mb-3 flex items-start justify-between gap-2 rounded-xl p-3 text-sm ${
            snapshot.notice.kind === "error"
              ? "bg-red-100 text-red-900"
              : "bg-emerald-100 text-emerald-900"
          }`}
        >
          <span>
            {snapshot.notice.message}
            {snapshot.notice.hint === "settings" && (
              <>
                {" "}
                <button
                  className="underline"
                  onClick={() => setSettingsOpen(true)}
                >
                  Open Settings
                </button>
              </>
            )}
          </span>
          <button aria-label="Dismiss" onClick={() => engine.dismissNotice()}>
            ×
          </button>
        </div>
      )}

      {configured === false && !settingsOpen && (
        <p className="py-10 text-center text-slate-500">
          Not connected yet. Open Settings to enter your web app URL and API
          key.
        </p>
      )}

      {state && (
        <div className="mb-3">
          <select
            aria-label="Pay period"
            className="min-h-11 w-full rounded-lg border border-slate-300 bg-white px-3 text-base dark:border-slate-600 dark:bg-slate-900"
            value={state.sheetName}
            onChange={(e) => void engine.selectSheet(e.target.value)}
          >
            {state.sheets.map((s) => (
              <option key={s.name} value={s.name}>
                {s.name}
              </option>
            ))}
          </select>
        </div>
      )}

      {state ? (
        tab === "expenses" ? (
          <ExpensesView
            engine={engine}
            sheetName={state.sheetName}
            state={state}
            rows={rows}
            snapshot={snapshot}
          />
        ) : (
          <EnvelopesView state={state} warnings={snapshot.warnings} />
        )
      ) : (
        configured && (
          <p className="py-10 text-center text-slate-500">
            {snapshot.refreshing ? "Loading…" : "No data yet."}
          </p>
        )
      )}

      <nav
        aria-label="Views"
        className="pb-safe fixed inset-x-0 bottom-0 z-20 border-t border-slate-200 bg-white/95 pt-2 backdrop-blur dark:border-slate-700 dark:bg-slate-900/95"
      >
        <div className="mx-auto flex max-w-2xl gap-2 px-4">
          {(["expenses", "envelopes"] as const).map((t) => (
            <button
              key={t}
              type="button"
              aria-current={tab === t ? "page" : undefined}
              onClick={() => setTab(t)}
              className={`min-h-12 flex-1 rounded-lg text-sm font-medium ${
                tab === t
                  ? "bg-teal-700 text-white"
                  : "text-slate-700 hover:bg-slate-100 dark:text-slate-200 dark:hover:bg-slate-800"
              }`}
            >
              {t === "expenses" ? "Expenses" : "Envelopes"}
            </button>
          ))}
        </div>
      </nav>

      {settingsOpen && (
        <SettingsDialog
          services={services}
          onClose={() => setSettingsOpen(false)}
          onChanged={() => {
            void vault.hasCredentials().then(setConfigured);
            void engine.refresh().then(() => engine.flush());
          }}
        />
      )}

      {conflict && (
        <ConflictDialog
          conflict={conflict}
          taxMultiplier={state?.taxMultiplier ?? 1}
          onResolve={(choice) => engine.resolveConflict(conflict.id, choice)}
        />
      )}
    </div>
  );
}
