import { useCallback, useEffect, useId, useMemo, useState } from "react";
import {
  compileReceiptRules,
  DEFAULT_RECEIPT_RULES,
  validateReceiptRulesConfig,
  type CompiledReceiptRules,
  type ReceiptEmailDto,
  type ReceiptRulesConfig,
  type ReceiptRulesStateDto,
  type RuleIssue,
} from "@budget/shared";
import { ApiError } from "../../api/client.ts";
import { previewReceiptRules } from "../../lib/receipt-rules-preview.ts";
import type { Services } from "../../services.ts";
import {
  Button,
  ConfirmDialog,
  Dialog,
  UnsavedPrompt,
  useUnsavedGuard,
} from "../ui.tsx";
import { EmailFetcher } from "./EmailFetcher.tsx";
import { EmailResults } from "./EmailResults.tsx";
import { RulesEditor, tabKeyDown } from "./RulesEditor.tsx";

type Section = "rules" | "emails" | "results";

type LoadState =
  | { kind: "loading" }
  | { kind: "error"; message: string }
  | { kind: "loaded"; state: ReceiptRulesStateDto };

type SaveState =
  | { kind: "idle" }
  | { kind: "saving" }
  | { kind: "saved" }
  | { kind: "conflict"; message: string }
  | { kind: "error"; message: string; issues: RuleIssue[] };

const errorMessage = (error: unknown) =>
  error instanceof Error ? error.message : String(error);

/** Edits the ReceiptRules global variable with a live preview over fetched emails. */
export function CustomizeVariablesDialog({
  services,
  onClose,
}: {
  services: Pick<Services, "api">;
  onClose: () => void;
}) {
  const { api } = services;
  const ids = useId();
  const [load, setLoad] = useState<LoadState>({ kind: "loading" });
  /** Bumped on every (re)load so the editor starts fresh. */
  const [loadId, setLoadId] = useState(0);
  const [candidate, setCandidate] = useState<ReceiptRulesConfig>(
    DEFAULT_RECEIPT_RULES,
  );
  const [dirty, setDirty] = useState(false);
  const [emails, setEmails] = useState<ReceiptEmailDto[]>([]);
  const [section, setSection] = useState<Section>("rules");
  const [save, setSave] = useState<SaveState>({ kind: "idle" });
  const [confirmClear, setConfirmClear] = useState(false);

  const applyState = useCallback((state: ReceiptRulesStateDto) => {
    setLoad({ kind: "loaded", state });
    setCandidate(state.config ?? DEFAULT_RECEIPT_RULES);
    setDirty(false);
  }, []);

  const reload = useCallback(async () => {
    setLoad({ kind: "loading" });
    setSave({ kind: "idle" });
    try {
      const { data } = await api<ReceiptRulesStateDto>({
        action: "getReceiptRules",
      });
      applyState(data);
      setLoadId((n) => n + 1);
    } catch (error) {
      setLoad({ kind: "error", message: errorMessage(error) });
    }
  }, [api, applyState]);

  useEffect(() => {
    void reload();
  }, [reload]);

  const loaded = load.kind === "loaded" ? load.state : null;
  const forwardEmailAddress = loaded?.forwardEmailAddress ?? "";
  const forwardName = loaded?.forwardName ?? "";

  const validation = useMemo(
    () =>
      validateReceiptRulesConfig(candidate, {
        forwardEmailAddress: forwardEmailAddress || undefined,
      }),
    [candidate, forwardEmailAddress],
  );
  const issues = validation.ok ? [] : validation.issues;
  const compiled = useMemo((): CompiledReceiptRules | null => {
    if (!validation.ok) return null;
    try {
      return compileReceiptRules(validation.config, {
        forwardEmailAddress: forwardEmailAddress || undefined,
      });
    } catch {
      return null;
    }
  }, [validation, forwardEmailAddress]);
  const preview = useMemo(
    () =>
      previewReceiptRules(compiled, emails, {
        forwardEmailAddress,
        forwardName,
      }),
    [compiled, emails, forwardEmailAddress, forwardName],
  );

  const busy = save.kind === "saving";
  const canSave = loaded !== null && dirty && issues.length === 0 && !busy;

  const send = async (rules: ReceiptRulesConfig | null): Promise<boolean> => {
    if (!loaded) return false;
    setSave({ kind: "saving" });
    try {
      const { data } = await api<ReceiptRulesStateDto>({
        action: "saveReceiptRules",
        rules,
        expectedRaw: loaded.raw,
      });
      applyState(data);
      setSave({ kind: "saved" });
      return true;
    } catch (error) {
      if (error instanceof ApiError && error.code === "conflict")
        setSave({ kind: "conflict", message: error.message });
      else
        setSave({
          kind: "error",
          message: errorMessage(error),
          issues: error instanceof ApiError ? (error.issues ?? []) : [],
        });
      return false;
    }
  };

  const guard = useUnsavedGuard({ dirty, onDiscard: onClose });

  const edit = (config: ReceiptRulesConfig) => {
    setCandidate(config);
    setDirty(true);
    if (save.kind !== "idle" && save.kind !== "saving")
      setSave({ kind: "idle" });
  };

  const sections: { id: Section; label: string }[] = [
    { id: "rules", label: "Rules" },
    { id: "emails", label: "Emails" },
    { id: "results", label: `Results (${preview.length})` },
  ];
  const panelProps = (id: Section) => ({
    role: "tabpanel" as const,
    id: `${ids}-panel-${id}`,
    "aria-label": sections.find((s) => s.id === id)?.label,
    "data-testid": `panel-${id}`,
    className: `${section === id ? "" : "hidden"} space-y-3 lg:block`,
  });

  return (
    <Dialog title="Customize variables" wide="xl" onClose={guard.requestClose}>
      {load.kind === "loading" && (
        <p role="status" className="text-sm text-slate-600 dark:text-slate-400">
          Loading the rules from the sheet…
        </p>
      )}
      {load.kind === "error" && (
        <div className="space-y-2">
          <p role="alert" className="text-sm text-red-600 dark:text-red-400">
            {load.message}
          </p>
          <Button onClick={() => void reload()}>Retry</Button>
        </div>
      )}
      {loaded && (
        <div className="space-y-3">
          {loaded.issues.length > 0 && (
            <div className="rounded-lg bg-amber-50 p-3 text-sm text-amber-900 dark:bg-amber-950/50 dark:text-amber-100">
              <p className="font-medium">
                The ReceiptRules value on the sheet is invalid, so receipt
                imports fail until it is fixed or cleared. The editor starts
                from the built-in defaults; the sheet&apos;s text is in the JSON
                tab.
              </p>
              <ul className="mt-1 list-disc pl-5">
                {loaded.issues.map((issue, i) => (
                  <li key={i} className="break-words">
                    <code>{issue.path}</code>: {issue.message}
                  </li>
                ))}
              </ul>
            </div>
          )}
          {loaded.isDefault && loaded.issues.length === 0 && (
            <p className="text-sm text-slate-600 dark:text-slate-400">
              The sheet uses the built-in defaults.
            </p>
          )}

          <div
            role="tablist"
            aria-label="Customize variables sections"
            className="flex gap-1 lg:hidden"
          >
            {sections.map((s) => (
              <button
                key={s.id}
                type="button"
                role="tab"
                id={`${ids}-tab-${s.id}`}
                aria-controls={`${ids}-panel-${s.id}`}
                aria-selected={section === s.id}
                tabIndex={section === s.id ? 0 : -1}
                onKeyDown={(e) =>
                  tabKeyDown(
                    e,
                    sections.map((x) => x.id),
                    section,
                    setSection,
                    (id) => `${ids}-tab-${id}`,
                  )
                }
                onClick={() => setSection(s.id)}
                className={`min-h-11 flex-1 rounded-lg px-3 text-sm font-medium ${
                  section === s.id
                    ? "bg-teal-700 text-white"
                    : "bg-slate-100 text-slate-700 hover:bg-slate-200 dark:bg-slate-800 dark:text-slate-200 dark:hover:bg-slate-700"
                }`}
              >
                {s.label}
              </button>
            ))}
          </div>

          <div className="lg:grid lg:grid-cols-2 lg:items-start lg:gap-4">
            <div {...panelProps("rules")}>
              <h3 className="hidden text-base font-semibold lg:block">Rules</h3>
              <RulesEditor
                key={loadId}
                config={candidate}
                onChange={edit}
                issues={issues}
                forwardEmailAddress={forwardEmailAddress}
                forwardName={forwardName}
                jsonSeed={loaded.config === null ? loaded.raw : undefined}
              />
            </div>
            <div className="space-y-4">
              <div {...panelProps("emails")}>
                <h3 className="hidden text-base font-semibold lg:block">
                  Test emails
                </h3>
                <EmailFetcher
                  senders={candidate.senders.map((s) => s.address)}
                  forwardEmailAddress={forwardEmailAddress}
                  api={api}
                  onFetched={(fetched) => setEmails(fetched)}
                />
              </div>
              <div {...panelProps("results")}>
                <h3 className="hidden text-base font-semibold lg:block">
                  Results ({preview.length})
                </h3>
                <EmailResults items={preview} />
              </div>
            </div>
          </div>

          <div className="sticky bottom-0 -mx-4 space-y-2 border-t border-slate-200 bg-white/95 px-4 py-3 dark:border-slate-700 dark:bg-slate-900/95">
            {save.kind === "saved" && (
              <p
                role="status"
                className="text-sm text-teal-700 dark:text-teal-300"
              >
                Saved to the Variables sheet
              </p>
            )}
            {save.kind === "conflict" && (
              <div
                role="alert"
                className="flex flex-wrap items-center gap-2 text-sm text-red-600 dark:text-red-400"
              >
                <span>{save.message}</span>
                <Button onClick={() => void reload()}>Reload</Button>
              </div>
            )}
            {save.kind === "error" && (
              <div
                role="alert"
                className="text-sm text-red-600 dark:text-red-400"
              >
                <p>{save.message}</p>
                {save.issues.length > 0 && (
                  <ul className="mt-1 list-disc pl-5">
                    {save.issues.map((issue, i) => (
                      <li key={i} className="break-words">
                        <code>{issue.path}</code>: {issue.message}
                      </li>
                    ))}
                  </ul>
                )}
              </div>
            )}
            <div className="flex flex-wrap gap-2">
              <Button
                variant="primary"
                disabled={!canSave}
                onClick={() => void send(candidate)}
              >
                {busy ? "Saving…" : "Save to sheet"}
              </Button>
              <Button
                disabled={busy}
                onClick={() => edit(DEFAULT_RECEIPT_RULES)}
              >
                Reset to built-in defaults
              </Button>
              <Button
                variant="ghost"
                disabled={busy || loaded.raw === ""}
                onClick={() => setConfirmClear(true)}
              >
                Clear variable (use built-in defaults)
              </Button>
            </div>
          </div>
        </div>
      )}

      {confirmClear && (
        <ConfirmDialog
          title="Clear the ReceiptRules variable?"
          confirmLabel="Clear variable"
          onConfirm={() => {
            setConfirmClear(false);
            void send(null);
          }}
          onCancel={() => setConfirmClear(false)}
        >
          The cell on the Variables sheet is emptied and receipt imports use the
          built-in defaults, including future updates to them.
        </ConfirmDialog>
      )}
      {guard.prompting && (
        <UnsavedPrompt
          canSave={canSave}
          onSave={() => {
            guard.setPrompting(false);
            void send(candidate).then((ok) => {
              if (ok) onClose();
            });
          }}
          onDiscard={onClose}
          onBack={() => guard.setPrompting(false)}
        />
      )}
    </Dialog>
  );
}
