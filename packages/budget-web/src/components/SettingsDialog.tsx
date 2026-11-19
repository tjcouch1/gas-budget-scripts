import { useEffect, useRef, useState } from "react";
import type { Services } from "../services.ts";
import { useInlineEnvelopeChips } from "../lib/preferences.ts";
import { useTestingMode } from "../lib/testing-mode.ts";
import { Button, Dialog, inputClass } from "./ui.tsx";

const normalizeUrl = (url: string) => url.trim().replace(/\/+$/, "");

type Message = { kind: "ok" | "error"; text: string } | null;

/**
 * Connection settings. The saved key is never shown or pre-filled: the inputs
 * start empty and plaintext only lives in this form while it is being typed.
 */
export function SettingsDialog({
  services,
  onClose,
  onOpenVariables,
  onChanged,
  autoCloseOnConnect = false,
}: {
  services: Services;
  onClose: () => void;
  /** Opens the Customize variables dialog (the caller closes Settings). */
  onOpenVariables: () => void;
  /** May resolve to whether the saved connection works (true loads data). */
  onChanged: () => void | Promise<boolean>;
  /** Close the dialog ~2 seconds after a save that connects successfully. */
  autoCloseOnConnect?: boolean;
}) {
  const closeTimer = useRef<ReturnType<typeof setTimeout> | undefined>(
    undefined,
  );
  useEffect(() => () => clearTimeout(closeTimer.current), []);
  const [url, setUrl] = useState("");
  const [apiKey, setApiKey] = useState("");
  const [saved, setSaved] = useState(false);
  const [busy, setBusy] = useState(false);
  const [message, setMessage] = useState<Message>(null);
  const [testingMode, setTestingMode] = useTestingMode();
  const [inlineChips, setInlineChips] = useInlineEnvelopeChips();
  // Number of unsynced changes that switching sheets would discard (awaiting a second click).
  const [discardCount, setDiscardCount] = useState<number | null>(null);

  useEffect(() => {
    void services.vault.hasCredentials().then(setSaved);
  }, [services.vault]);

  const candidate = { gasWebAppUrl: url, apiKey };

  async function test() {
    setBusy(true);
    setMessage(null);
    try {
      const state = await services.testConnection(candidate);
      setMessage({
        kind: "ok",
        text: `Connected. ${state.sheets.length} pay-period sheet${state.sheets.length === 1 ? "" : "s"} found.`,
      });
    } catch (e) {
      setMessage({ kind: "error", text: (e as Error).message });
    } finally {
      setBusy(false);
    }
  }

  async function save() {
    setBusy(true);
    setMessage(null);
    try {
      // Plaintext only lives in this function scope.
      const previous = await services.vault.load().catch(() => null);
      if (
        previous &&
        normalizeUrl(previous.gasWebAppUrl) !== normalizeUrl(url)
      ) {
        // Cached rows, queued edits and conflicts belong to the old sheet.
        const snap = services.engine.getSnapshot();
        const unsynced = snap.queue.length + snap.conflicts.length;
        if (unsynced > 0 && discardCount === null) {
          setDiscardCount(unsynced);
          return;
        }
        await services.clearLocalData();
      }
      setDiscardCount(null);
      await services.vault.save(candidate);
      // Ask the browser not to evict our storage (important on iOS Safari).
      if (services.vault.persistent) void navigator.storage?.persist?.();
      setUrl("");
      setApiKey("");
      setSaved(true);
      setMessage({
        kind: "ok",
        text: services.vault.persistent
          ? "Saved (encrypted on this device)."
          : "Connected for this session only. Nothing was saved on this device.",
      });
      const connected = await onChanged();
      if (connected === true && autoCloseOnConnect) {
        setMessage({ kind: "ok", text: "Saved. Connected." });
        clearTimeout(closeTimer.current);
        closeTimer.current = setTimeout(onClose, 2000);
      }
    } catch (e) {
      setMessage({ kind: "error", text: (e as Error).message });
    } finally {
      setBusy(false);
    }
  }

  async function forget() {
    await services.clearLocalData();
    await services.vault.forget();
    setDiscardCount(null);
    setSaved(false);
    setMessage({
      kind: "ok",
      text: services.vault.persistent
        ? "Credentials removed from this device."
        : "Credentials removed from this session.",
    });
    onChanged();
  }

  async function clearData() {
    await services.clearLocalData();
    setMessage({ kind: "ok", text: "Cached budget data cleared." });
  }

  return (
    <Dialog title="Connection settings" onClose={onClose}>
      <form
        className="space-y-3"
        onSubmit={(e) => {
          e.preventDefault();
          void save();
        }}
      >
        {!services.vault.persistent && (
          <div
            role="alert"
            className="rounded-lg border border-amber-400 bg-amber-50 p-3 text-sm text-amber-900 dark:border-amber-600 dark:bg-amber-950 dark:text-amber-200"
          >
            <p className="font-semibold">This page is not a secure context</p>
            <p>
              Browsers only allow encryption on HTTPS pages and on localhost.
              Here your URL and API key are kept in memory for this session
              only: they are not saved and must be entered again each time the
              page reloads. Use the HTTPS (GitHub Pages) address or localhost
              instead to save them encrypted.
            </p>
          </div>
        )}
        <p className="text-sm text-slate-600 dark:text-slate-400">
          {saved
            ? services.vault.persistent
              ? "Credentials are saved on this device. Enter new values to replace them."
              : "Credentials are held in memory for this session. Enter new values to replace them."
            : "Enter the web app URL and API key from your Apps Script deployment."}
        </p>
        <label className="block text-sm font-medium">
          Web app URL
          <input
            className={`${inputClass} mt-1`}
            type="url"
            autoComplete="off"
            spellCheck={false}
            placeholder="https://script.google.com/macros/s/<deployment-id>/exec"
            value={url}
            onChange={(e) => {
              setUrl(e.target.value);
              setDiscardCount(null);
            }}
          />
        </label>
        <label className="block text-sm font-medium">
          API key
          <input
            className={`${inputClass} mt-1`}
            type="password"
            autoComplete="off"
            spellCheck={false}
            placeholder={saved ? "•••••••• (saved)" : ""}
            value={apiKey}
            onChange={(e) => setApiKey(e.target.value)}
          />
        </label>
        {discardCount !== null && (
          <p role="alert" className="text-sm text-red-600">
            {discardCount} unsynced change{discardCount === 1 ? "" : "s"} from
            the old sheet will be discarded.
          </p>
        )}
        {message && (
          <p
            role={message.kind === "error" ? "alert" : "status"}
            className={`text-sm ${message.kind === "error" ? "text-red-600" : "text-emerald-700 dark:text-emerald-400"}`}
          >
            {message.text}
          </p>
        )}
        <div className="flex flex-wrap gap-2">
          <Button variant="primary" type="submit" disabled={busy}>
            {discardCount !== null ? "Discard and switch" : "Save"}
          </Button>
          <Button
            onClick={() => void test()}
            disabled={busy || !url || !apiKey}
          >
            Test connection
          </Button>
        </div>
      </form>
      <label className="flex items-start gap-2 text-sm">
        <input
          type="checkbox"
          className="mt-1"
          checked={testingMode}
          onChange={(e) => setTestingMode(e.target.checked)}
        />
        <span>
          Testing mode: import receipts without marking them in Gmail
          <span className="block text-xs text-slate-500">
            Unmarked receipts are imported again on the next import, so rows may
            duplicate.
          </span>
        </span>
      </label>
      <label className="mt-3 flex items-start gap-2 text-sm">
        <input
          type="checkbox"
          className="mt-1"
          checked={inlineChips}
          onChange={(e) => setInlineChips(e.target.checked)}
        />
        <span>
          Choose envelopes with chips instead of the dropdown
          <span className="block text-xs text-slate-500">
            Uncategorized rows in the list show a scrolling row of envelope
            chips under the transaction.
          </span>
        </span>
      </label>
      <hr className="my-4 border-slate-200 dark:border-slate-700" />
      <div className="mb-3">
        <Button
          onClick={onOpenVariables}
          disabled={!saved}
          title={
            saved
              ? undefined
              : "Save your connection first; the variables are loaded from the sheet."
          }
        >
          Customize variables
        </Button>
      </div>
      <div className="flex flex-wrap gap-2">
        <Button
          variant="danger"
          onClick={() => void forget()}
          disabled={!saved}
        >
          Forget credentials
        </Button>
        <Button onClick={() => void clearData()}>Clear local data</Button>
      </div>
      <p className="mt-4 text-xs text-slate-500">
        <a
          className="underline"
          href="./licenses.html"
          target="_blank"
          rel="noreferrer"
        >
          Open source licenses
        </a>
      </p>
    </Dialog>
  );
}
