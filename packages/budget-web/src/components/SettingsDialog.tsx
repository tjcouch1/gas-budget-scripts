import { useEffect, useState } from "react";
import type { Services } from "../services.ts";
import { Button, Dialog, inputClass } from "./ui.tsx";

type Message = { kind: "ok" | "error"; text: string } | null;

/**
 * Connection settings. The saved key is never shown or pre-filled: the inputs
 * start empty and plaintext only lives in this form while it is being typed.
 */
export function SettingsDialog({
  services,
  onClose,
  onChanged,
}: {
  services: Services;
  onClose: () => void;
  onChanged: () => void;
}) {
  const [url, setUrl] = useState("");
  const [apiKey, setApiKey] = useState("");
  const [saved, setSaved] = useState(false);
  const [busy, setBusy] = useState(false);
  const [message, setMessage] = useState<Message>(null);

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
      await services.vault.save(candidate);
      // Ask the browser not to evict our storage (important on iOS Safari).
      void navigator.storage?.persist?.();
      setUrl("");
      setApiKey("");
      setSaved(true);
      setMessage({ kind: "ok", text: "Saved (encrypted on this device)." });
      onChanged();
    } catch (e) {
      setMessage({ kind: "error", text: (e as Error).message });
    } finally {
      setBusy(false);
    }
  }

  async function forget() {
    await services.vault.forget();
    setSaved(false);
    setMessage({ kind: "ok", text: "Credentials removed from this device." });
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
        <p className="text-sm text-slate-600 dark:text-slate-400">
          {saved
            ? "Credentials are saved on this device. Enter new values to replace them."
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
            onChange={(e) => setUrl(e.target.value)}
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
            Save
          </Button>
          <Button
            onClick={() => void test()}
            disabled={busy || !url || !apiKey}
          >
            Test connection
          </Button>
        </div>
      </form>
      <hr className="my-4 border-slate-200 dark:border-slate-700" />
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
