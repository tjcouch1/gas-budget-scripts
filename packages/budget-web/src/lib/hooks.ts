import { useEffect, useState, useSyncExternalStore } from "react";
import type { SyncEngine } from "../sync/engine.ts";

export function useEngineSnapshot(engine: SyncEngine) {
  return useSyncExternalStore(engine.subscribe, engine.getSnapshot);
}

export type Tab = "expenses" | "envelopes";

function readTab(): Tab {
  return window.location.hash.startsWith("#/envelopes")
    ? "envelopes"
    : "expenses";
}

/** Hash-based tab routing: deep links work on GitHub Pages without a 404 fallback. */
export function useHashTab(): [Tab, (tab: Tab) => void] {
  const [tab, setTab] = useState<Tab>(readTab);
  useEffect(() => {
    const onHash = () => setTab(readTab());
    window.addEventListener("hashchange", onHash);
    return () => window.removeEventListener("hashchange", onHash);
  }, []);
  return [
    tab,
    (next) => {
      window.location.hash = `#/${next}`;
      setTab(next);
    },
  ];
}

export function useOnlineEvents(engine: SyncEngine): void {
  useEffect(() => {
    const online = () => engine.setOnline(true);
    const offline = () => engine.setOnline(false);
    const focus = () => {
      if (document.visibilityState === "visible" && navigator.onLine) {
        void engine.refresh().then(() => engine.flush());
      }
    };
    window.addEventListener("online", online);
    window.addEventListener("offline", offline);
    document.addEventListener("visibilitychange", focus);
    return () => {
      window.removeEventListener("online", online);
      window.removeEventListener("offline", offline);
      document.removeEventListener("visibilitychange", focus);
    };
  }, [engine]);
}
