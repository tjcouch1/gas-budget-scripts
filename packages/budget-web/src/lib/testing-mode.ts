import { useSyncExternalStore } from "react";

/** Per-device preference (not a secret): import receipts without marking them in Gmail. */
export const TESTING_MODE_KEY = "budget.testingMode";

const listeners = new Set<() => void>();

export function getTestingMode(): boolean {
  try {
    return localStorage.getItem(TESTING_MODE_KEY) === "1";
  } catch {
    return false;
  }
}

export function setTestingMode(on: boolean): void {
  try {
    if (on) localStorage.setItem(TESTING_MODE_KEY, "1");
    else localStorage.removeItem(TESTING_MODE_KEY);
  } catch {
    // storage unavailable: the preference just does not persist
  }
  listeners.forEach((l) => l());
}

function subscribe(listener: () => void): () => void {
  listeners.add(listener);
  window.addEventListener("storage", listener);
  return () => {
    listeners.delete(listener);
    window.removeEventListener("storage", listener);
  };
}

export function useTestingMode(): [boolean, (on: boolean) => void] {
  return [useSyncExternalStore(subscribe, getTestingMode), setTestingMode];
}
