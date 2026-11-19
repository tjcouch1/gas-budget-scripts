import type { EngineSnapshot } from "../sync/types.ts";

export type SyncStatus =
  "conflict" | "offline" | "syncing" | "refreshing" | "synced";

export function statusOf(s: EngineSnapshot): { id: SyncStatus; label: string } {
  const queued = s.queue.length;
  if (s.conflicts.length > 0) return { id: "conflict", label: "Conflict" };
  if (s.offline) {
    return {
      id: "offline",
      label: queued > 0 ? `Offline (${queued} queued)` : "Offline",
    };
  }
  if (queued > 0) {
    return {
      id: "syncing",
      label: `Syncing ${queued} change${queued === 1 ? "" : "s"}…`,
    };
  }
  if (s.refreshing || s.syncingReceipts) {
    return { id: "refreshing", label: "Refreshing…" };
  }
  return { id: "synced", label: "Synced" };
}

const COLORS: Record<SyncStatus, string> = {
  conflict: "bg-red-100 text-red-800 dark:bg-red-900/40 dark:text-red-200",
  offline:
    "bg-amber-100 text-amber-800 dark:bg-amber-900/40 dark:text-amber-200",
  syncing: "bg-sky-100 text-sky-800 dark:bg-sky-900/40 dark:text-sky-200",
  refreshing: "bg-sky-100 text-sky-800 dark:bg-sky-900/40 dark:text-sky-200",
  synced:
    "bg-emerald-100 text-emerald-800 dark:bg-emerald-900/40 dark:text-emerald-200",
};

export function StatusPill({ snapshot }: { snapshot: EngineSnapshot }) {
  const status = statusOf(snapshot);
  return (
    <span
      role="status"
      data-status={status.id}
      className={`rounded-full px-3 py-1 text-xs font-medium ${COLORS[status.id]}`}
    >
      {status.label}
    </span>
  );
}
