import type {
  AllTransactionsDto,
  PayPeriodStateDto,
  SplitItemInput,
  TransactionChanges,
  TransactionRowDto,
} from "@budget/shared";

interface MutationBase {
  id: string;
  sheetName: string;
  /** 0-based row index as of the last pull; re-based after a split. */
  rowIndex: number;
  /** Fingerprint of the row at the client's last pull (the server's conflict check). */
  baseFingerprint: string;
  /** The row as the user saw it when making the change ("Original Value"). */
  baseRow: TransactionRowDto;
  createdAt: number;
}

export interface UpdateMutation extends MutationBase {
  action: "updateTransaction";
  payload: { changes: TransactionChanges };
}

export interface SplitMutation extends MutationBase {
  action: "splitTransaction";
  payload: {
    splits: SplitItemInput[];
    /** New item note after " | " for the original row (written with the split). */
    parentItemNote?: string;
    /** New envelope for the original row (written with the split). */
    parentEnvelope?: string;
  };
}

export type QueuedMutation = UpdateMutation | SplitMutation;

/** A mutation the server rejected because the row changed since `baseRow`. */
export interface Conflict {
  id: string;
  mutation: QueuedMutation;
  /** The sheet's current row ("Current Remote Value"); null if the row can no longer be found. */
  current: TransactionRowDto | null;
}

/** A row the sheet confirmed a moment ago; the list shows "Saved ✓" for it briefly. */
export interface SavedMark {
  sheetName: string;
  rowIndex: number;
}

export interface Notice {
  kind: "info" | "error";
  message: string;
  /** Set for failures that need the owner to fix the sheet or Settings. */
  hint?: "settings" | "no-sheet";
}

export interface EngineSnapshot {
  /** True once the cache has been loaded. */
  ready: boolean;
  currentSheet: string | null;
  states: Record<string, PayPeriodStateDto>;
  /** Every pay-period tab's transactions; loaded on demand by the "All pay periods" view. */
  all: AllTransactionsDto | null;
  /** True while `getAllTransactions` is in flight. */
  loadingAll: boolean;
  queue: QueuedMutation[];
  conflicts: Conflict[];
  /** Rows whose changes reached the sheet in the last few seconds (not persisted). */
  recentlySaved: SavedMark[];
  refreshing: boolean;
  flushing: boolean;
  syncingReceipts: boolean;
  offline: boolean;
  notice: Notice | null;
  /** Warnings from the last getState (e.g. envelopes missing on this tab). */
  warnings: string[];
}

export interface PersistedEngine {
  currentSheet: string | null;
  states: Record<string, PayPeriodStateDto>;
  all?: AllTransactionsDto | null;
  queue: QueuedMutation[];
  conflicts: Conflict[];
}
