import type {
  AllTransactionsDto,
  NewTransactionFields,
  PayPeriodStateDto,
  SplitGroupRowChanges,
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

/** One row of a queued delete, with its fingerprint at the last pull (re-based like `MutationBase`). */
export interface DeleteRowBase {
  rowIndex: number;
  baseFingerprint: string;
  baseRow: TransactionRowDto;
}

/**
 * Deletes one row or a whole split group. `rowIndex`/`baseFingerprint`/`baseRow` mirror
 * `payload.rows[0]`; every row is re-based by fingerprint when earlier changes move rows.
 */
export interface DeleteMutation extends MutationBase {
  action: "deleteTransaction";
  payload: { rows: DeleteRowBase[] };
}

/** One row of a queued split group save: its base (re-based like `DeleteRowBase`) and its own edits. */
export interface GroupRowBase extends DeleteRowBase {
  changes?: SplitGroupRowChanges;
}

/**
 * Saves every edit of one split group atomically (`updateSplitGroup`): every row of the group
 * with its fingerprint, the group's new title and total, and new split items.
 * `rowIndex`/`baseFingerprint`/`baseRow` mirror `payload.rows[0]` (the split parent).
 */
export interface GroupMutation extends MutationBase {
  action: "updateSplitGroup";
  payload: {
    rows: GroupRowBase[];
    /** New text before " | " for every row. */
    title?: string;
    /** The parent's new total (a plain amount). */
    splitTotal?: string;
    /** New split items appended to the group. */
    newItems?: SplitItemInput[];
  };
}

/** A mutation of existing rows (it carries base rows and can conflict). */
export type RowMutation =
  UpdateMutation | SplitMutation | DeleteMutation | GroupMutation;

/**
 * Appends a new row ("New expense"). It has no base row and never conflicts: the server
 * writes it after the last non-empty row, so no other row moves. `requestId` stays the same
 * across retries, so the server inserts it at most once.
 */
export interface AddMutation {
  id: string;
  sheetName: string;
  createdAt: number;
  action: "addTransaction";
  payload: { requestId: string; fields: NewTransactionFields };
}

export type QueuedMutation = RowMutation | AddMutation;

/** A mutation the server rejected because the row changed since `baseRow`. */
export interface Conflict {
  id: string;
  mutation: RowMutation;
  /** The sheet's current row ("Current Remote Value"); null if the row can no longer be found. */
  current: TransactionRowDto | null;
  /** For a split group save: the group's current rows (rows that are now empty are missing). */
  currentRows?: TransactionRowDto[];
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
  hint?: "settings" | "no-sheet" | "variables";
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
  /** True when the last multi-period load stopped early (the periods that arrived are kept). */
  loadAllFailed: boolean;
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
