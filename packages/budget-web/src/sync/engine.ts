import type {
  AddTransactionResultDto,
  AllTransactionsDto,
  NewTransactionFields,
  PayPeriodStateDto,
  SplitItemInput,
  SyncReceiptsResultDto,
  TransactionChanges,
  TransactionRowDto,
} from "@budget/shared";
import { ApiError, type CallApi } from "../api/client.ts";
import type { KeyValueStore } from "../crypto/kv-store.ts";
import type {
  AddMutation,
  Conflict,
  DeleteMutation,
  DeleteRowBase,
  EngineSnapshot,
  Notice,
  PersistedEngine,
  QueuedMutation,
  RowMutation,
  SavedMark,
  SplitMutation,
  UpdateMutation,
} from "./types.ts";

const PERSIST_KEY = "engine";
const DEFAULT_RETRY_DELAYS_MS = [2000, 5000, 15000, 30000];
/** How long a row shows "Saved ✓". */
const SAVED_MARK_MS = 3000;

export interface SyncEngineOptions {
  callApi: CallApi;
  /** Cache and queue storage (unencrypted budget data; see web-plan.md section 4.5). */
  store: KeyValueStore;
  /** Backoff delays for transient failures; an empty array disables automatic retries. */
  retryDelaysMs?: number[];
  now?: () => number;
  newId?: () => string;
  /** Idempotency key of a new row (`addTransaction`); default `crypto.randomUUID()`. */
  newRequestId?: () => string;
}

/** Updates return the changed row; splits return the refreshed state. Accept both shapes defensively. */
function isState(data: unknown): data is PayPeriodStateDto {
  return (
    typeof data === "object" &&
    data !== null &&
    Array.isArray((data as PayPeriodStateDto).transactions)
  );
}

function isRow(data: unknown): data is TransactionRowDto {
  return (
    typeof data === "object" &&
    data !== null &&
    typeof (data as TransactionRowDto).fingerprint === "string"
  );
}

function replaceInList(
  list: (TransactionRowDto | null)[],
  row: TransactionRowDto,
): (TransactionRowDto | null)[] {
  const transactions = [...list];
  while (transactions.length <= row.rowIndex) transactions.push(null);
  transactions[row.rowIndex] = row;
  return transactions;
}

function replaceRow(
  state: PayPeriodStateDto,
  row: TransactionRowDto,
): PayPeriodStateDto {
  return { ...state, transactions: replaceInList(state.transactions, row) };
}

/** Replaces one sheet's transactions in the all-sheets cache (no-op when it is not cached). */
function patchAll(
  all: AllTransactionsDto | null,
  sheetName: string,
  update: (
    transactions: (TransactionRowDto | null)[],
  ) => (TransactionRowDto | null)[],
): AllTransactionsDto | null {
  if (!all || !all.sheets.some((x) => x.name === sheetName)) return all;
  return {
    ...all,
    sheets: all.sheets.map((x) =>
      x.name === sheetName ? { ...x, transactions: update(x.transactions) } : x,
    ),
  };
}

/** The rows a mutation is based on: a delete's rows, or its one row. */
function baseRowsOf(q: QueuedMutation): DeleteRowBase[] {
  if (q.action === "addTransaction") return [];
  return q.action === "deleteTransaction"
    ? q.payload.rows
    : [
        {
          rowIndex: q.rowIndex,
          baseFingerprint: q.baseFingerprint,
          baseRow: q.baseRow,
        },
      ];
}

/** `q` re-based onto `rows` (the mirrored top-level fields follow the first row). */
function withBaseRows(q: RowMutation, rows: DeleteRowBase[]): RowMutation {
  const [first] = rows;
  const top = {
    rowIndex: first.rowIndex,
    baseFingerprint: first.baseFingerprint,
    baseRow: first.baseRow,
  };
  return q.action === "deleteTransaction"
    ? { ...q, ...top, payload: { rows } }
    : { ...q, ...top };
}

export class SyncEngine {
  private snapshot: EngineSnapshot = {
    ready: false,
    currentSheet: null,
    states: {},
    all: null,
    loadingAll: false,
    queue: [],
    conflicts: [],
    recentlySaved: [],
    refreshing: false,
    flushing: false,
    syncingReceipts: false,
    offline: false,
    notice: null,
    warnings: [],
  };
  private listeners = new Set<() => void>();
  private tail: Promise<unknown> = Promise.resolve();
  private flushPromise: Promise<void> | null = null;
  private inFlightId: string | null = null;
  private retryAttempt = 0;
  private retryTimer: ReturnType<typeof setTimeout> | null = null;
  private readonly retryDelays: number[];
  private readonly now: () => number;
  private readonly newId: () => string;
  private readonly newRequestId: () => string;
  private idCounter = 0;

  constructor(private readonly options: SyncEngineOptions) {
    this.retryDelays = options.retryDelaysMs ?? DEFAULT_RETRY_DELAYS_MS;
    this.now = options.now ?? Date.now;
    this.newId =
      options.newId ??
      (() => `${this.now().toString(36)}-${(this.idCounter++).toString(36)}`);
    this.newRequestId =
      options.newRequestId ?? (() => globalThis.crypto.randomUUID());
  }

  // ---- external store plumbing -------------------------------------------

  subscribe = (listener: () => void): (() => void) => {
    this.listeners.add(listener);
    return () => this.listeners.delete(listener);
  };

  getSnapshot = (): EngineSnapshot => this.snapshot;

  private set(patch: Partial<EngineSnapshot>, persist = false): void {
    this.snapshot = { ...this.snapshot, ...patch };
    for (const l of this.listeners) l();
    if (persist) void this.persist();
  }

  private async persist(): Promise<void> {
    const s = this.snapshot;
    const data: PersistedEngine = {
      currentSheet: s.currentSheet,
      states: s.states,
      all: s.all,
      queue: s.queue,
      conflicts: s.conflicts,
    };
    try {
      await this.options.store.set(PERSIST_KEY, data);
    } catch {
      // Storage may be unavailable (private mode, quota). The app still works in memory.
    }
  }

  private run<T>(fn: () => Promise<T>): Promise<T> {
    const result = this.tail.then(fn, fn);
    this.tail = result.catch(() => undefined);
    return result;
  }

  // ---- lifecycle ---------------------------------------------------------

  /** Loads the cache so the UI can render immediately, then refreshes and flushes. */
  async init(): Promise<void> {
    try {
      const saved = await this.options.store.get<PersistedEngine>(PERSIST_KEY);
      if (saved) {
        this.snapshot = {
          ...this.snapshot,
          currentSheet: saved.currentSheet,
          states: saved.states ?? {},
          all: saved.all ?? null,
          queue: saved.queue ?? [],
          conflicts: saved.conflicts ?? [],
        };
      }
    } catch {
      // Ignore a broken cache.
    }
    this.set({ ready: true });
    await this.refresh(this.snapshot.currentSheet ?? undefined);
    await this.flush();
  }

  setOnline(online: boolean): void {
    if (!online) {
      this.set({ offline: true });
      return;
    }
    this.set({ offline: false });
    this.retryAttempt = 0;
    void this.refresh().then(() => this.flush());
  }

  dismissNotice(): void {
    this.set({ notice: null });
  }

  /** Removes every cached state, queued change, and conflict. */
  async clearLocalData(): Promise<void> {
    const wipe = () =>
      this.set({
        currentSheet: null,
        states: {},
        all: null,
        queue: [],
        conflicts: [],
        notice: null,
        warnings: [],
      });
    if (this.retryTimer) clearTimeout(this.retryTimer);
    this.retryTimer = null;
    this.retryAttempt = 0;
    // Empty the queue first so a running flush loop sends nothing more, then let any
    // in-flight request settle and wipe again so its result cannot repopulate old data.
    wipe();
    await this.tail;
    if (this.retryTimer) clearTimeout(this.retryTimer);
    this.retryTimer = null;
    wipe();
    await this.options.store.clear();
  }

  // ---- reading -----------------------------------------------------------

  selectSheet(name: string): Promise<void> {
    this.set({ currentSheet: name }, true);
    return this.refresh(name).then(() => undefined);
  }

  /** Pulls state from the server (stale-while-revalidate: the cache stays visible meanwhile). */
  refresh(sheetName?: string): Promise<boolean> {
    return this.run(async () => {
      const target = sheetName ?? this.snapshot.currentSheet ?? undefined;
      this.set({ refreshing: true });
      const fetchState = (name?: string) =>
        this.options.callApi<PayPeriodStateDto>(
          name
            ? { action: "getState", sheetName: name }
            : { action: "getState" },
        );
      try {
        let result;
        try {
          result = await fetchState(target);
        } catch (error) {
          if (
            !target ||
            !(error instanceof ApiError) ||
            error.code !== "not_found"
          )
            throw error;
          // The cached pay period is gone from the spreadsheet: drop it and show the default one.
          const states = { ...this.snapshot.states };
          delete states[target];
          this.set({ states, currentSheet: null }, true);
          result = await fetchState();
          this.set({
            notice: {
              kind: "info",
              message: `Pay period "${target}" is no longer in the spreadsheet, so the current pay period is shown.`,
            },
          });
        }
        const { data, warnings } = result;
        this.acceptState(data);
        // A successful read means the connection works now, so a "go to Settings" notice
        // (e.g. from the first load before credentials were saved) is stale.
        const notice =
          this.snapshot.notice?.hint === "settings"
            ? null
            : this.snapshot.notice;
        this.set({ offline: false, warnings, notice }, true);
        this.retryAttempt = 0;
        return true;
      } catch (error) {
        this.handleReadError(error);
        return false;
      } finally {
        this.set({ refreshing: false });
      }
    });
  }

  /**
   * Loads every pay-period tab's transactions (the "All pay periods" view). On demand only:
   * it reads every tab, so it is heavier than `refresh`. The cache stays visible meanwhile.
   */
  loadAll(): Promise<void> {
    return this.run(async () => {
      this.set({ loadingAll: true });
      try {
        const { data } = await this.options.callApi<AllTransactionsDto>({
          action: "getAllTransactions",
        });
        this.set({ all: data, offline: false }, true);
        this.retryAttempt = 0;
      } catch (error) {
        this.handleReadError(error);
      } finally {
        this.set({ loadingAll: false });
      }
    });
  }

  private acceptState(state: PayPeriodStateDto): void {
    const current =
      this.snapshot.currentSheet &&
      this.snapshot.states[this.snapshot.currentSheet]
        ? this.snapshot.currentSheet
        : state.sheetName;
    this.set(
      {
        states: { ...this.snapshot.states, [state.sheetName]: state },
        all: patchAll(
          this.snapshot.all,
          state.sheetName,
          () => state.transactions,
        ),
        currentSheet: this.snapshot.currentSheet ?? current,
      },
      true,
    );
  }

  private handleReadError(error: unknown): void {
    if (error instanceof ApiError) {
      if (error.isTransient) {
        if (error.code !== "busy") this.set({ offline: true });
        this.scheduleRetry();
        return;
      }
      this.set({
        notice: {
          kind: "error",
          message: error.message,
          hint:
            error.code === "unauthorized" || error.code === "not_configured"
              ? "settings"
              : undefined,
        },
      });
      return;
    }
    this.set({ notice: { kind: "error", message: String(error) } });
  }

  // ---- writing -----------------------------------------------------------

  private chainBase(
    sheetName: string,
    rowIndex: number,
    row: TransactionRowDto,
  ) {
    // If an earlier queued change touches the same row, ride on its base fingerprint;
    // it is re-based to the server's new fingerprint once that change succeeds.
    for (const m of [...this.snapshot.queue].reverse()) {
      if (m.sheetName !== sheetName) continue;
      const base = baseRowsOf(m).find((r) => r.rowIndex === rowIndex);
      if (base) return base.baseFingerprint;
    }
    return row.fingerprint;
  }

  /** Queues a field edit. Successive edits to the same row are merged. */
  enqueueUpdate(
    sheetName: string,
    row: TransactionRowDto,
    changes: TransactionChanges,
  ): void {
    const queue = this.snapshot.queue;
    const mergeIndex = queue.findIndex(
      (m) =>
        m.action === "updateTransaction" &&
        m.sheetName === sheetName &&
        m.rowIndex === row.rowIndex &&
        m.id !== this.inFlightId &&
        // Only merge when nothing later in the queue touches the same row.
        !queue
          .slice(queue.indexOf(m) + 1)
          .some(
            (n) =>
              n.sheetName === sheetName &&
              n.action !== "addTransaction" &&
              n.rowIndex === row.rowIndex,
          ),
    );
    if (mergeIndex >= 0) {
      const existing = queue[mergeIndex] as UpdateMutation;
      const merged: UpdateMutation = {
        ...existing,
        payload: { changes: { ...existing.payload.changes, ...changes } },
      };
      const next = [...queue];
      next[mergeIndex] = merged;
      this.set({ queue: next }, true);
    } else {
      const mutation: UpdateMutation = {
        id: this.newId(),
        sheetName,
        rowIndex: row.rowIndex,
        baseFingerprint: this.chainBase(sheetName, row.rowIndex, row),
        baseRow: row,
        createdAt: this.now(),
        action: "updateTransaction",
        payload: { changes },
      };
      this.set({ queue: [...queue, mutation] }, true);
    }
    void this.flush();
  }

  enqueueSplit(
    sheetName: string,
    row: TransactionRowDto,
    splits: SplitItemInput[],
    parent: { itemNote?: string; envelope?: string } = {},
  ): void {
    const mutation: SplitMutation = {
      id: this.newId(),
      sheetName,
      rowIndex: row.rowIndex,
      baseFingerprint: this.chainBase(sheetName, row.rowIndex, row),
      baseRow: row,
      createdAt: this.now(),
      action: "splitTransaction",
      payload: {
        splits,
        ...(parent.itemNote !== undefined
          ? { parentItemNote: parent.itemNote }
          : {}),
        ...(parent.envelope !== undefined
          ? { parentEnvelope: parent.envelope }
          : {}),
      },
    };
    this.set({ queue: [...this.snapshot.queue, mutation] }, true);
    void this.flush();
  }

  /**
   * Queues deleting a row or a whole split group (all its rows, contiguous). `rows` are the rows
   * as last pulled; each carries its fingerprint for the server's conflict check.
   */
  enqueueDelete(sheetName: string, rows: TransactionRowDto[]): void {
    if (rows.length === 0) return;
    const base: DeleteRowBase[] = [...rows]
      .sort((a, b) => a.rowIndex - b.rowIndex)
      .map((row) => ({
        rowIndex: row.rowIndex,
        baseFingerprint: this.chainBase(sheetName, row.rowIndex, row),
        baseRow: row,
      }));
    const mutation: DeleteMutation = {
      id: this.newId(),
      sheetName,
      rowIndex: base[0].rowIndex,
      baseFingerprint: base[0].baseFingerprint,
      baseRow: base[0].baseRow,
      createdAt: this.now(),
      action: "deleteTransaction",
      payload: { rows: base },
    };
    this.set({ queue: [...this.snapshot.queue, mutation] }, true);
    void this.flush();
  }

  /**
   * Queues a new row ("New expense"). The list shows it as a pending row at the end until the
   * sheet confirms it; retries reuse its `requestId`, so the server inserts it at most once.
   */
  enqueueAdd(sheetName: string, fields: NewTransactionFields): void {
    const mutation: AddMutation = {
      id: this.newId(),
      sheetName,
      createdAt: this.now(),
      action: "addTransaction",
      payload: { requestId: this.newRequestId(), fields },
    };
    this.set({ queue: [...this.snapshot.queue, mutation] }, true);
    void this.flush();
  }

  /** Sends queued mutations one at a time, in order. Stops at the first conflict, busy, or offline. */
  flush(): Promise<void> {
    if (this.flushPromise) return this.flushPromise;
    this.flushPromise = this.run(() => this.flushLoop()).finally(() => {
      this.flushPromise = null;
    });
    return this.flushPromise;
  }

  private async flushLoop(): Promise<void> {
    this.set({ flushing: true });
    try {
      while (
        this.snapshot.queue.length > 0 &&
        this.snapshot.conflicts.length === 0
      ) {
        const head = this.snapshot.queue[0];
        this.inFlightId = head.id;
        const keepGoing = await this.sendOne(head);
        this.inFlightId = null;
        if (!keepGoing) break;
      }
    } finally {
      this.inFlightId = null;
      this.set({ flushing: false });
    }
  }

  private removeFromQueue(id: string): QueuedMutation[] {
    return this.snapshot.queue.filter((m) => m.id !== id);
  }

  /** Returns true when the loop may continue with the next mutation. */
  private async sendOne(m: QueuedMutation): Promise<boolean> {
    try {
      if (m.action === "updateTransaction") {
        const { data } = await this.options.callApi<
          TransactionRowDto | PayPeriodStateDto
        >({
          action: "updateTransaction",
          sheetName: m.sheetName,
          rowIndex: m.rowIndex,
          baseFingerprint: m.baseFingerprint,
          changes: m.payload.changes,
          ...(m.payload.changes.note !== undefined
            ? { baseNote: m.baseRow.note }
            : {}),
        });
        this.onUpdateSucceeded(m, data);
      } else if (m.action === "deleteTransaction") {
        const { data } = await this.options.callApi<PayPeriodStateDto>({
          action: "deleteTransaction",
          sheetName: m.sheetName,
          rowIndex: m.rowIndex,
          rows: m.payload.rows.map((r) => ({
            rowIndex: r.rowIndex,
            baseFingerprint: r.baseFingerprint,
          })),
        });
        this.onDeleteSucceeded(m, data);
      } else if (m.action === "addTransaction") {
        const { data } = await this.options.callApi<AddTransactionResultDto>({
          action: "addTransaction",
          sheetName: m.sheetName,
          requestId: m.payload.requestId,
          fields: m.payload.fields,
        });
        this.onAddSucceeded(m, data);
      } else {
        const { data } = await this.options.callApi<PayPeriodStateDto>({
          action: "splitTransaction",
          sheetName: m.sheetName,
          rowIndex: m.rowIndex,
          baseFingerprint: m.baseFingerprint,
          splits: m.payload.splits,
          ...(m.payload.parentItemNote !== undefined
            ? { parentItemNote: m.payload.parentItemNote }
            : {}),
          ...(m.payload.parentEnvelope !== undefined
            ? { parentEnvelope: m.payload.parentEnvelope }
            : {}),
        });
        this.onSplitSucceeded(m, data);
      }
      this.retryAttempt = 0;
      this.set({ offline: false });
      return true;
    } catch (error) {
      return this.onSendFailed(m, error);
    }
  }

  private onUpdateSucceeded(
    m: UpdateMutation,
    data: TransactionRowDto | PayPeriodStateDto,
  ): void {
    let queue = this.removeFromQueue(m.id);
    let states = this.snapshot.states;
    let all = this.snapshot.all;
    const state = states[m.sheetName];
    let newRow: TransactionRowDto | null = null;
    if (isState(data)) {
      states = { ...states, [data.sheetName]: data };
      all = patchAll(all, data.sheetName, () => data.transactions);
      newRow = data.transactions[m.rowIndex] ?? null;
    } else if (isRow(data)) {
      if (state) states = { ...states, [m.sheetName]: replaceRow(state, data) };
      all = patchAll(all, m.sheetName, (t) => replaceInList(t, data));
      newRow = data;
    }
    // Re-base later mutations that were chained on this row's previous fingerprint.
    if (newRow) {
      const fresh = newRow;
      const chained = (r: DeleteRowBase) =>
        r.rowIndex === m.rowIndex && r.baseFingerprint === m.baseFingerprint;
      queue = queue.map((q) => {
        if (q.sheetName !== m.sheetName || q.action === "addTransaction")
          return q;
        const rows = baseRowsOf(q);
        if (!rows.some(chained)) return q;
        return withBaseRows(
          q,
          rows.map((r) =>
            chained(r)
              ? {
                  rowIndex: r.rowIndex,
                  baseFingerprint: fresh.fingerprint,
                  baseRow: fresh,
                }
              : r,
          ),
        );
      });
    }
    this.set({ queue, states, all }, true);
    this.markSaved(m.sheetName, [m.rowIndex]);
  }

  /** Shows "Saved ✓" on rows the sheet just confirmed; it clears itself after a few seconds. */
  private markSaved(sheetName: string, rowIndexes: number[]): void {
    const marks: SavedMark[] = rowIndexes.map((rowIndex) => ({
      sheetName,
      rowIndex,
    }));
    this.set({ recentlySaved: [...this.snapshot.recentlySaved, ...marks] });
    setTimeout(() => {
      this.set({
        recentlySaved: this.snapshot.recentlySaved.filter(
          (s) => !marks.includes(s),
        ),
      });
    }, SAVED_MARK_MS);
  }

  private onSplitSucceeded(m: SplitMutation, data: PayPeriodStateDto): void {
    let queue = this.removeFromQueue(m.id);
    const conflicts: Conflict[] = [...this.snapshot.conflicts];
    let states = this.snapshot.states;
    let all = this.snapshot.all;
    if (isState(data)) {
      states = { ...states, [data.sheetName]: data };
      all = patchAll(all, data.sheetName, () => data.transactions);
      // Rows below the split moved down.
      const shift = m.payload.splits.length;
      queue = this.rebaseAfterShift(
        queue,
        data,
        (rowIndex) => (rowIndex > m.rowIndex ? rowIndex + shift : rowIndex),
        conflicts,
      );
    } else {
      // Unexpected shape: pull fresh state; later mutations will be conflict-checked by the server.
      void this.refresh(m.sheetName);
    }
    this.set({ queue, states, all, conflicts }, true);
    this.markSaved(
      m.sheetName,
      Array.from(
        { length: m.payload.splits.length + 1 },
        (_, i) => m.rowIndex + i,
      ),
    );
  }

  private onDeleteSucceeded(m: DeleteMutation, data: PayPeriodStateDto): void {
    let queue = this.removeFromQueue(m.id);
    const conflicts: Conflict[] = [...this.snapshot.conflicts];
    let states = this.snapshot.states;
    let all = this.snapshot.all;
    if (isState(data)) {
      states = { ...states, [data.sheetName]: data };
      all = patchAll(all, data.sheetName, () => data.transactions);
      // Deleting clears rows in place (a gap), so no row moves: only later mutations that are
      // based on a deleted row become conflicts (their row is now blank).
      const first = m.rowIndex;
      const last = m.rowIndex + m.payload.rows.length - 1;
      queue = queue.filter((q) => {
        if (
          q.sheetName !== data.sheetName ||
          q.action === "addTransaction" ||
          !baseRowsOf(q).some((r) => r.rowIndex >= first && r.rowIndex <= last)
        )
          return true;
        conflicts.push({
          id: this.newId(),
          mutation: q,
          current: data.transactions[q.rowIndex] ?? null,
        });
        return false;
      });
    } else {
      void this.refresh(m.sheetName);
    }
    this.set({ queue, states, all, conflicts }, true);
  }

  /** Accepts the state after a new row was written. No row moved, so the queue is unchanged. */
  private onAddSucceeded(m: AddMutation, data: AddTransactionResultDto): void {
    const queue = this.removeFromQueue(m.id);
    if (!isState(data)) {
      this.set({ queue }, true);
      void this.refresh(m.sheetName);
      return;
    }
    const { addedRowIndex, ...state } = data;
    this.set(
      {
        queue,
        states: { ...this.snapshot.states, [state.sheetName]: state },
        all: patchAll(
          this.snapshot.all,
          state.sheetName,
          () => state.transactions,
        ),
      },
      true,
    );
    if (typeof addedRowIndex === "number")
      this.markSaved(state.sheetName, [addedRowIndex]);
  }

  /**
   * After a split moved rows of `data.sheetName`, finds each later mutation's rows by
   * fingerprint near their expected new index. Mutations whose rows cannot be found become conflicts.
   */
  private rebaseAfterShift(
    queue: QueuedMutation[],
    data: PayPeriodStateDto,
    expectedIndex: (rowIndex: number) => number,
    conflicts: Conflict[],
  ): QueuedMutation[] {
    const rebased: QueuedMutation[] = [];
    for (const q of queue) {
      if (q.sheetName !== data.sheetName || q.action === "addTransaction") {
        rebased.push(q);
        continue;
      }
      const baseRows = baseRowsOf(q);
      const rows: DeleteRowBase[] = [];
      for (const r of baseRows) {
        const match = this.findRowByFingerprint(
          data,
          r.baseRow.fingerprint,
          expectedIndex(r.rowIndex),
        );
        if (!match) break;
        rows.push({
          rowIndex: match.rowIndex,
          baseFingerprint: match.fingerprint,
          baseRow: match,
        });
      }
      if (rows.length === baseRows.length) rebased.push(withBaseRows(q, rows));
      else
        conflicts.push({
          id: this.newId(),
          mutation: q,
          current: data.transactions[q.rowIndex] ?? null,
        });
    }
    return rebased;
  }

  /** A sheet's cached transactions: its pay-period state, else its tab in the all-sheets view. */
  private sheetTransactions(
    sheetName: string,
  ): (TransactionRowDto | null)[] | null {
    return (
      this.snapshot.states[sheetName]?.transactions ??
      this.snapshot.all?.sheets.find((x) => x.name === sheetName)
        ?.transactions ??
      null
    );
  }

  private findRowByFingerprint(
    state: PayPeriodStateDto,
    fingerprint: string,
    expectedIndex: number,
  ): TransactionRowDto | null {
    let best: TransactionRowDto | null = null;
    for (const t of state.transactions) {
      if (!t || t.fingerprint !== fingerprint) continue;
      if (
        best === null ||
        Math.abs(t.rowIndex - expectedIndex) <
          Math.abs(best.rowIndex - expectedIndex)
      ) {
        best = t;
      }
    }
    return best;
  }

  private onSendFailed(m: QueuedMutation, error: unknown): boolean {
    if (!(error instanceof ApiError)) {
      this.set({ notice: { kind: "error", message: String(error) } });
      return false;
    }
    // An add has no base row, so it cannot conflict: treat it as a permanent rejection.
    if (error.code === "conflict" && m.action !== "addTransaction") {
      const current = error.conflictCurrent ?? null;
      let states = this.snapshot.states;
      const state = states[m.sheetName];
      if (current && state) {
        states = { ...states, [m.sheetName]: replaceRow(state, current) };
      }
      const all = current
        ? patchAll(this.snapshot.all, m.sheetName, (t) =>
            replaceInList(t, current),
          )
        : this.snapshot.all;
      this.set(
        {
          queue: this.removeFromQueue(m.id),
          states,
          all,
          conflicts: [
            ...this.snapshot.conflicts,
            { id: this.newId(), mutation: m, current },
          ],
        },
        true,
      );
      // Only the first changed row comes back: pull the sheet so "Overwrite with Mine" can
      // resend the delete with every row's current fingerprint.
      if (m.action === "deleteTransaction") void this.refresh(m.sheetName);
      return false;
    }
    if (error.isTransient) {
      if (error.code !== "busy") this.set({ offline: true });
      this.scheduleRetry();
      return false;
    }
    // Permanent rejection: drop the change (reverting the optimistic edit) and tell the user.
    const notice: Notice = {
      kind: "error",
      message: `A change was rejected: ${error.message}`,
      hint:
        error.code === "unauthorized" || error.code === "not_configured"
          ? "settings"
          : undefined,
    };
    if (error.code === "unauthorized" || error.code === "not_configured") {
      // Keep the change queued; it can be retried once the credentials are fixed.
      this.set({ notice });
      return false;
    }
    this.set({ queue: this.removeFromQueue(m.id), notice }, true);
    void this.refresh(m.sheetName);
    return true;
  }

  private scheduleRetry(): void {
    if (this.retryTimer || this.retryDelays.length === 0) return;
    const delay =
      this.retryDelays[
        Math.min(this.retryAttempt, this.retryDelays.length - 1)
      ];
    this.retryAttempt++;
    this.retryTimer = setTimeout(() => {
      this.retryTimer = null;
      void this.refresh().then(() => this.flush());
    }, delay);
  }

  // ---- conflicts ---------------------------------------------------------

  /**
   * Accept Remote: drop the local change (the remote row is already shown).
   * Overwrite with Mine: resend against the *current* fingerprint (not `force`),
   * so a third concurrent change is still caught by the server.
   */
  resolveConflict(id: string, choice: "accept" | "overwrite"): void {
    const conflict = this.snapshot.conflicts.find((c) => c.id === id);
    if (!conflict) return;
    const conflicts = this.snapshot.conflicts.filter((c) => c.id !== id);
    let queue = this.snapshot.queue;
    if (choice === "overwrite" && conflict.current) {
      const current = conflict.current;
      const mutation = conflict.mutation;
      if (mutation.action === "deleteTransaction") {
        // Resend the same rows with their current fingerprints (still checked by the server).
        const transactions = this.sheetTransactions(mutation.sheetName);
        const rows: DeleteRowBase[] = [];
        for (const r of mutation.payload.rows) {
          const row =
            transactions?.[r.rowIndex] ??
            (r.rowIndex === current.rowIndex ? current : null);
          if (row)
            rows.push({
              rowIndex: row.rowIndex,
              baseFingerprint: row.fingerprint,
              baseRow: row,
            });
        }
        if (rows.length === mutation.payload.rows.length)
          queue = [
            withBaseRows({ ...mutation, id: this.newId() }, rows),
            ...queue,
          ];
      } else {
        const resend: QueuedMutation = {
          ...mutation,
          id: this.newId(),
          rowIndex: current.rowIndex,
          baseFingerprint: current.fingerprint,
          baseRow: current,
        };
        queue = [resend, ...queue];
      }
    }
    this.set({ conflicts, queue }, true);
    void this.flush();
  }

  // ---- receipts ----------------------------------------------------------

  /** Imports Gmail receipts. Online only (it is not queued). */
  syncReceipts(options?: { mark?: boolean }): Promise<void> {
    return this.run(async () => {
      this.set({ syncingReceipts: true });
      try {
        const { data } = await this.options.callApi<SyncReceiptsResultDto>({
          action: "syncReceipts",
          ...(options?.mark === false ? { mark: false } : {}),
        });
        if (isState(data.state)) this.acceptState(data.state);
        const errors =
          data.numErrors > 0 ? ` (${data.numErrors} with errors)` : "";
        this.set({
          offline: false,
          notice: {
            kind: data.numErrors > 0 ? "error" : "info",
            message: `Imported ${data.receiptInfosAdded} receipt${data.receiptInfosAdded === 1 ? "" : "s"}${errors}.${
              data.errors.length ? ` ${data.errors.join(" ")}` : ""
            }`,
          },
        });
      } catch (error) {
        if (error instanceof ApiError) {
          const noSheet = /transaction sheet/i.test(error.message);
          this.set({
            notice: {
              kind: "error",
              message: noSheet
                ? "No sheet for this pay period. Add the new pay-period sheet in the spreadsheet, then sync again."
                : error.message,
              hint: noSheet
                ? "no-sheet"
                : error.code === "invalid_config"
                  ? "variables"
                  : error.code === "unauthorized" ||
                      error.code === "not_configured"
                    ? "settings"
                    : undefined,
            },
            offline:
              error.code === "network" || error.code === "timeout"
                ? true
                : this.snapshot.offline,
          });
        } else {
          this.set({ notice: { kind: "error", message: String(error) } });
        }
      } finally {
        this.set({ syncingReceipts: false });
      }
    });
  }
}
