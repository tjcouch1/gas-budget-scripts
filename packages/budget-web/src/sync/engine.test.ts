import { describe, expect, it, vi } from "vitest";
import type {
  AllTransactionsDto,
  PayPeriodStateDto,
  TransactionRowDto,
} from "@budget/shared";
import { ApiError, type ApiRequestBody, type CallApi } from "../api/client.ts";
import { MemoryKeyValueStore } from "../crypto/kv-store.ts";
import { flush, makeRow, makeState } from "../test/fixtures.ts";
import { SyncEngine } from "./engine.ts";
import { buildViewRows } from "./optimistic.ts";
import type { RowMutation } from "./types.ts";

const SHEET = "3/3/26 - 3/16/26";

type Handler = (req: ApiRequestBody) => unknown;

/** Builds a callApi mock from a handler; thrown ApiErrors reject, other values resolve as data. */
function mockApi(handler: Handler) {
  const calls: ApiRequestBody[] = [];
  const callApi = (async (req: ApiRequestBody) => {
    calls.push(req);
    const data = await handler(req);
    return { data, warnings: [] };
  }) as CallApi;
  return { callApi, calls };
}

function newEngine(
  handler: Handler,
  store = new MemoryKeyValueStore(),
  retryDelaysMs: number[] = [],
) {
  const api = mockApi(handler);
  const engine = new SyncEngine({
    callApi: api.callApi,
    store,
    retryDelaysMs,
    now: () => 1000,
  });
  return { engine, store, ...api };
}

const rowsOf = (engine: SyncEngine) =>
  buildViewRows(engine.getSnapshot().states[SHEET], engine.getSnapshot().queue);

const twoRows = () => [
  makeRow(0, { name: "Aldi" }),
  makeRow(1, { name: "Shell", cost: 30 }),
];

describe("SyncEngine reads", () => {
  it("shows the cache immediately, then refreshes", async () => {
    const store = new MemoryKeyValueStore();
    const cached = makeState(twoRows());
    await store.set("engine", {
      currentSheet: SHEET,
      states: { [SHEET]: cached },
      queue: [],
      conflicts: [],
    });
    let release: (s: PayPeriodStateDto) => void = () => {};
    const { engine } = newEngine(
      () => new Promise<PayPeriodStateDto>((r) => (release = r)),
      store,
    );
    const init = engine.init();
    await flush();
    // Cached data is visible while the refresh is still in flight.
    expect(engine.getSnapshot().ready).toBe(true);
    expect(engine.getSnapshot().refreshing).toBe(true);
    expect(engine.getSnapshot().states[SHEET].transactions).toHaveLength(2);
    release(makeState([makeRow(0, { name: "Fresh" })]));
    await init;
    expect(engine.getSnapshot().states[SHEET].transactions).toHaveLength(1);
    expect(engine.getSnapshot().refreshing).toBe(false);
  });

  it("adopts the server sheet when none is selected and persists it", async () => {
    const { engine, store } = newEngine(() => makeState(twoRows()));
    await engine.init();
    expect(engine.getSnapshot().currentSheet).toBe(SHEET);
    await flush();
    expect(
      (await store.get<{ currentSheet: string }>("engine"))?.currentSheet,
    ).toBe(SHEET);
  });

  it("stays usable offline when the refresh fails with a network error", async () => {
    const store = new MemoryKeyValueStore();
    await store.set("engine", {
      currentSheet: SHEET,
      states: { [SHEET]: makeState(twoRows()) },
      queue: [],
      conflicts: [],
    });
    const { engine } = newEngine(() => {
      throw new ApiError("network", "down");
    }, store);
    await engine.init();
    expect(engine.getSnapshot().offline).toBe(true);
    expect(engine.getSnapshot().notice).toBeNull();
    expect(rowsOf(engine)).toHaveLength(2);
  });

  it("points to Settings on unauthorized", async () => {
    const { engine } = newEngine(() => {
      throw new ApiError("unauthorized", "bad key");
    });
    await engine.init();
    expect(engine.getSnapshot().notice).toMatchObject({ hint: "settings" });
  });

  it("clears the Settings notice once a refresh succeeds (credentials were saved)", async () => {
    let configured = false;
    const { engine } = newEngine(() => {
      if (!configured) throw new ApiError("not_configured", "no credentials");
      return makeState(twoRows());
    });
    await engine.init();
    expect(engine.getSnapshot().notice).toMatchObject({ hint: "settings" });
    configured = true;
    await engine.refresh();
    expect(engine.getSnapshot().notice).toBeNull();
  });
});

describe("SyncEngine stale pay period and note edits", () => {
  const GONE = "1/1/26 - 1/14/26";

  it("falls back to the default pay period when the cached sheet no longer exists", async () => {
    const store = new MemoryKeyValueStore();
    await store.set("engine", {
      currentSheet: GONE,
      states: { [GONE]: makeState(twoRows(), { sheetName: GONE }) },
      queue: [],
      conflicts: [],
    });
    const { engine, calls } = newEngine((req) => {
      if (req.action === "getState" && req.sheetName === GONE)
        throw new ApiError("not_found", `No transaction sheet named "${GONE}"`);
      return makeState(twoRows());
    }, store);
    await engine.init();
    const snap = engine.getSnapshot();
    expect(
      calls.map((c) => ("sheetName" in c ? c.sheetName : undefined)),
    ).toEqual([GONE, undefined]);
    expect(snap.states[GONE]).toBeUndefined();
    expect(snap.currentSheet).toBe(SHEET);
    expect(snap.states[SHEET]).toBeDefined();
    expect(snap.notice).toMatchObject({
      kind: "info",
      message: `Pay period "${GONE}" is no longer in the spreadsheet, so the current pay period is shown.`,
    });
    const saved = await store.get<{ states: object }>("engine");
    expect(Object.keys(saved!.states)).toEqual([SHEET]);
  });

  it("refresh resolves true on success and false on failure", async () => {
    let fail = false;
    const { engine } = newEngine(() => {
      if (fail) throw new ApiError("network", "down");
      return makeState(twoRows());
    });
    await engine.init();
    expect(await engine.refresh()).toBe(true);
    fail = true;
    expect(await engine.refresh()).toBe(false);
  });

  it("sends baseNote with a note change", async () => {
    const { engine, calls } = newEngine((req) =>
      req.action === "getState"
        ? makeState([makeRow(0, { name: "Aldi", note: "old" })])
        : makeRow(0, { name: "Aldi", note: "new", fingerprint: "newfp000" }),
    );
    await engine.init();
    const row = engine.getSnapshot().states[SHEET].transactions[0]!;
    engine.enqueueUpdate(SHEET, row, { note: "new" });
    await engine.flush();
    expect(calls.find((c) => c.action === "updateTransaction")).toMatchObject({
      changes: { note: "new" },
      baseNote: "old",
    });
  });

  it("does not send baseNote without a note change", async () => {
    const { engine, calls } = newEngine((req) =>
      req.action === "getState" ? makeState(twoRows()) : makeRow(0),
    );
    await engine.init();
    const row = engine.getSnapshot().states[SHEET].transactions[0]!;
    engine.enqueueUpdate(SHEET, row, { envelope: "Gas" });
    await engine.flush();
    expect(
      calls.find((c) => c.action === "updateTransaction"),
    ).not.toHaveProperty("baseNote");
  });
});

describe("SyncEngine updates", () => {
  async function setup(handler: Handler) {
    const ctx = newEngine((req) =>
      req.action === "getState" ? makeState(twoRows()) : handler(req),
    );
    await ctx.engine.init();
    return ctx;
  }

  it("applies the change optimistically, sends it with the base fingerprint, and clears the queue", async () => {
    let release: (r: TransactionRowDto) => void = () => {};
    const { engine, calls } = await setup(
      () => new Promise<TransactionRowDto>((r) => (release = r)),
    );
    const row = engine.getSnapshot().states[SHEET].transactions[0]!;
    engine.enqueueUpdate(SHEET, row, { envelope: "Groceries" });
    await flush();
    expect(rowsOf(engine)[0]).toMatchObject({
      envelope: "Groceries",
      pending: true,
    });
    expect(engine.getSnapshot().queue).toHaveLength(1);
    expect(calls.at(-1)).toMatchObject({
      action: "updateTransaction",
      sheetName: SHEET,
      rowIndex: 0,
      baseFingerprint: row.fingerprint,
      changes: { envelope: "Groceries" },
    });
    expect(calls.at(-1)).not.toHaveProperty("force");
    release({ ...row, envelope: "Groceries", fingerprint: "newfp000" });
    await engine.flush();
    await flush();
    expect(engine.getSnapshot().queue).toHaveLength(0);
    expect(engine.getSnapshot().states[SHEET].transactions[0]).toMatchObject({
      envelope: "Groceries",
      fingerprint: "newfp000",
    });
  });

  it("optimistically shows an edited cost, evaluating inline math", async () => {
    const { engine } = await setup(() => new Promise(() => {}));
    const row = engine.getSnapshot().states[SHEET].transactions[0]!;
    engine.enqueueUpdate(SHEET, row, { costExpression: "2*3+1" });
    expect(rowsOf(engine)[0]).toMatchObject({ cost: 7, costFormula: "=2*3+1" });
  });

  it("merges successive edits to the same row into one request", async () => {
    const { engine, calls } = newEngine((req) =>
      req.action === "getState" ? makeState(twoRows()) : new Promise(() => {}),
    );
    await engine.init();
    const row0 = engine.getSnapshot().states[SHEET].transactions[0]!;
    const row1 = engine.getSnapshot().states[SHEET].transactions[1]!;
    // Row 0 is in flight, so row 1 edits stay queued and can merge.
    engine.enqueueUpdate(SHEET, row0, { envelope: "Gas" });
    await flush();
    engine.enqueueUpdate(SHEET, row1, { envelope: "Gas" });
    engine.enqueueUpdate(SHEET, row1, { paidWith: "Pat Venmo" });
    const queue = engine.getSnapshot().queue;
    expect(queue).toHaveLength(2);
    expect(queue[1].payload).toEqual({
      changes: { envelope: "Gas", paidWith: "Pat Venmo" },
    });
    expect(calls.filter((c) => c.action === "updateTransaction")).toHaveLength(
      1,
    );
  });

  it("keeps changes queued while offline and sends them when back online", async () => {
    let online = false;
    const { engine, calls } = newEngine((req) => {
      if (!online) throw new ApiError("network", "offline");
      return req.action === "getState"
        ? makeState(twoRows())
        : { ...twoRows()[0], envelope: "Gas", fingerprint: "n1" };
    });
    await engine.init();
    // init refresh failed, so seed state via a successful refresh later
    online = true;
    await engine.refresh();
    online = false;
    const row = engine.getSnapshot().states[SHEET].transactions[0]!;
    engine.enqueueUpdate(SHEET, row, { envelope: "Gas" });
    await engine.flush();
    expect(engine.getSnapshot().offline).toBe(true);
    expect(engine.getSnapshot().queue).toHaveLength(1);
    expect(rowsOf(engine)[0].envelope).toBe("Gas"); // still shown optimistically
    online = true;
    engine.setOnline(true);
    await flush();
    await engine.flush();
    await flush();
    expect(engine.getSnapshot().queue).toHaveLength(0);
    expect(engine.getSnapshot().offline).toBe(false);
    expect(
      calls.filter((c) => c.action === "updateTransaction").length,
    ).toBeGreaterThanOrEqual(2);
  });

  it("persists the queue so it survives a reload", async () => {
    const store = new MemoryKeyValueStore();
    const first = newEngine((req) => {
      if (req.action === "getState") return makeState(twoRows());
      throw new ApiError("network", "offline");
    }, store);
    await first.engine.init();
    first.engine.enqueueUpdate(SHEET, twoRows()[0], { envelope: "Gas" });
    await first.engine.flush();
    await flush();
    const second = newEngine(() => {
      throw new ApiError("network", "offline");
    }, store);
    await second.engine.init();
    expect(second.engine.getSnapshot().queue).toHaveLength(1);
    expect(rowsOf(second.engine)[0].envelope).toBe("Gas");
  });

  it("retries busy responses with backoff", async () => {
    vi.useFakeTimers();
    try {
      let busy = true;
      const { engine, calls } = newEngine(
        (req) => {
          if (req.action === "getState") return makeState(twoRows());
          if (busy) throw new ApiError("busy", "locked");
          return { ...twoRows()[0], envelope: "Gas", fingerprint: "n2" };
        },
        new MemoryKeyValueStore(),
        [50],
      );
      await engine.init();
      engine.enqueueUpdate(SHEET, twoRows()[0], { envelope: "Gas" });
      await vi.advanceTimersByTimeAsync(0);
      expect(engine.getSnapshot().queue).toHaveLength(1);
      expect(engine.getSnapshot().offline).toBe(false);
      busy = false;
      await vi.advanceTimersByTimeAsync(100);
      expect(engine.getSnapshot().queue).toHaveLength(0);
      expect(
        calls.filter((c) => c.action === "updateTransaction"),
      ).toHaveLength(2);
    } finally {
      vi.useRealTimers();
    }
  });

  it("drops a permanently rejected change, reverts it, and tells the user", async () => {
    const { engine, calls } = await setup(() => {
      throw new ApiError("bad_request", "Name has a formula");
    });
    engine.enqueueUpdate(
      SHEET,
      engine.getSnapshot().states[SHEET].transactions[0]!,
      {
        name: "x",
      },
    );
    await engine.flush();
    await flush();
    expect(engine.getSnapshot().queue).toHaveLength(0);
    expect(engine.getSnapshot().notice?.message).toContain(
      "Name has a formula",
    );
    expect(rowsOf(engine)[0].name).toBe("Aldi");
    expect(
      calls.filter((c) => c.action === "getState").length,
    ).toBeGreaterThanOrEqual(2);
  });
});

describe("SyncEngine conflicts", () => {
  const remote = makeRow(0, {
    name: "Aldi",
    envelope: "Gas",
    fingerprint: "remote01",
  });

  async function conflicted() {
    const ctx = newEngine((req) => {
      if (req.action === "getState") return makeState(twoRows());
      if (
        req.action === "updateTransaction" &&
        req.baseFingerprint === "remote01"
      ) {
        return { ...remote, envelope: "Groceries", fingerprint: "after001" };
      }
      throw new ApiError("conflict", "row changed", remote);
    });
    await ctx.engine.init();
    const row = ctx.engine.getSnapshot().states[SHEET].transactions[0]!;
    ctx.engine.enqueueUpdate(SHEET, row, { envelope: "Groceries" });
    await ctx.engine.flush();
    return { ...ctx, original: row };
  }

  it("reverts the optimistic change and records original, proposed and remote values", async () => {
    const { engine, original } = await conflicted();
    const snap = engine.getSnapshot();
    expect(snap.queue).toHaveLength(0);
    expect(snap.conflicts).toHaveLength(1);
    expect(snap.conflicts[0].mutation.baseRow).toEqual(original);
    expect(snap.conflicts[0].mutation.payload).toEqual({
      changes: { envelope: "Groceries" },
    });
    expect(snap.conflicts[0].current).toEqual(remote);
    // Remote value is what the list shows now.
    expect(rowsOf(engine)[0].envelope).toBe("Gas");
  });

  it("Accept Remote drops the local change without another write", async () => {
    const { engine, calls } = await conflicted();
    const writes = () =>
      calls.filter((c) => c.action === "updateTransaction").length;
    const before = writes();
    engine.resolveConflict(engine.getSnapshot().conflicts[0].id, "accept");
    await engine.flush();
    expect(engine.getSnapshot().conflicts).toHaveLength(0);
    expect(engine.getSnapshot().queue).toHaveLength(0);
    expect(writes()).toBe(before);
    expect(rowsOf(engine)[0].envelope).toBe("Gas");
  });

  it("Overwrite with Mine resends against the current fingerprint, without force", async () => {
    const { engine, calls } = await conflicted();
    engine.resolveConflict(engine.getSnapshot().conflicts[0].id, "overwrite");
    await engine.flush();
    await flush();
    const resend = calls
      .filter((c) => c.action === "updateTransaction")
      .at(-1)!;
    expect(resend).toMatchObject({
      baseFingerprint: "remote01",
      changes: { envelope: "Groceries" },
    });
    expect(resend).not.toHaveProperty("force");
    expect(engine.getSnapshot().conflicts).toHaveLength(0);
    expect(engine.getSnapshot().queue).toHaveLength(0);
    expect(engine.getSnapshot().states[SHEET].transactions[0]?.envelope).toBe(
      "Groceries",
    );
  });

  it("surfaces a second conflict if a third change lands during Overwrite", async () => {
    const { engine } = newEngine((req) => {
      if (req.action === "getState") return makeState(twoRows());
      throw new ApiError(
        "conflict",
        "again",
        makeRow(0, { envelope: "Gas", fingerprint: "third001" }),
      );
    });
    await engine.init();
    engine.enqueueUpdate(SHEET, twoRows()[0], { envelope: "Groceries" });
    await engine.flush();
    engine.resolveConflict(engine.getSnapshot().conflicts[0].id, "overwrite");
    await engine.flush();
    await flush();
    expect(engine.getSnapshot().conflicts).toHaveLength(1);
    expect(engine.getSnapshot().conflicts[0].current?.fingerprint).toBe(
      "third001",
    );
  });

  it("pauses the queue until the conflict is resolved", async () => {
    const { engine, calls } = newEngine((req) => {
      if (req.action === "getState") return makeState(twoRows());
      if ("rowIndex" in req && req.rowIndex === 0)
        throw new ApiError("conflict", "x", remote);
      return { ...twoRows()[1], fingerprint: "ok" };
    });
    await engine.init();
    engine.enqueueUpdate(SHEET, twoRows()[0], { envelope: "Groceries" });
    await engine.flush();
    engine.enqueueUpdate(SHEET, twoRows()[1], { envelope: "Gas" });
    await engine.flush();
    expect(engine.getSnapshot().queue).toHaveLength(1);
    expect(calls.filter((c) => c.action === "updateTransaction")).toHaveLength(
      1,
    );
    engine.resolveConflict(engine.getSnapshot().conflicts[0].id, "accept");
    await engine.flush();
    await flush();
    expect(engine.getSnapshot().queue).toHaveLength(0);
  });
});

describe("SyncEngine splits", () => {
  const splits = [
    { itemNote: "milk", expression: "3", preTax: false, envelope: "Groceries" },
    { itemNote: "eggs", expression: "4", preTax: true, envelope: "Groceries" },
  ];

  it("shows placeholder children optimistically and the parent remainder", async () => {
    const { engine } = newEngine((req) =>
      req.action === "getState" ? makeState(twoRows()) : new Promise(() => {}),
    );
    await engine.init();
    engine.enqueueSplit(SHEET, twoRows()[0], splits);
    const rows = rowsOf(engine);
    expect(rows.map((r) => r.name)).toEqual([
      "Aldi | ",
      "Aldi | milk",
      "Aldi | eggs",
      "Shell",
    ]);
    expect(rows[0]).toMatchObject({
      splitPending: true,
      cost: 10 - 3 - 4 * 1.0825,
    });
    expect(rows[1].synthetic).toBe(true);
  });

  it("shows each placeholder child's tax deductible and note", async () => {
    const { engine } = newEngine((req) =>
      req.action === "getState" ? makeState(twoRows()) : new Promise(() => {}),
    );
    await engine.init();
    engine.enqueueSplit(SHEET, twoRows()[0], [
      { ...splits[0], taxDeductible: "Yes", note: "for Pat" },
      splits[1],
    ]);
    const rows = rowsOf(engine);
    expect(rows[1]).toMatchObject({
      taxDeductible: "Yes",
      note: "for Pat",
    });
    expect(rows[2]).toMatchObject({ taxDeductible: "", note: "" });
  });

  it("re-bases later queued edits by fingerprint after the rows shift", async () => {
    const shifted = makeState(
      [
        makeRow(0, { name: "Aldi | ", cost: 2.67 }),
        makeRow(1, { name: "Aldi | milk" }),
        makeRow(2, { name: "Aldi | eggs" }),
        twoRows()[1],
      ].map((r, i) => ({ ...r, rowIndex: i })),
    );
    shifted.transactions[3] = { ...twoRows()[1], rowIndex: 3 };
    let releaseSplit: (s: PayPeriodStateDto) => void = () => {};
    const { engine, calls } = newEngine((req) => {
      if (req.action === "getState") return makeState(twoRows());
      if (req.action === "splitTransaction") {
        return new Promise<PayPeriodStateDto>((r) => (releaseSplit = r));
      }
      return {
        ...twoRows()[1],
        rowIndex: 3,
        envelope: "Gas",
        fingerprint: "fin",
      };
    });
    await engine.init();
    engine.enqueueSplit(SHEET, twoRows()[0], splits);
    await flush();
    engine.enqueueUpdate(SHEET, twoRows()[1], { envelope: "Gas" });
    releaseSplit(shifted);
    await engine.flush();
    await flush();
    const update = calls.find((c) => c.action === "updateTransaction")!;
    expect(update).toMatchObject({
      rowIndex: 3,
      baseFingerprint: twoRows()[1].fingerprint,
    });
    expect(engine.getSnapshot().queue).toHaveLength(0);
  });

  it("turns an edit into a conflict when its row cannot be found after a split", async () => {
    const shifted = makeState([
      makeRow(0, { name: "Aldi | " }),
      makeRow(1, { name: "Aldi | milk" }),
      makeRow(2, { name: "Aldi | eggs" }),
      makeRow(3, { name: "Completely different", cost: 99 }),
    ]);
    let releaseSplit: (s: PayPeriodStateDto) => void = () => {};
    const { engine } = newEngine((req) => {
      if (req.action === "getState") return makeState(twoRows());
      return new Promise<PayPeriodStateDto>((r) => (releaseSplit = r));
    });
    await engine.init();
    engine.enqueueSplit(SHEET, twoRows()[0], splits);
    await flush();
    engine.enqueueUpdate(SHEET, twoRows()[1], { envelope: "Gas" });
    releaseSplit(shifted);
    await engine.flush();
    await flush();
    expect(engine.getSnapshot().queue).toHaveLength(0);
    expect(engine.getSnapshot().conflicts).toHaveLength(1);
    expect(engine.getSnapshot().conflicts[0].mutation.baseRow.name).toBe(
      "Shell",
    );
  });

  it("chains an edit queued behind an edit on the same row", async () => {
    let release: (r: TransactionRowDto) => void = () => {};
    const { engine, calls } = newEngine((req) => {
      if (req.action === "getState") return makeState(twoRows());
      if (req.action === "updateTransaction") {
        return new Promise<TransactionRowDto>((r) => (release = r));
      }
      return makeState(twoRows());
    });
    await engine.init();
    const row = twoRows()[0];
    engine.enqueueUpdate(SHEET, row, { envelope: "Gas" });
    await flush();
    engine.enqueueSplit(SHEET, row, splits);
    expect((engine.getSnapshot().queue[1] as RowMutation).baseFingerprint).toBe(
      row.fingerprint,
    );
    release({ ...row, envelope: "Gas", fingerprint: "fresh001" });
    await engine.flush();
    await flush();
    const split = calls.find((c) => c.action === "splitTransaction")!;
    expect(split).toMatchObject({ baseFingerprint: "fresh001" });
  });

  it("handles a split conflict like any other", async () => {
    const current = makeRow(0, { name: "Edited", fingerprint: "cur" });
    const { engine } = newEngine((req) => {
      if (req.action === "getState") return makeState(twoRows());
      throw new ApiError("conflict", "x", current);
    });
    await engine.init();
    engine.enqueueSplit(SHEET, twoRows()[0], splits);
    await engine.flush();
    expect(engine.getSnapshot().conflicts[0].mutation.action).toBe(
      "splitTransaction",
    );
    expect(rowsOf(engine)[0].name).toBe("Edited");
  });
});

describe("SyncEngine receipts and local data", () => {
  it("imports receipts and reports the result", async () => {
    const { engine } = newEngine((req) =>
      req.action === "syncReceipts"
        ? {
            receiptInfosAdded: 3,
            numErrors: 0,
            errors: [],
            state: makeState([makeRow(0)]),
          }
        : makeState(twoRows()),
    );
    await engine.init();
    await engine.syncReceipts();
    expect(engine.getSnapshot().notice).toMatchObject({
      kind: "info",
      message: "Imported 3 receipts.",
    });
    expect(engine.getSnapshot().states[SHEET].transactions).toHaveLength(1);
  });

  it("uses the state in the sync response instead of fetching it again", async () => {
    const { engine, calls } = newEngine((req) =>
      req.action === "syncReceipts"
        ? {
            receiptInfosAdded: 1,
            numErrors: 0,
            errors: [],
            state: makeState([makeRow(0, { name: "Imported" })]),
          }
        : makeState(twoRows()),
    );
    await engine.init();
    const before = calls.length;
    await engine.syncReceipts();
    expect(calls.slice(before).map((c) => c.action)).toEqual(["syncReceipts"]);
    expect(rowsOf(engine)[0].name).toBe("Imported");
  });

  it("reports receipts left in Gmail because their pay period has no sheet", async () => {
    const { engine } = newEngine((req) =>
      req.action === "syncReceipts"
        ? {
            receiptInfosAdded: 2,
            numErrors: 0,
            errors: [],
            receiptsWithoutSheet: 3,
            state: makeState([makeRow(0)]),
          }
        : makeState(twoRows()),
    );
    await engine.init();
    await engine.syncReceipts();
    expect(engine.getSnapshot().notice).toEqual({
      kind: "error",
      message:
        "Imported 2 receipts. 3 receipts have no sheet for their pay period and stayed in Gmail. Add the new pay-period sheet in the spreadsheet, then sync again.",
      hint: "no-sheet",
    });
    // The imported receipts are shown
    expect(engine.getSnapshot().states[SHEET].transactions).toHaveLength(1);
  });

  it("sends mark:false only when asked", async () => {
    const { engine, calls } = newEngine((req) =>
      req.action === "syncReceipts"
        ? {
            receiptInfosAdded: 0,
            numErrors: 0,
            errors: [],
            state: makeState(twoRows()),
          }
        : makeState(twoRows()),
    );
    await engine.init();
    await engine.syncReceipts();
    await engine.syncReceipts({ mark: false });
    const sent = calls.filter((c) => c.action === "syncReceipts");
    expect(sent[0]).not.toHaveProperty("mark");
    expect(sent[1]).toMatchObject({ mark: false });
  });

  it("clearLocalData also resets the selected sheet and drops in-flight results", async () => {
    let release!: () => void;
    const gate = new Promise<void>((r) => (release = r));
    const { engine } = newEngine(async (req) => {
      if (req.action === "getState") return makeState(twoRows());
      await gate;
      return makeRow(0, { name: "Aldi", note: "x", fingerprint: "fp2" });
    });
    await engine.init();
    const row = engine.getSnapshot().states[SHEET].transactions[0]!;
    engine.enqueueUpdate(SHEET, row, { note: "x" });
    const cleared = engine.clearLocalData();
    release();
    await cleared;
    const snap = engine.getSnapshot();
    expect(snap.currentSheet).toBeNull();
    expect(snap.states).toEqual({});
    expect(snap.queue).toEqual([]);
  });

  it("explains a missing pay-period sheet", async () => {
    const { engine } = newEngine((req) => {
      if (req.action === "syncReceipts") {
        throw new ApiError(
          "bad_request",
          "Could not find transaction sheet for 3/20/26",
        );
      }
      return makeState(twoRows());
    });
    await engine.init();
    await engine.syncReceipts();
    expect(engine.getSnapshot().notice).toMatchObject({
      kind: "error",
      hint: "no-sheet",
    });
  });

  it("points to Customize variables when the rules are invalid", async () => {
    const { engine } = newEngine((req) => {
      if (req.action === "syncReceipts")
        throw new ApiError(
          "invalid_config",
          "ReceiptRules variable is invalid",
        );
      return makeState(twoRows());
    });
    await engine.init();
    await engine.syncReceipts();
    expect(engine.getSnapshot().notice).toMatchObject({
      kind: "error",
      hint: "variables",
    });
  });

  it("clears cached data, the queue, and conflicts", async () => {
    const store = new MemoryKeyValueStore();
    const { engine } = newEngine((req) => {
      if (req.action === "getState") return makeState(twoRows());
      throw new ApiError("network", "down");
    }, store);
    await engine.init();
    engine.enqueueUpdate(SHEET, twoRows()[0], { envelope: "Gas" });
    await engine.flush();
    await engine.clearLocalData();
    expect(engine.getSnapshot().queue).toHaveLength(0);
    expect(engine.getSnapshot().states).toEqual({});
    expect(await store.get("engine")).toBeUndefined();
  });
});

describe("SyncEngine all pay periods", () => {
  const OLD = "2/17/26 - 3/2/26";
  const makeAll = (): AllTransactionsDto => ({
    taxMultiplier: 1.0825,
    sheets: [
      {
        name: SHEET,
        startDate: "2026-03-03",
        endDate: "2026-03-16",
        isGap: false,
        transactions: twoRows(),
        envelopeOptions: ["Groceries"],
        paidWithOptions: null,
      },
      {
        name: OLD,
        startDate: "2026-02-17",
        endDate: "2026-03-02",
        isGap: false,
        transactions: [makeRow(0, { name: "Old shop" })],
        envelopeOptions: null,
        paidWithOptions: null,
      },
    ],
  });
  const oldRows = (engine: SyncEngine) =>
    buildViewRows(
      {
        sheetName: OLD,
        taxMultiplier: 1.0825,
        transactions: engine.getSnapshot().all!.sheets[1].transactions,
      },
      engine.getSnapshot().queue,
      true,
    );

  it("loads every tab on demand only, caches it, and persists it", async () => {
    const { engine, calls, store } = newEngine((req) =>
      req.action === "getAllTransactions" ? makeAll() : makeState(twoRows()),
    );
    await engine.init();
    expect(calls.map((c) => c.action)).toEqual(["getState"]);
    await engine.loadAll();
    expect(calls.map((c) => c.action)).toEqual([
      "getState",
      "getAllTransactions",
    ]);
    expect(engine.getSnapshot().all?.sheets).toHaveLength(2);
    expect(engine.getSnapshot().loadingAll).toBe(false);
    await flush();
    expect((await store.get<{ all: unknown }>("engine"))?.all).toBeTruthy();
    // A new engine shows the cached all-tabs data before any request
    const again = newEngine(() => makeState(twoRows()), store);
    await again.engine.init();
    expect(again.engine.getSnapshot().all?.sheets).toHaveLength(2);
  });

  it("edits a row of another tab with that tab's sheetName and fingerprint", async () => {
    const all = makeAll();
    const target = all.sheets[1].transactions[0]!;
    const { engine, calls } = newEngine((req) => {
      if (req.action === "getAllTransactions") return all;
      if (req.action === "updateTransaction")
        return { ...target, envelope: "Groceries", fingerprint: "new-fp" };
      return makeState(twoRows());
    });
    await engine.init();
    await engine.loadAll();
    engine.enqueueUpdate(OLD, target, { envelope: "Groceries" });
    // Optimistic while queued
    expect(oldRows(engine)[0]).toMatchObject({
      envelope: "Groceries",
      pending: true,
      sheetName: OLD,
    });
    await engine.flush();
    expect(calls.at(-1)).toMatchObject({
      action: "updateTransaction",
      sheetName: OLD,
      rowIndex: 0,
      baseFingerprint: target.fingerprint,
    });
    expect(engine.getSnapshot().queue).toEqual([]);
    const cached = engine.getSnapshot().all!.sheets[1].transactions[0];
    expect(cached).toMatchObject({
      envelope: "Groceries",
      fingerprint: "new-fp",
    });
  });

  it("scopes keys and split group ids by sheet when rows of several tabs are shown", () => {
    const grouped = makeRow(0, { splitGroupId: "0", name: "A | " });
    const rows = [SHEET, OLD].flatMap((name) =>
      buildViewRows(
        { sheetName: name, taxMultiplier: 1, transactions: [grouped] },
        [],
        true,
      ),
    );
    expect(new Set(rows.map((r) => r.key)).size).toBe(2);
    expect(new Set(rows.map((r) => r.splitGroupId)).size).toBe(2);
  });

  it("clears the all-tabs cache with the local data", async () => {
    const { engine } = newEngine((req) =>
      req.action === "getAllTransactions" ? makeAll() : makeState(twoRows()),
    );
    await engine.init();
    await engine.loadAll();
    await engine.clearLocalData();
    expect(engine.getSnapshot().all).toBeNull();
  });
});

describe("SyncEngine split parent fields and saved marks", () => {
  const items = [
    { itemNote: "milk", expression: "3", preTax: false, envelope: "Gas" },
  ];

  it("sends the original's note and envelope with the split and shows them optimistically", async () => {
    const { engine, calls } = newEngine((req) =>
      req.action === "getState" ? makeState(twoRows()) : new Promise(() => {}),
    );
    await engine.init();
    engine.enqueueSplit(SHEET, twoRows()[0], items, {
      itemNote: "household",
      envelope: "Groceries",
    });
    await flush();
    expect(calls.find((c) => c.action === "splitTransaction")).toMatchObject({
      splits: items,
      parentItemNote: "household",
      parentEnvelope: "Groceries",
    });
    const rows = rowsOf(engine);
    expect(rows[0]).toMatchObject({
      name: "Aldi | household",
      envelope: "Groceries",
      splitPending: true,
    });
    expect(rows[1].name).toBe("Aldi | milk");
  });

  it("omits the parent fields when none were edited", async () => {
    const { engine, calls } = newEngine((req) =>
      req.action === "getState" ? makeState(twoRows()) : new Promise(() => {}),
    );
    await engine.init();
    engine.enqueueSplit(SHEET, twoRows()[0], items);
    await flush();
    const sent = calls.find((c) => c.action === "splitTransaction")!;
    expect(sent).not.toHaveProperty("parentItemNote");
    expect(sent).not.toHaveProperty("parentEnvelope");
  });

  it("marks saved rows briefly after the sheet confirms a change", async () => {
    vi.useFakeTimers();
    try {
      const saved = { ...twoRows()[0], envelope: "Gas", fingerprint: "n" };
      const { engine } = newEngine((req) =>
        req.action === "getState" ? makeState(twoRows()) : saved,
      );
      await engine.init();
      engine.enqueueUpdate(SHEET, twoRows()[0], { envelope: "Gas" });
      await vi.advanceTimersByTimeAsync(10);
      expect(engine.getSnapshot().queue).toHaveLength(0);
      expect(engine.getSnapshot().recentlySaved).toEqual([
        { sheetName: SHEET, rowIndex: 0 },
      ]);
      await vi.advanceTimersByTimeAsync(5000);
      expect(engine.getSnapshot().recentlySaved).toEqual([]);
    } finally {
      vi.useRealTimers();
    }
  });
});

describe("SyncEngine deletes", () => {
  /** Coffee (0), a split group Aldi (1-2), Shell (3) */
  const groupRows = () => [
    makeRow(0, { name: "Coffee", cost: 4 }),
    makeRow(1, { name: "Aldi | ", cost: 6, splitGroupId: "1" }),
    makeRow(2, { name: "Aldi | milk", cost: 4, splitGroupId: "1" }),
    makeRow(3, { name: "Shell", cost: 30 }),
  ];
  /** The sheet after deleting the Aldi group: rows 1-2 are blank gaps and Shell stays at row 3 */
  const afterGroupDelete = () =>
    makeState([groupRows()[0], null, null, groupRows()[3]]);

  it("hides the rows optimistically, sends every row's fingerprint, and accepts the returned state", async () => {
    let release: (s: PayPeriodStateDto) => void = () => {};
    const { engine, calls, store } = newEngine((req) => {
      if (req.action === "getState") return makeState(groupRows());
      return new Promise<PayPeriodStateDto>((r) => (release = r));
    });
    await engine.init();
    const [, parent, child] = groupRows();
    // Passed out of order: rows are sorted by index
    engine.enqueueDelete(SHEET, [child, parent]);
    expect(rowsOf(engine).map((r) => r.name)).toEqual(["Coffee", "Shell"]);
    await flush();
    // Persisted while queued (offline-capable)
    const persisted = await store.get<{ queue: { action: string }[] }>(
      "engine",
    );
    expect(persisted?.queue.map((q) => q.action)).toEqual([
      "deleteTransaction",
    ]);
    expect(calls.at(-1)).toEqual({
      action: "deleteTransaction",
      sheetName: SHEET,
      rowIndex: 1,
      rows: [
        { rowIndex: 1, baseFingerprint: parent.fingerprint },
        { rowIndex: 2, baseFingerprint: child.fingerprint },
      ],
    });
    release(afterGroupDelete());
    await engine.flush();
    await flush();
    expect(engine.getSnapshot().queue).toEqual([]);
    expect(rowsOf(engine).map((r) => [r.rowIndex, r.name])).toEqual([
      [0, "Coffee"],
      [3, "Shell"],
    ]);
  });

  it("keeps a later queued edit on a row below the delete unchanged (no rows move)", async () => {
    let release: (s: PayPeriodStateDto) => void = () => {};
    const { engine, calls } = newEngine((req) => {
      if (req.action === "getState") return makeState(groupRows());
      if (req.action === "deleteTransaction")
        return new Promise<PayPeriodStateDto>((r) => (release = r));
      return {
        ...groupRows()[3],
        rowIndex: 1,
        envelope: "Gas",
        fingerprint: "s2",
      };
    });
    await engine.init();
    engine.enqueueDelete(SHEET, groupRows().slice(1, 3));
    await flush();
    engine.enqueueUpdate(SHEET, groupRows()[3], { envelope: "Gas" });
    release(afterGroupDelete());
    await engine.flush();
    await flush();
    expect(calls.find((c) => c.action === "updateTransaction")).toMatchObject({
      rowIndex: 3,
      baseFingerprint: groupRows()[3].fingerprint,
    });
    expect(engine.getSnapshot().queue).toHaveLength(0);
    expect(engine.getSnapshot().conflicts).toHaveLength(0);
  });

  it("turns a later queued edit of a deleted row into a conflict instead of sending it", async () => {
    let release: (s: PayPeriodStateDto) => void = () => {};
    const { engine, calls } = newEngine((req) => {
      if (req.action === "getState") return makeState(groupRows());
      if (req.action === "deleteTransaction")
        return new Promise<PayPeriodStateDto>((r) => (release = r));
      throw new Error("unexpected " + req.action);
    });
    await engine.init();
    engine.enqueueDelete(SHEET, groupRows().slice(1, 3));
    await flush();
    engine.enqueueUpdate(SHEET, groupRows()[2], { envelope: "Gas" });
    release(afterGroupDelete());
    await engine.flush();
    await flush();
    expect(calls.some((c) => c.action === "updateTransaction")).toBe(false);
    const snap = engine.getSnapshot();
    expect(snap.queue).toHaveLength(0);
    expect(snap.conflicts).toHaveLength(1);
    expect(snap.conflicts[0].current).toBeNull();
    expect(snap.conflicts[0].mutation.rowIndex).toBe(2);
  });

  it("re-bases a queued delete after an earlier split and chains it on an earlier edit", async () => {
    const shell = groupRows()[3];
    let releaseUpdate: (r: TransactionRowDto) => void = () => {};
    const { engine, calls } = newEngine((req) => {
      if (req.action === "getState") return makeState(groupRows());
      if (req.action === "updateTransaction")
        return new Promise<TransactionRowDto>((r) => (releaseUpdate = r));
      if (req.action === "splitTransaction")
        // Coffee split into 2 rows: Shell moves from 3 to 4
        return makeState([
          makeRow(0, { name: "Coffee | ", splitGroupId: "0" }),
          makeRow(1, { name: "Coffee | x", splitGroupId: "0" }),
          { ...groupRows()[1], rowIndex: 2, splitGroupId: "2" },
          { ...groupRows()[2], rowIndex: 3, splitGroupId: "2" },
          { ...shell, rowIndex: 4, envelope: "Gas", fingerprint: "edited" },
        ]);
      return afterGroupDelete();
    });
    await engine.init();
    engine.enqueueUpdate(SHEET, shell, { envelope: "Gas" });
    await flush();
    engine.enqueueSplit(SHEET, groupRows()[0], [
      { itemNote: "x", expression: "1", preTax: false, envelope: "" },
    ]);
    engine.enqueueDelete(SHEET, [shell]);
    // Chained on the in-flight edit of the same row
    expect((engine.getSnapshot().queue[2] as RowMutation).baseFingerprint).toBe(
      shell.fingerprint,
    );
    releaseUpdate({ ...shell, envelope: "Gas", fingerprint: "edited" });
    await engine.flush();
    await flush();
    expect(calls.find((c) => c.action === "deleteTransaction")).toMatchObject({
      rowIndex: 4,
      rows: [{ rowIndex: 4, baseFingerprint: "edited" }],
    });
  });

  it("surfaces a conflict, restores the rows, and Accept Remote cancels the delete", async () => {
    const remote = {
      ...groupRows()[2],
      name: "Aldi | eggs",
      fingerprint: "remote",
    };
    let pulls = 0;
    const { engine, calls } = newEngine((req) => {
      // The sheet is pulled again after a delete conflict
      if (req.action === "getState")
        return pulls++ === 0
          ? makeState(groupRows())
          : makeState([groupRows()[0], groupRows()[1], remote, groupRows()[3]]);
      throw new ApiError("conflict", "changed", remote);
    });
    await engine.init();
    engine.enqueueDelete(SHEET, groupRows().slice(1, 3));
    await engine.flush();
    const snap = engine.getSnapshot();
    expect(snap.queue).toEqual([]);
    expect(snap.conflicts).toHaveLength(1);
    expect(snap.conflicts[0].mutation.action).toBe("deleteTransaction");
    expect(snap.conflicts[0].current).toEqual(remote);
    // The rows are listed again, with the remote value
    expect(rowsOf(engine).map((r) => r.name)).toEqual([
      "Coffee",
      "Aldi | ",
      "Aldi | eggs",
      "Shell",
    ]);
    const deletes = () =>
      calls.filter((c) => c.action === "deleteTransaction").length;
    const before = deletes();
    engine.resolveConflict(snap.conflicts[0].id, "accept");
    await engine.flush();
    expect(deletes()).toBe(before);
    expect(engine.getSnapshot().conflicts).toEqual([]);
  });

  it("Overwrite with Mine resends the delete with every row's current fingerprint, without force", async () => {
    const remoteChild = {
      ...groupRows()[2],
      name: "Aldi | eggs",
      fingerprint: "remote",
    };
    const remoteState = () =>
      makeState([groupRows()[0], groupRows()[1], remoteChild, groupRows()[3]]);
    let pulled = 0;
    const { engine, calls } = newEngine((req) => {
      if (req.action === "getState")
        return pulled++ === 0 ? makeState(groupRows()) : remoteState();
      const rows = (req as { rows: { baseFingerprint: string }[] }).rows;
      if (rows[1].baseFingerprint !== "remote")
        throw new ApiError("conflict", "changed", remoteChild);
      return afterGroupDelete();
    });
    await engine.init();
    engine.enqueueDelete(SHEET, groupRows().slice(1, 3));
    await engine.flush();
    await flush();
    engine.resolveConflict(engine.getSnapshot().conflicts[0].id, "overwrite");
    await engine.flush();
    await flush();
    const resend = calls
      .filter((c) => c.action === "deleteTransaction")
      .at(-1)!;
    expect(resend).toEqual({
      action: "deleteTransaction",
      sheetName: SHEET,
      rowIndex: 1,
      rows: [
        { rowIndex: 1, baseFingerprint: groupRows()[1].fingerprint },
        { rowIndex: 2, baseFingerprint: "remote" },
      ],
    });
    expect(engine.getSnapshot().conflicts).toEqual([]);
    expect(engine.getSnapshot().queue).toEqual([]);
    expect(rowsOf(engine).map((r) => r.name)).toEqual(["Coffee", "Shell"]);
  });

  it("deletes a row of another tab in the All view with that tab's sheetName", async () => {
    const OLD = "2/17/26 - 3/2/26";
    const oldRow = makeRow(0, { name: "Old shop" });
    const all: AllTransactionsDto = {
      taxMultiplier: 1.0825,
      sheets: [
        {
          name: OLD,
          startDate: "2026-02-17",
          endDate: "2026-03-02",
          isGap: false,
          transactions: [oldRow, makeRow(1, { name: "Old two" })],
          envelopeOptions: null,
          paidWithOptions: null,
        },
      ],
    };
    const { engine, calls } = newEngine((req) => {
      if (req.action === "getAllTransactions") return all;
      if (req.action === "deleteTransaction")
        return makeState([makeRow(0, { name: "Old two" })], { sheetName: OLD });
      return makeState(twoRows());
    });
    await engine.init();
    await engine.loadAll();
    engine.enqueueDelete(OLD, [oldRow]);
    const view = () =>
      buildViewRows(
        {
          sheetName: OLD,
          taxMultiplier: 1,
          transactions: engine.getSnapshot().all!.sheets[0].transactions,
        },
        engine.getSnapshot().queue,
        true,
      ).map((r) => r.name);
    expect(view()).toEqual(["Old two"]);
    await engine.flush();
    expect(calls.at(-1)).toMatchObject({
      action: "deleteTransaction",
      sheetName: OLD,
      rows: [{ rowIndex: 0, baseFingerprint: oldRow.fingerprint }],
    });
    expect(view()).toEqual(["Old two"]);
    expect(engine.getSnapshot().all!.sheets[0].transactions).toHaveLength(1);
  });

  it("ignores an empty delete", async () => {
    const { engine } = newEngine(() => makeState(twoRows()));
    await engine.init();
    engine.enqueueDelete(SHEET, []);
    expect(engine.getSnapshot().queue).toEqual([]);
  });
});

describe("SyncEngine adds", () => {
  const FIELDS = {
    date: "2026-03-07",
    name: "Farmers market",
    costExpression: "4.25*2",
    envelope: "Groceries",
    paidWith: "Pat Venmo",
    taxDeductible: "" as const,
    note: "cash",
  };
  const added = () =>
    makeRow(2, {
      name: "Farmers market",
      cost: 8.5,
      costFormula: "=4.25*2",
      envelope: "Groceries",
      paidWith: "Pat Venmo",
      date: "2026-03-07",
    });
  const afterAdd = () => ({
    ...makeState([...twoRows(), added()]),
    addedRowIndex: 2,
  });

  it("shows a pending row at the end, persists it, sends a requestId, and accepts the state", async () => {
    let release: (s: unknown) => void = () => {};
    const { engine, calls, store } = newEngine((req) => {
      if (req.action === "getState") return makeState(twoRows());
      return new Promise((r) => (release = r));
    });
    await engine.init();
    engine.enqueueAdd(SHEET, FIELDS);
    const pending = rowsOf(engine).at(-1)!;
    expect(rowsOf(engine)).toHaveLength(3);
    expect(pending).toMatchObject({
      synthetic: true,
      pending: true,
      name: "Farmers market",
      date: "2026-03-07",
      cost: 8.5,
      // A plain number; the UI formats `cost` for display
      costDisplay: "",
      envelope: "Groceries",
      paidWith: "Pat Venmo",
      note: "cash",
      splitGroupId: null,
    });
    await flush();
    const persisted = await store.get<{ queue: { action: string }[] }>(
      "engine",
    );
    expect(persisted?.queue.map((q) => q.action)).toEqual(["addTransaction"]);
    const sent = calls.at(-1) as {
      action: string;
      sheetName: string;
      requestId: string;
      fields: unknown;
    };
    expect(sent).toMatchObject({
      action: "addTransaction",
      sheetName: SHEET,
      fields: FIELDS,
    });
    expect(sent.requestId).toMatch(/^[A-Za-z0-9_-]{8,64}$/);
    release(afterAdd());
    await engine.flush();
    expect(engine.getSnapshot().queue).toEqual([]);
    expect(
      rowsOf(engine).map((r) => [r.rowIndex, r.name, !!r.synthetic]),
    ).toEqual([
      [0, "Aldi", false],
      [1, "Shell", false],
      [2, "Farmers market", false],
    ]);
    expect(engine.getSnapshot().states[SHEET]).not.toHaveProperty(
      "addedRowIndex",
    );
    expect(engine.getSnapshot().recentlySaved).toEqual([
      { sheetName: SHEET, rowIndex: 2 },
    ]);
  });

  it("stays queued offline and retries with the same requestId", async () => {
    let online = false;
    const { engine, calls } = newEngine((req) => {
      if (req.action === "getState") return makeState(twoRows());
      if (!online) throw new ApiError("network", "down");
      return afterAdd();
    });
    await engine.init();
    engine.enqueueAdd(SHEET, FIELDS);
    await engine.flush();
    expect(engine.getSnapshot().offline).toBe(true);
    expect(engine.getSnapshot().queue).toHaveLength(1);
    expect(rowsOf(engine).at(-1)).toMatchObject({ synthetic: true });
    online = true;
    engine.setOnline(true);
    await flush();
    await engine.flush();
    const adds = calls.filter((c) => c.action === "addTransaction") as {
      requestId: string;
    }[];
    expect(adds).toHaveLength(2);
    expect(adds[1].requestId).toBe(adds[0].requestId);
    expect(engine.getSnapshot().queue).toEqual([]);
  });

  it("gives each add its own requestId", async () => {
    const { engine, calls } = newEngine((req) =>
      req.action === "getState" ? makeState(twoRows()) : afterAdd(),
    );
    await engine.init();
    engine.enqueueAdd(SHEET, FIELDS);
    engine.enqueueAdd(SHEET, { ...FIELDS, name: "Second" });
    await engine.flush();
    const ids = calls
      .filter((c) => c.action === "addTransaction")
      .map((c) => (c as { requestId: string }).requestId);
    expect(ids).toHaveLength(2);
    expect(new Set(ids).size).toBe(2);
  });

  it("drops a rejected add and shows the error notice", async () => {
    const { engine } = newEngine((req) => {
      if (req.action === "getState") return makeState(twoRows());
      throw new ApiError("bad_request", "There is no empty transaction row");
    });
    await engine.init();
    engine.enqueueAdd(SHEET, FIELDS);
    await engine.flush();
    expect(engine.getSnapshot().queue).toEqual([]);
    expect(engine.getSnapshot().notice).toMatchObject({
      kind: "error",
      message: expect.stringContaining("no empty transaction row"),
    });
    expect(rowsOf(engine)).toHaveLength(2);
  });

  it("does not re-base a later queued edit (rows do not move)", async () => {
    let release: (s: unknown) => void = () => {};
    const { engine, calls } = newEngine((req) => {
      if (req.action === "getState") return makeState(twoRows());
      if (req.action === "addTransaction")
        return new Promise((r) => (release = r));
      return makeRow(1, { name: "Shell 2", cost: 30 });
    });
    await engine.init();
    const shell = twoRows()[1];
    engine.enqueueAdd(SHEET, FIELDS);
    engine.enqueueUpdate(SHEET, shell, { name: "Shell 2" });
    await flush();
    release(afterAdd());
    await engine.flush();
    expect(calls.at(-1)).toMatchObject({
      action: "updateTransaction",
      rowIndex: 1,
      baseFingerprint: shell.fingerprint,
    });
    expect(engine.getSnapshot().queue).toEqual([]);
  });

  it("keeps a queued add behind a split (never re-based or a conflict)", async () => {
    const aldi = twoRows()[0];
    const { engine, calls } = newEngine((req) => {
      if (req.action === "getState") return makeState(twoRows());
      if (req.action === "splitTransaction")
        return makeState([
          makeRow(0, { name: "Aldi | ", splitGroupId: "0" }),
          makeRow(1, { name: "Aldi | x", splitGroupId: "0" }),
          makeRow(2, { name: "Shell", cost: 30 }),
        ]);
      return afterAdd();
    });
    await engine.init();
    engine.enqueueSplit(SHEET, aldi, [
      { itemNote: "x", expression: "1", preTax: false, envelope: "" },
    ]);
    engine.enqueueAdd(SHEET, FIELDS);
    await engine.flush();
    expect(engine.getSnapshot().conflicts).toEqual([]);
    expect(calls.map((c) => c.action)).toContain("addTransaction");
    expect(engine.getSnapshot().queue).toEqual([]);
  });
});

describe("SyncEngine batched multi-period loading", () => {
  /** Ten pay-period tabs, newest first. */
  const sheetInfos = Array.from({ length: 10 }, (_, i) => {
    const day = String(28 - i * 2).padStart(2, "0");
    return {
      name: `tab ${i}`,
      startDate: `2026-01-${day}`,
      endDate: `2026-01-${day}`,
      isGap: false,
    };
  });
  const dataFor = (names: string[]): AllTransactionsDto => ({
    taxMultiplier: 1.0825,
    sheets: sheetInfos
      .filter((s) => names.includes(s.name))
      .map((s) => ({
        ...s,
        transactions: [makeRow(0, { name: `row of ${s.name}` })],
        envelopeOptions: null,
        paidWithOptions: null,
      })),
  });
  const handlerFor =
    (onAll: (names: string[]) => unknown = dataFor) =>
    (req: ApiRequestBody) =>
      req.action === "getAllTransactions"
        ? onAll((req as { sheetNames?: string[] }).sheetNames ?? [])
        : makeState([makeRow(0)], { sheets: sheetInfos });

  it("loads the tabs in small requests, showing each batch as it arrives", async () => {
    const seen: number[] = [];
    const { engine, calls } = newEngine(handlerFor());
    await engine.init();
    engine.subscribe(() =>
      seen.push(engine.getSnapshot().all?.sheets.length ?? 0),
    );
    await engine.loadAll();
    const loads = calls.filter((c) => c.action === "getAllTransactions");
    expect(
      loads.map((c) => (c as { sheetNames: string[] }).sheetNames.length),
    ).toEqual([4, 4, 2]);
    const all = engine.getSnapshot().all!;
    expect(all.sheets.map((s) => s.name)).toEqual(
      sheetInfos.map((s) => s.name),
    );
    // Rows showed up batch by batch, not all at once at the end
    expect(seen).toContain(4);
    expect(seen).toContain(8);
    expect(engine.getSnapshot()).toMatchObject({
      loadingAll: false,
      loadAllFailed: false,
      offline: false,
    });
  });

  it("loads only the requested tabs", async () => {
    const { engine, calls } = newEngine(handlerFor());
    await engine.init();
    await engine.loadAll(["tab 1", "tab 2"]);
    const loads = calls.filter((c) => c.action === "getAllTransactions");
    expect(loads).toHaveLength(1);
    expect(loads[0]).toMatchObject({ sheetNames: ["tab 1", "tab 2"] });
    expect(engine.getSnapshot().all!.sheets.map((s) => s.name)).toEqual([
      "tab 1",
      "tab 2",
    ]);
  });

  it("keeps the batches that arrived and does not go offline on a timeout after progress", async () => {
    let n = 0;
    const { engine } = newEngine(
      handlerFor((names) => {
        if (++n === 2) throw new ApiError("timeout", "The request timed out.");
        return dataFor(names);
      }),
    );
    await engine.init();
    await engine.loadAll();
    const snap = engine.getSnapshot();
    expect(snap.all!.sheets).toHaveLength(4);
    expect(snap.loadAllFailed).toBe(true);
    expect(snap.loadingAll).toBe(false);
    expect(snap.offline).toBe(false);
  });

  it("does not mark offline for a timeout of the first batch either, but does for no network", async () => {
    const timeout = newEngine(
      handlerFor(() => {
        throw new ApiError("timeout", "The request timed out.");
      }),
    );
    await timeout.engine.init();
    await timeout.engine.loadAll();
    expect(timeout.engine.getSnapshot()).toMatchObject({
      loadAllFailed: true,
      offline: false,
    });

    const network = newEngine(
      handlerFor(() => {
        throw new ApiError("network", "Network error.");
      }),
    );
    await network.engine.init();
    await network.engine.loadAll();
    expect(network.engine.getSnapshot()).toMatchObject({
      loadAllFailed: true,
      offline: true,
    });
  });

  it("lets a newer load replace an older one", async () => {
    const { engine, calls } = newEngine(handlerFor());
    await engine.init();
    const first = engine.loadAll();
    const second = engine.loadAll(["tab 0"]);
    await Promise.all([first, second]);
    const loads = calls.filter((c) => c.action === "getAllTransactions");
    // The first load stopped before its first request ran; only the second one read anything
    expect(loads).toHaveLength(1);
    expect(loads[0]).toMatchObject({ sheetNames: ["tab 0"] });
  });
});

describe("SyncEngine split groups", () => {
  /** Aldi, $30 in total, split into milk and soap. */
  const groupRows = () => [
    makeRow(0, {
      name: "Aldi | ",
      cost: 18,
      costFormula: "=30-R[1]C[0]-R[2]C[0]",
      splitGroupId: "0",
    }),
    makeRow(1, { name: "Aldi | milk", cost: 10.83, splitGroupId: "0" }),
    makeRow(2, { name: "Aldi | soap", cost: 1.17, splitGroupId: "0" }),
    makeRow(3, { name: "Shell", cost: 30 }),
  ];
  const pendingForever = (req: ApiRequestBody) =>
    req.action === "getState" ? makeState(groupRows()) : new Promise(() => {});
  const groupTotal = (rows: { cost: number | null }[]) =>
    rows.slice(0, 3).reduce((s, r) => s + (r.cost ?? 0), 0);

  it("keeps the parent's remainder up to date when a child's cost is edited", async () => {
    const { engine } = newEngine(pendingForever);
    await engine.init();
    const soap = engine.getSnapshot().states[SHEET].transactions[2]!;
    engine.enqueueUpdate(SHEET, soap, { costExpression: "5" });
    const rows = rowsOf(engine);
    expect(rows[2].cost).toBe(5);
    expect(rows[0]).toMatchObject({ cost: 14.17, pending: true });
    expect(groupTotal(rows)).toBeCloseTo(30, 9);
    // Other rows are untouched
    expect(rows[3].cost).toBe(30);
  });

  it("moves the remainder by a queued change of the split total", async () => {
    const { engine } = newEngine(pendingForever);
    await engine.init();
    const parent = engine.getSnapshot().states[SHEET].transactions[0]!;
    engine.enqueueUpdate(SHEET, parent, { splitTotal: "40" });
    const rows = rowsOf(engine);
    expect(rows[0]).toMatchObject({
      cost: 28,
      costFormula: "=40-R[1]C[0]-R[2]C[0]",
    });
    expect(groupTotal(rows)).toBeCloseTo(40, 9);
  });

  it("moves the remainder by a queued expression total", async () => {
    const { engine } = newEngine(pendingForever);
    await engine.init();
    const parent = engine.getSnapshot().states[SHEET].transactions[0]!;
    engine.enqueueUpdate(SHEET, parent, { splitTotal: "35+5" });
    const rows = rowsOf(engine);
    expect(rows[0]).toMatchObject({
      cost: 28,
      costFormula: "=(35+5)-R[1]C[0]-R[2]C[0]",
    });
    expect(groupTotal(rows)).toBeCloseTo(40, 9);
  });

  it("shows a renamed split's new prefix on every child", async () => {
    const { engine } = newEngine(pendingForever);
    await engine.init();
    const parent = engine.getSnapshot().states[SHEET].transactions[0]!;
    engine.enqueueUpdate(SHEET, parent, { name: "Lidl | " });
    expect(rowsOf(engine).map((r) => r.name)).toEqual([
      "Lidl | ",
      "Lidl | milk",
      "Lidl | soap",
      "Shell",
    ]);
  });

  it("takes the whole state back from a child's cost edit, so the parent is fresh after sync", async () => {
    const after = groupRows();
    after[0] = { ...after[0], cost: 14.17, fingerprint: after[0].fingerprint };
    after[2] = makeRow(2, { name: "Aldi | soap", cost: 5, splitGroupId: "0" });
    const { engine } = newEngine((req) =>
      req.action === "getState" ? makeState(groupRows()) : makeState(after),
    );
    await engine.init();
    const soap = engine.getSnapshot().states[SHEET].transactions[2]!;
    engine.enqueueUpdate(SHEET, soap, { costExpression: "5" });
    await engine.flush();
    expect(engine.getSnapshot().queue).toHaveLength(0);
    const rows = rowsOf(engine);
    expect(rows[0].cost).toBe(14.17);
    expect(rows[0].pending).toBeUndefined();
    expect(groupTotal(rows)).toBeCloseTo(30, 9);
  });

  it("re-bases a queued child change after the parent's rename renamed it", async () => {
    const renamed = groupRows().map((r) =>
      r.splitGroupId === "0"
        ? {
            ...r,
            name: r.name.replace("Aldi", "Lidl"),
            fingerprint: `new${r.rowIndex}`,
          }
        : r,
    );
    const { engine, calls } = newEngine((req) => {
      if (req.action !== "updateTransaction") return makeState(groupRows());
      if (req.rowIndex === 0) return makeState(renamed);
      if (req.baseFingerprint !== "new1")
        throw new ApiError("conflict", "row changed", renamed[1]);
      return { ...renamed[1], envelope: "Gas", fingerprint: "after1" };
    });
    await engine.init();
    const [parent, milk] = engine.getSnapshot().states[SHEET].transactions;
    engine.enqueueUpdate(SHEET, parent!, { name: "Lidl | " });
    engine.enqueueUpdate(SHEET, milk!, { envelope: "Gas" });
    await engine.flush();
    await flush();
    expect(engine.getSnapshot().conflicts).toHaveLength(0);
    expect(engine.getSnapshot().queue).toHaveLength(0);
    expect(
      calls.filter((c) => c.action === "updateTransaction").at(-1),
    ).toMatchObject({
      rowIndex: 1,
      baseFingerprint: "new1",
    });
  });
});

describe("SyncEngine replayed changes", () => {
  it("accepts a conflict silently when the sheet already has the change", async () => {
    const base = twoRows()[0];
    const current = {
      ...base,
      envelope: "Groceries",
      cost: 12.5,
      costFormula: null,
      costDisplay: "$12.50",
      fingerprint: "remote02",
    };
    const { engine, calls } = newEngine((req) => {
      if (req.action === "getState") return makeState(twoRows());
      throw new ApiError("conflict", "row changed", current);
    });
    await engine.init();
    engine.enqueueUpdate(
      SHEET,
      engine.getSnapshot().states[SHEET].transactions[0]!,
      {
        envelope: "Groceries",
        costExpression: "12.50",
      },
    );
    await engine.flush();
    await flush();
    expect(engine.getSnapshot().conflicts).toHaveLength(0);
    expect(engine.getSnapshot().queue).toHaveLength(0);
    expect(engine.getSnapshot().states[SHEET].transactions[0]).toEqual(current);
    expect(calls.filter((c) => c.action === "updateTransaction")).toHaveLength(
      1,
    );
  });

  it("still asks when the remote value differs from the local change", async () => {
    const current = {
      ...twoRows()[0],
      envelope: "Gas",
      fingerprint: "remote03",
    };
    const { engine } = newEngine((req) => {
      if (req.action === "getState") return makeState(twoRows());
      throw new ApiError("conflict", "row changed", current);
    });
    await engine.init();
    engine.enqueueUpdate(
      SHEET,
      engine.getSnapshot().states[SHEET].transactions[0]!,
      {
        envelope: "Groceries",
      },
    );
    await engine.flush();
    expect(engine.getSnapshot().conflicts).toHaveLength(1);
  });
  /** A replayed cost change against `remote` (the sheet's row): accepted silently or a conflict. */
  async function replayCost(
    costExpression: string,
    remote: Partial<TransactionRowDto>,
  ) {
    const current = { ...twoRows()[0], ...remote, fingerprint: "remote04" };
    const { engine } = newEngine((req) => {
      if (req.action === "getState") return makeState(twoRows());
      throw new ApiError("conflict", "row changed", current);
    });
    await engine.init();
    engine.enqueueUpdate(
      SHEET,
      engine.getSnapshot().states[SHEET].transactions[0]!,
      { costExpression },
    );
    await engine.flush();
    await flush();
    return engine.getSnapshot().conflicts.length;
  }

  it("compares formula text, not values, when either side is a formula", async () => {
    // Same value, but a formula is not a plain number (and the other way around)
    expect(await replayCost("1*1.09", { cost: 1.09, costFormula: null })).toBe(
      1,
    );
    expect(
      await replayCost("1.09", { cost: 1.09, costFormula: "=1*1.09" }),
    ).toBe(1);
    expect(await replayCost("2+1", { cost: 3, costFormula: "=1+2" })).toBe(1);
    // The same formula, spelled as the sheet returns it
    expect(
      await replayCost("1*1.09", { cost: 1.09, costFormula: "= 1 * 1.09" }),
    ).toBe(0);
    expect(await replayCost(".5*2", { cost: 1, costFormula: "=0.5*2" })).toBe(
      0,
    );
    // Plain numbers compare by value
    expect(await replayCost("1.090", { cost: 1.09, costFormula: null })).toBe(
      0,
    );
  });
});

describe("SyncEngine split group saves (updateSplitGroup)", () => {
  /** Aldi, $30 in total, split into milk and soap, then an unrelated row. */
  const groupRows = () => [
    makeRow(0, {
      name: "Aldi | ",
      cost: 18,
      costFormula: "=30-R[1]C[0]-R[2]C[0]",
      splitGroupId: "0",
    }),
    makeRow(1, { name: "Aldi | milk", cost: 10.83, splitGroupId: "0" }),
    makeRow(2, {
      name: "Aldi | soap",
      cost: 1.17,
      splitGroupId: "0",
      note: "old note",
    }),
    makeRow(3, { name: "Shell", cost: 30 }),
  ];
  const groupOf = (engine: SyncEngine) =>
    engine
      .getSnapshot()
      .states[SHEET].transactions.slice(0, 3) as TransactionRowDto[];
  /** The rows the sheet holds after the edit below (title Lidl, soap $5, total $40). */
  const savedRows = () => [
    makeRow(0, {
      name: "Lidl | ",
      cost: 24.17,
      costFormula: "=40-R[1]C[0]-R[2]C[0]",
      splitGroupId: "0",
    }),
    makeRow(1, { name: "Lidl | milk", cost: 10.83, splitGroupId: "0" }),
    makeRow(2, {
      name: "Lidl | soap",
      cost: 5,
      splitGroupId: "0",
      note: "old note",
    }),
  ];
  const edit = {
    rowChanges: [undefined, undefined, { costExpression: "5" }],
    title: "Lidl",
    splitTotal: "40",
  };

  it("queues one request with every row's fingerprint and note and shows the whole group optimistically", async () => {
    const { engine, calls } = newEngine((req) =>
      req.action === "getState"
        ? makeState(groupRows())
        : new Promise(() => {}),
    );
    await engine.init();
    engine.enqueueGroupUpdate(SHEET, groupOf(engine), edit);
    expect(engine.getSnapshot().queue).toHaveLength(1);
    const rows = rowsOf(engine);
    expect(rows.map((r) => [r.name, r.cost])).toEqual([
      ["Lidl | ", 24.17],
      ["Lidl | milk", 10.83],
      ["Lidl | soap", 5],
      ["Shell", 30],
    ]);
    expect(rows.slice(0, 3).every((r) => r.pending)).toBe(true);
    expect(rows[3].pending).toBeUndefined();
    await flush();
    const sent = calls.filter((c) => c.action !== "getState");
    expect(sent).toEqual([
      {
        action: "updateSplitGroup",
        sheetName: SHEET,
        rows: groupRows()
          .slice(0, 3)
          .map((r, i) => ({
            rowIndex: r.rowIndex,
            baseFingerprint: r.fingerprint,
            baseNote: r.note,
            ...(i === 2 ? { changes: { costExpression: "5" } } : {}),
          })),
        title: "Lidl",
        splitTotal: "40",
      },
    ]);
  });

  it("accepts the returned group rows without a whole-state response", async () => {
    const { engine } = newEngine((req) =>
      req.action === "getState"
        ? makeState(groupRows())
        : { rows: savedRows() },
    );
    await engine.init();
    engine.enqueueGroupUpdate(SHEET, groupOf(engine), edit);
    await engine.flush();
    const snap = engine.getSnapshot();
    expect(snap.queue).toHaveLength(0);
    expect(snap.states[SHEET].transactions).toEqual([
      ...savedRows(),
      groupRows()[3],
    ]);
    expect(snap.recentlySaved.map((s) => s.rowIndex)).toEqual([0, 1, 2]);
  });

  it("adds new items in the same request: placeholders first, then the returned state", async () => {
    const after = makeState([
      ...groupRows().slice(0, 3),
      makeRow(3, { name: "Aldi | jam", cost: 4, splitGroupId: "0" }),
      makeRow(4, { name: "Shell", cost: 30 }),
    ]);
    let release: (v: unknown) => void = () => {};
    const { engine, calls } = newEngine((req) =>
      req.action === "getState"
        ? makeState(groupRows())
        : new Promise((r) => (release = r)),
    );
    await engine.init();
    const newItems = [
      { itemNote: "jam", expression: "4", preTax: false, envelope: "" },
    ];
    engine.enqueueGroupUpdate(SHEET, groupOf(engine), {
      rowChanges: [],
      newItems,
    });
    const rows = rowsOf(engine);
    expect(rows.map((r) => [r.name, r.cost, !!r.synthetic])).toEqual([
      ["Aldi | ", 14, false],
      ["Aldi | milk", 10.83, false],
      ["Aldi | soap", 1.17, false],
      ["Aldi | jam", 4, true],
      ["Shell", 30, false],
    ]);
    await flush();
    expect(calls.at(-1)).toMatchObject({
      action: "updateSplitGroup",
      newItems,
    });
    release({ rows: after.transactions.slice(0, 4), state: after });
    await engine.flush();
    expect(engine.getSnapshot().states[SHEET]).toEqual(after);
  });

  it("on conflict shows the group's current rows; Accept Remote drops the whole save", async () => {
    const current = groupRows().slice(0, 3);
    current[1] = { ...current[1], envelope: "Gas", fingerprint: "remote01" };
    const { engine, calls } = newEngine((req) => {
      if (req.action === "getState") return makeState(groupRows());
      throw new ApiError(
        "conflict",
        "split changed",
        current[1],
        undefined,
        current,
      );
    });
    await engine.init();
    engine.enqueueGroupUpdate(SHEET, groupOf(engine), edit);
    await engine.flush();
    const snap = engine.getSnapshot();
    expect(snap.queue).toHaveLength(0);
    expect(snap.conflicts).toHaveLength(1);
    expect(snap.conflicts[0].currentRows).toEqual(current);
    expect(snap.conflicts[0].mutation.action).toBe("updateSplitGroup");
    // The list shows the remote rows, without the discarded edit
    expect(rowsOf(engine).map((r) => [r.name, r.envelope])).toEqual([
      ["Aldi | ", ""],
      ["Aldi | milk", "Gas"],
      ["Aldi | soap", ""],
      ["Shell", ""],
    ]);
    const writes = calls.length;
    engine.resolveConflict(snap.conflicts[0].id, "accept");
    await engine.flush();
    expect(engine.getSnapshot().conflicts).toHaveLength(0);
    expect(engine.getSnapshot().queue).toHaveLength(0);
    expect(calls).toHaveLength(writes);
  });

  it("Overwrite with Mine resends the whole group with every row's current fingerprint", async () => {
    const current = groupRows().slice(0, 3);
    current[1] = { ...current[1], envelope: "Gas", fingerprint: "remote01" };
    let conflicted = false;
    const { engine, calls } = newEngine((req) => {
      if (req.action === "getState") return makeState(groupRows());
      if (!conflicted) {
        conflicted = true;
        throw new ApiError("conflict", "x", current[1], undefined, current);
      }
      return { rows: savedRows() };
    });
    await engine.init();
    engine.enqueueGroupUpdate(SHEET, groupOf(engine), edit);
    await engine.flush();
    engine.resolveConflict(engine.getSnapshot().conflicts[0].id, "overwrite");
    await engine.flush();
    await flush();
    const resend = calls.filter((c) => c.action === "updateSplitGroup").at(-1);
    expect(resend).toMatchObject({
      title: "Lidl",
      splitTotal: "40",
      rows: [
        { rowIndex: 0, baseFingerprint: current[0].fingerprint },
        { rowIndex: 1, baseFingerprint: "remote01" },
        {
          rowIndex: 2,
          baseFingerprint: current[2].fingerprint,
          changes: { costExpression: "5" },
        },
      ],
    });
    expect(resend).not.toHaveProperty("force");
    expect(engine.getSnapshot().conflicts).toHaveLength(0);
    expect(engine.getSnapshot().queue).toHaveLength(0);
  });

  it("accepts a conflict silently when the sheet already has every row's change", async () => {
    const current = savedRows().map((r) => ({
      ...r,
      fingerprint: `x${r.rowIndex}`,
    }));
    const { engine, calls } = newEngine((req) => {
      if (req.action === "getState") return makeState(groupRows());
      throw new ApiError("conflict", "x", current[0], undefined, current);
    });
    await engine.init();
    engine.enqueueGroupUpdate(SHEET, groupOf(engine), edit);
    await engine.flush();
    expect(engine.getSnapshot().conflicts).toHaveLength(0);
    expect(engine.getSnapshot().queue).toHaveLength(0);
    expect(engine.getSnapshot().states[SHEET].transactions.slice(0, 3)).toEqual(
      current,
    );
    expect(calls.filter((c) => c.action === "updateSplitGroup")).toHaveLength(
      1,
    );
  });

  it("compares the parent's total formula text when accepting a replayed group save", async () => {
    const replay = async (splitTotal: string, remoteFormula: string) => {
      const current = savedRows().map((r, i) => ({
        ...r,
        ...(i === 0 ? { costFormula: remoteFormula } : {}),
        fingerprint: `x${r.rowIndex}`,
      }));
      const { engine } = newEngine((req) => {
        if (req.action === "getState") return makeState(groupRows());
        throw new ApiError("conflict", "x", current[0], undefined, current);
      });
      await engine.init();
      engine.enqueueGroupUpdate(SHEET, groupOf(engine), {
        ...edit,
        splitTotal,
      });
      await engine.flush();
      await flush();
      return engine.getSnapshot().conflicts.length;
    };
    expect(await replay("35+5", "=(35+5)-R[1]C[0]-R[2]C[0]")).toBe(0);
    // Same total value, different formula: not this save
    expect(await replay("35+5", "=40-R[1]C[0]-R[2]C[0]")).toBe(1);
    expect(await replay("40", "=(35+5)-R[1]C[0]-R[2]C[0]")).toBe(1);
    // Real Sheets returns the remainder in A1 notation: same total and references by value
    expect(await replay("40", "=40-C12-C13")).toBe(0);
    expect(await replay("40", "=40-C12")).toBe(1);
  });

  it("re-bases a later queued edit of a group row onto the saved row", async () => {
    const saved = savedRows();
    const { engine, calls } = newEngine((req) => {
      if (req.action === "getState") return makeState(groupRows());
      if (req.action === "updateSplitGroup") return { rows: saved };
      if (
        req.action === "updateTransaction" &&
        req.baseFingerprint !== saved[1].fingerprint
      )
        throw new ApiError("conflict", "stale", saved[1]);
      return { ...saved[1], envelope: "Gas", fingerprint: "after1" };
    });
    await engine.init();
    const rows = groupOf(engine);
    engine.enqueueGroupUpdate(SHEET, rows, edit);
    engine.enqueueUpdate(SHEET, rows[1], { envelope: "Gas" });
    await engine.flush();
    await flush();
    expect(engine.getSnapshot().conflicts).toHaveLength(0);
    expect(calls.at(-1)).toMatchObject({
      action: "updateTransaction",
      rowIndex: 1,
      baseFingerprint: saved[1].fingerprint,
    });
  });
});
