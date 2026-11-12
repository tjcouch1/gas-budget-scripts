import { describe, expect, it, vi } from "vitest";
import type { PayPeriodStateDto, TransactionRowDto } from "@budget/shared";
import { ApiError, type ApiRequestBody, type CallApi } from "../api/client.ts";
import { MemoryKeyValueStore } from "../crypto/kv-store.ts";
import { flush, makeRow, makeState } from "../test/fixtures.ts";
import { SyncEngine } from "./engine.ts";
import { buildViewRows } from "./optimistic.ts";

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
    expect(engine.getSnapshot().queue[1].baseFingerprint).toBe(row.fingerprint);
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
