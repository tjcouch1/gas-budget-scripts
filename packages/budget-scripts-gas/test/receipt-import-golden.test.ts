/* eslint-disable @typescript-eslint/no-explicit-any */
/**
 * Golden characterization tests for receipt import: they pin the exact end state of the budget
 * sheet (every non-empty cell's value, formula, note, and background) and of Gmail (labels and
 * archive state of every thread) after importing a large mixed inbox, plus the web API responses
 * built from it. Performance changes to the import must keep these snapshots unchanged.
 */
import { afterEach, beforeEach, describe, expect, it, vi } from "vitest";
import {
  addTransactionSheet,
  addVariablesSheet,
  getEnv,
  installGasGlobals,
  makeThreadWithId,
  MockTextOutput,
  type GasEnv,
} from "./gas-mocks";
import {
  chasePurchaseEmail,
  receiptEdgeFixtures,
  receiptFixtures,
  receiptThreadFixtures,
  type EmailFixture,
} from "./fixtures/receipts";

const loadBudgeting = () => import("../src/util/budgeting.util");
const loadIndex = () => import("../src/index");

const KEY = "test-api-key-0123456789abcdef";
const SHEET_A = "2/18/24 - 3/2/24";
const SHEET_B = "3/3/24 - 3/16/24";

/** Dates in sheet A / sheet B. `i % 5` hours apart so some receipts share a date (sort ties) */
function scenarioDate(i: number) {
  const base = i % 2 === 0 ? "2024-02-20T12:00:00Z" : "2024-03-05T12:00:00Z";
  return new Date(new Date(base).getTime() + (i % 5) * 60 * 60 * 1000);
}

/** Every single-email fixture as its own thread, plus multi-message and failure threads */
function buildScenarioThreads(): any[] {
  const singles = Object.entries({
    ...receiptFixtures,
    ...receiptEdgeFixtures,
  }).map(([name, fixture], i) =>
    makeThreadWithId(`t-${name}`, {
      ...(fixture.email as EmailFixture),
      date: scenarioDate(i),
    }),
  );
  const multis = Object.entries(receiptThreadFixtures).map(
    ([name, messages], i) =>
      makeThreadWithId(
        `t-thread-${name}`,
        ...messages.map((m, j) => ({ ...m, date: scenarioDate(i + j) })),
      ),
  );
  const errorAndNote = makeThreadWithId(
    "t-error-and-note",
    chasePurchaseEmail("3.21", "Combo Shop", "2024-02-22T15:00:00Z"),
    {
      from: "Chase <no.reply.alerts@chase.com>",
      subject: "Promo",
      body: "hi",
      date: new Date("2024-02-22T16:00:00Z"),
    },
    {
      from: "z at example.org",
      subject: "z",
      body: "",
      date: new Date("2024-02-22T17:00:00Z"),
    },
  );
  const failing = makeThreadWithId(
    "t-thread-fails",
    chasePurchaseEmail("9.99", "Never Read", "2024-03-06T12:00:00Z"),
  );
  failing.messagesError = new Error("Gmail thread unavailable");
  const acrossSheets = makeThreadWithId(
    "t-across-sheets",
    chasePurchaseEmail("1.11", "Early Half", "2024-02-28T12:00:00Z"),
    chasePurchaseEmail("2.22", "Late Half", "2024-03-04T12:00:00Z"),
  );
  return [...singles, ...multis, errorAndNote, failing, acrossSheets];
}

/** Sheets with existing content, stale formatting in empty rows, and banding-like fills */
function setupScenarioSheets(env: GasEnv) {
  env.spreadsheet.addSheet("Main");
  addVariablesSheet(env);
  const a = addTransactionSheet(env, SHEET_A, { max: 30 });
  a.seed("A5", [
    [new Date("2024-02-19T12:00:00Z"), "Existing", 1, "", "Food", "Credit"],
  ]);
  // Row 6 has no values (so receipts start there) but keeps a note and fills from earlier use
  a.getRange("B6").setNote("stale note");
  a.getRange("B6:C6").setBackground("#00ff00");
  a.getRange("E6").setBackground("#abcdef");
  a.getRange("B8").setNote("another stale note");
  const b = addTransactionSheet(env, SHEET_B, { max: 30 });
  b.getRange("A5:F12").setBackground("#f3f3f3");
  return { a, b };
}

function gmailState(threads: any[]) {
  return threads.map((t) => ({
    id: t.getId(),
    labels: t.labels.map((l: any) => l.getName()),
    archived: t.archived,
  }));
}

async function api(action: string, fields: Record<string, unknown> = {}) {
  const { doPost } = await loadIndex();
  const output = doPost({
    postData: {
      contents: JSON.stringify({
        apiVersion: 1,
        apiKey: KEY,
        action,
        ...fields,
      }),
    },
  } as unknown as GoogleAppsScript.Events.DoPost) as unknown as MockTextOutput;
  return JSON.parse(output.getContent());
}

let env: GasEnv;
let threads: any[];

beforeEach(() => {
  env = getEnv();
  env.PropertiesService.scriptProperties.setProperty("API_KEY", KEY);
  setupScenarioSheets(env);
  threads = buildScenarioThreads();
  env.GmailApp.threads = threads;
});

afterEach(() => {
  vi.useRealTimers();
});

describe("receipt import golden end state", () => {
  it("recordReceiptsCore with marking: sheet cells and Gmail labels/archive", async () => {
    const { getChaseReceipts, recordReceiptsCore } = await loadBudgeting();
    const { skippedReceipts, ...result } = recordReceiptsCore(
      getChaseReceipts(null, null),
      true,
    );
    // Added with GitHub #1 (no receipt here lacks a sheet); the snapshot predates the field
    expect(skippedReceipts).toEqual([]);
    expect(result).toMatchSnapshot("result");
    expect(env.spreadsheet.dumpSheets()).toMatchSnapshot("sheets");
    expect(gmailState(threads)).toMatchSnapshot("gmail");
    expect(env.GmailApp.createdLabels).toMatchSnapshot("created labels");
  });

  it("menu import without marking: same sheet, Gmail untouched, one alert", async () => {
    const { getAndRecordReceipts } = await loadBudgeting();
    const result = getAndRecordReceipts(null, null, false);
    expect(result).toMatchSnapshot("result");
    expect(env.spreadsheet.dumpSheets()).toMatchSnapshot("sheets");
    expect(gmailState(threads)).toMatchSnapshot("gmail");
    expect(env.GmailApp.createdLabels).toEqual([]);
    expect(env.SpreadsheetApp.ui.alerts).toMatchSnapshot("alerts");
  });

  it("syncReceipts (25 threads) and getAllTransactions API responses", async () => {
    vi.useFakeTimers();
    vi.setSystemTime(new Date("2024-03-05T18:00:00Z"));
    const sync = await api("syncReceipts");
    // Added with GitHub #1 (no receipt here lacks a sheet); the snapshot predates the field
    expect(sync.data.receiptsWithoutSheet).toBe(0);
    delete sync.data.receiptsWithoutSheet;
    expect(sync).toMatchSnapshot("syncReceipts response");
    expect(env.spreadsheet.dumpSheets()).toMatchSnapshot("sheets");
    expect(gmailState(threads)).toMatchSnapshot("gmail");
    const all = await api("getAllTransactions");
    expect(all).toMatchSnapshot("getAllTransactions response");
  });

  it("fetchReceiptEmails response (read-only)", async () => {
    const res = await api("fetchReceiptEmails", {
      maxThreads: 50,
      includeHtml: true,
    });
    expect(res).toMatchSnapshot("fetchReceiptEmails response");
    expect(gmailState(threads).every((t) => !t.archived)).toBe(true);
  });
});

/** Threads whose receipts fit no transaction sheet (alone or next to a receipt that does fit) */
function buildUnmatchedThreads(): any[] {
  const future = makeThreadWithId(
    "u-future",
    chasePurchaseEmail("4.00", "Future Shop", "2030-01-01T12:00:00Z"),
  );
  const past = makeThreadWithId(
    "u-past-with-note",
    chasePurchaseEmail("5.00", "Old Shop", "2023-06-01T12:00:00Z"),
    {
      from: "Chase <no.reply.alerts@chase.com>",
      subject: "Promo",
      body: "hi",
      date: new Date("2023-06-01T13:00:00Z"),
    },
  );
  // An errored thread with no receipt gets an EmptyReceipt dated at its last message
  const erroredOutside = makeThreadWithId("u-errored-outside", {
    from: "Weird <weird at example.org>",
    subject: "huh",
    body: "",
    date: new Date("2031-05-05T12:00:00Z"),
  });
  // One receipt fits sheet A, the other fits nothing: the whole thread is skipped
  const straddling = makeThreadWithId(
    "u-straddling",
    chasePurchaseEmail("6.00", "Fits A", "2024-02-21T12:00:00Z"),
    chasePurchaseEmail("7.00", "Fits Nothing", "2030-02-02T12:00:00Z"),
  );
  return [future, past, erroredOutside, straddling];
}

/** Runs the import in a fresh environment and returns everything it changed */
async function runFreshImport(withUnmatched: boolean, mark: boolean) {
  vi.resetModules();
  const fresh = installGasGlobals();
  setupScenarioSheets(fresh);
  const matched = buildScenarioThreads();
  const unmatched = withUnmatched ? buildUnmatchedThreads() : [];
  // Interleave: first, middle and last positions
  fresh.GmailApp.threads = [
    unmatched[0],
    ...matched.slice(0, 10),
    unmatched[1],
    unmatched[2],
    ...matched.slice(10),
    unmatched[3],
  ].filter(Boolean);
  const { getChaseReceipts, recordReceiptsCore } = await loadBudgeting();
  const result = recordReceiptsCore(getChaseReceipts(null, null), mark);
  return {
    result,
    sheets: fresh.spreadsheet.dumpSheets(),
    matchedGmail: gmailState(matched),
    unmatched,
    createdLabels: fresh.GmailApp.createdLabels,
  };
}

describe("receipts that fit no transaction sheet (GitHub #1)", () => {
  it.each([true, false])(
    "are skipped whole-thread while the rest imports exactly as without them (mark: %s)",
    async (mark) => {
      const baseline = await runFreshImport(false, mark);
      const mixed = await runFreshImport(true, mark);

      // The matched receipts produce exactly what they produce alone
      expect(mixed.sheets).toEqual(baseline.sheets);
      expect(mixed.matchedGmail).toEqual(baseline.matchedGmail);
      expect(mixed.createdLabels).toEqual(baseline.createdLabels);
      expect(mixed.result.receiptInfosAdded).toEqual(
        baseline.result.receiptInfosAdded,
      );
      expect(mixed.result.numErrors).toBe(baseline.result.numErrors);
      expect(mixed.result.errors).toEqual(baseline.result.errors);

      // The unmatched threads are untouched: not written, not labeled, still in the inbox
      expect(gmailState(mixed.unmatched)).toEqual(
        mixed.unmatched.map((t) => ({
          id: t.getId(),
          labels: [],
          archived: false,
        })),
      );
      // ...and reported, one message per receipt not recorded
      expect(baseline.result.skippedReceipts).toEqual([]);
      expect(mixed.result.skippedReceipts).toHaveLength(5);
      expect(mixed.result.skippedReceipts.join("\n")).toContain("Future Shop");
      expect(mixed.result.skippedReceipts.join("\n")).toContain("Fits A");
    },
  );
});

describe("GAS calls made by the import (end state pinned by the snapshots above)", () => {
  /** The scenario without the thread whose messages cannot be read */
  const readableThreads = () =>
    threads.filter((t) => t.getId() !== "t-thread-fails");

  it("reads all thread messages with one getMessagesForThreads call", async () => {
    env.GmailApp.threads = readableThreads();
    const { getChaseReceipts, recordReceiptsCore } = await loadBudgeting();
    env.calls.reset();
    recordReceiptsCore(getChaseReceipts(null, null), true);
    expect(env.calls.get("GmailApp.getMessagesForThreads")).toBe(1);
    expect(env.calls.get("GmailThread.getMessages")).toBe(0);
    expect(env.calls.get("GmailMessage.getThread")).toBe(0);
    expect(env.calls.get("Spreadsheet.getSheets")).toBe(1);
  });

  it("falls back to per-thread reads when the batch read fails, with the same result", async () => {
    const { getChaseReceipts, recordReceiptsCore } = await loadBudgeting();
    env.calls.reset();
    recordReceiptsCore(getChaseReceipts(null, null), true);
    expect(env.calls.get("GmailApp.getMessagesForThreads")).toBe(1);
    expect(env.calls.get("GmailThread.getMessages")).toBe(threads.length);
    // (sheet and Gmail end state: the first golden test runs this same scenario)
  });

  it("writes notes per run of consecutive rows and backgrounds per color, never per cell", async () => {
    const bad = (id: string, hour: number) =>
      makeThreadWithId(id, {
        from: "Weird <weird at example.org>",
        subject: "huh",
        body: "",
        date: new Date(Date.UTC(2024, 1, 21, hour)),
      });
    env.GmailApp.threads = [
      makeThreadWithId(
        "g1",
        chasePurchaseEmail("1.00", "A", "2024-02-20T12:00:00Z"),
      ),
      bad("b1", 1),
      bad("b2", 2),
      makeThreadWithId(
        "g2",
        chasePurchaseEmail("2.00", "B", "2024-02-22T12:00:00Z"),
      ),
      bad("b3", 3),
    ];
    const { getChaseReceipts, recordReceiptsCore } = await loadBudgeting();
    env.calls.reset();
    recordReceiptsCore(getChaseReceipts(null, null), true);
    // b1, b2, b3 sort into consecutive rows: one setNotes, one red and one pink RangeList write
    expect(env.calls.get("Range.setNotes")).toBe(1);
    expect(env.calls.get("RangeList.setBackground")).toBe(2);
    expect(env.calls.get("Range.setNote")).toBe(0);
    expect(env.calls.get("Range.setBackground")).toBe(0);
    const a = env.spreadsheet.getSheetByName(SHEET_A);
    expect(a.getRange("B6:C9").getBackgrounds()).toEqual([
      // g1's row: untouched cells keep the stale fill
      ["#00ff00", "#00ff00"],
      ["#FF0000", "#E8A9CA"],
      ["#FF0000", "#E8A9CA"],
      ["#FF0000", "#E8A9CA"],
    ]);
    expect(
      a
        .getRange("B6:B10")
        .getNotes()
        .map(([note]: string[]) => note.slice(0, 60)),
    ).toEqual([
      "stale note",
      "Error while processing message b1-m1 with subject huh from d",
      "Error while processing message b2-m1 with subject huh from d",
      "Error while processing message b3-m1 with subject huh from d",
      "",
    ]);
  });

  it("syncReceipts reads the sheet list and time zone once", async () => {
    env.GmailApp.threads = readableThreads();
    vi.useFakeTimers();
    vi.setSystemTime(new Date("2024-03-05T18:00:00Z"));
    env.calls.reset();
    await api("syncReceipts");
    expect(env.calls.get("Spreadsheet.getSheets")).toBe(1);
    expect(env.calls.get("Spreadsheet.getSpreadsheetTimeZone")).toBe(1);
    expect(env.calls.get("GmailThread.getMessages")).toBe(0);
  });

  it("getAllTransactions reads the time zone once for all sheets", async () => {
    env.calls.reset();
    await api("getAllTransactions");
    expect(env.calls.get("Spreadsheet.getSpreadsheetTimeZone")).toBe(1);
    expect(env.calls.get("Spreadsheet.getSheets")).toBe(1);
  });

  it("each API request reads the sheet list again (like a new GAS execution)", async () => {
    await api("getAllTransactions");
    addTransactionSheet(env, "3/17/24 - 3/30/24");
    const all = await api("getAllTransactions");
    expect(all.data.sheets.map((s: any) => s.name)).toContain(
      "3/17/24 - 3/30/24",
    );
  });

  it("fetchReceiptEmails reads messages in one batch (labels still per thread)", async () => {
    env.GmailApp.threads = readableThreads();
    env.calls.reset();
    const res = await api("fetchReceiptEmails", { maxThreads: 50 });
    expect(res.ok).toBe(true);
    expect(env.calls.get("GmailApp.getMessagesForThreads")).toBe(1);
    expect(env.calls.get("GmailThread.getMessages")).toBe(0);
    expect(env.calls.get("GmailThread.getLabels")).toBe(threads.length - 1);
  });
});
