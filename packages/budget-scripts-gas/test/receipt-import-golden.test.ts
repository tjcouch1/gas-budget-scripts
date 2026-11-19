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
    const result = recordReceiptsCore(getChaseReceipts(null, null), true);
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
