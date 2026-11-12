import { afterEach, beforeEach, describe, expect, it, vi } from "vitest";
import {
  addTransactionSheet,
  addVariablesSheet,
  getEnv,
  makeThread,
  AnySheet,
} from "./gas-mocks";

const load = () => import("../src/util/budgeting.util");

/** Format like the sheet names: M/D/YY */
const fmt = (d: Date) =>
  `${d.getUTCMonth() + 1}/${d.getUTCDate()}/${String(d.getUTCFullYear()).slice(2)}`;

beforeEach(() => {
  const env = getEnv();
  // First sheet is the "active" one (menu lives there in index tests)
  env.spreadsheet.addSheet("Main");
  addVariablesSheet(env);
});

afterEach(() => {
  vi.useRealTimers();
});

describe("getTransactionSheetInfos", () => {
  it("finds sheets named 'start - end' and parses their dates at GMT-06:00", async () => {
    const env = getEnv();
    addTransactionSheet(env, "2/18/24 - 3/2/24");
    addTransactionSheet(env, "2/4/24 - 2/17/24 (Gap)");
    env.spreadsheet.addSheet("Template");
    const { getTransactionSheetInfos } = await load();
    const infos = getTransactionSheetInfos();
    expect(infos.map((i) => i.sheet.getName())).toEqual([
      "2/18/24 - 3/2/24",
      "2/4/24 - 2/17/24 (Gap)",
    ]);
    expect(infos[0].startDate.toISOString()).toBe("2024-02-18T06:00:00.000Z");
    expect(infos[0].endDate.toISOString()).toBe("2024-03-03T05:59:59.999Z");
    expect(infos[0].receiptInfosToAdd).toEqual([]);
  });
});

describe("addTransactionSheet(s)", () => {
  function setupSheets(count = 2) {
    const env = getEnv();
    env.spreadsheet.addSheet("Template").seed("A1", [["template-cell"]]);
    // Chronological by position, newest first (matches how new sheets are inserted)
    let start = Date.UTC(2024, 1, 18, 12);
    const names: string[] = [];
    for (let i = 0; i < count; i++) {
      names.push(
        `${fmt(new Date(start))} - ${fmt(new Date(start + 13 * 86400000))}`,
      );
      start -= 14 * 86400000;
    }
    names.forEach((n) => addTransactionSheet(env, n));
    return { env, names };
  }

  it("throws if there are no transaction sheets", async () => {
    getEnv().spreadsheet.addSheet("Template");
    const { addTransactionSheet: add } = await load();
    expect(() => add()).toThrow(
      "No transaction sheets found in order to add more",
    );
  });

  it("returns false and adds nothing when the latest sheet has not ended", async () => {
    const { env } = setupSheets();
    vi.useFakeTimers();
    vi.setSystemTime(new Date("2024-03-01T12:00:00Z"));
    const before = env.spreadsheet.getSheets().length;
    const { addTransactionSheet: add } = await load();
    expect(add()).toBe(false);
    expect(env.spreadsheet.getSheets()).toHaveLength(before);
  });

  it("adds the next pay period sheet from the template, before the latest sheet", async () => {
    const { env, names } = setupSheets();
    vi.useFakeTimers();
    vi.setSystemTime(new Date("2024-03-10T12:00:00Z"));
    const { addTransactionSheet: add } = await load();
    expect(add()).toBe(true);

    const added = env.spreadsheet.getSheetByName("3/3/24 - 3/16/24");
    expect(added).not.toBeNull();
    // Copied from the template
    expect(added!.getRange("A1").getValue()).toBe("template-cell");
    // Inserted at (latest sheet's index - 1), i.e. directly before the latest transaction sheet
    const order = env.spreadsheet.getSheets().map((s) => s.getName());
    expect(order.indexOf("3/3/24 - 3/16/24")).toBe(order.indexOf(names[0]) - 1);
    expect(added!.tabColor).toBeNull();
  });

  it("throws if the template sheet is missing", async () => {
    const { env } = setupSheets();
    env.spreadsheet.sheets = env.spreadsheet.sheets.filter(
      (s) => s.getName() !== "Template",
    );
    vi.useFakeTimers();
    vi.setSystemTime(new Date("2024-03-10T12:00:00Z"));
    const { addTransactionSheet: add } = await load();
    expect(() => add()).toThrow(
      "Template sheet not found! Cannot duplicate to create sheet 3/3/24 - 3/16/24.",
    );
  });

  it("addTransactionSheets keeps adding until caught up and returns the count", async () => {
    const { env } = setupSheets();
    vi.useFakeTimers();
    vi.setSystemTime(new Date("2024-04-01T12:00:00Z"));
    const { addTransactionSheets } = await load();
    expect(addTransactionSheets()).toBe(3);
    const names = env.spreadsheet.getSheets().map((s) => s.getName());
    expect(names).toContain("3/3/24 - 3/16/24");
    expect(names).toContain("3/17/24 - 3/30/24");
    expect(names).toContain("3/31/24 - 4/13/24");
    expect(addTransactionSheets()).toBe(0);
  });

  it("marks the new sheet as a Gap period when the 14th newest sheet is a gap", async () => {
    const { env, names } = setupSheets(14);
    // 14th newest sheet (index 13 in descending order) is a gap
    env.spreadsheet.getSheetByName(names[13])!.rename(`${names[13]} (Gap)`);
    vi.useFakeTimers();
    vi.setSystemTime(new Date("2024-03-10T12:00:00Z"));
    const { addTransactionSheet: add } = await load();
    expect(add()).toBe(true);
    const added = env.spreadsheet.getSheetByName("3/3/24 - 3/16/24 (Gap)");
    expect(added).not.toBeNull();
    expect(added!.tabColor).toBe("#E8A9CA");
  });
});

describe("recording receipts", () => {
  const SHEET = "2/18/24 - 3/2/24";
  let sheet: AnySheet;

  beforeEach(() => {
    sheet = addTransactionSheet(getEnv(), SHEET, { max: 5 });
  });

  const chase = (cost: string, merchant: string, date: string) => ({
    from: "Chase <no.reply.alerts@chase.com>",
    subject: `You made a $${cost} transaction with ${merchant}`,
    body: "",
    date: new Date(date),
  });

  it("queries Gmail with the receipt search and the start/max arguments", async () => {
    const env = getEnv();
    const { getAndRecordReceipts } = await load();
    getAndRecordReceipts(5, 20);
    getAndRecordReceipts(null, null);
    expect(env.GmailApp.searches[0].start).toBe(5);
    expect(env.GmailApp.searches[0].max).toBe(20);
    expect(env.GmailApp.searches[0].query).toBe(
      "from:(no.reply.alerts@chase.com OR service@paypal.com OR venmo@venmo.com OR fwd@example.com) in:inbox NOT label:receipts NOT label:receipts-cru-reimburse NOT label:receipts-tax-deductible NOT label:receipts-scripted ",
    );
    expect(env.GmailApp.searches[1].start).toBeUndefined();
  });

  it("writes receipts in date order into the first empty rows with metadata columns", async () => {
    const env = getEnv();
    sheet.seed("A5", [[new Date("2024-02-19T12:00:00Z"), "Existing", 1]]);
    env.GmailApp.threads = [
      makeThread(chase("20.00", "Later Shop", "2024-02-25T12:00:00Z")),
      makeThread(chase("10.50", "Earlier Shop", "2024-02-20T12:00:00Z")),
    ];
    const { getAndRecordReceipts } = await load();
    const result = getAndRecordReceipts(null, null, false);

    expect(result).toEqual({ [SHEET]: 2 });
    expect(sheet.getRange("A5:C7").getValues()).toEqual([
      [new Date("2024-02-19T12:00:00Z"), "Existing", 1],
      [new Date("2024-02-20T12:00:00Z"), "Earlier Shop", 10.5],
      [new Date("2024-02-25T12:00:00Z"), "Later Shop", 20],
    ]);
    // Category (E) and type (F) columns
    expect(sheet.getRange("E6:F7").getValues()).toEqual([
      ["", "Credit"],
      ["", "Credit"],
    ]);
  });

  it("does not mark threads processed unless asked", async () => {
    const env = getEnv();
    const thread = makeThread(chase("1.00", "Shop", "2024-02-20T12:00:00Z"));
    env.GmailApp.threads = [thread];
    const { getAndRecordReceipts } = await load();
    getAndRecordReceipts(null, null, false);
    expect(thread.archived).toBe(false);
    expect(thread.labels).toEqual([]);
    expect(env.logs).toContain(`Would mark thread ${thread.getId()} processed`);
  });

  it("labels (Receipts/Scripted + Receipts) and archives threads when marking", async () => {
    const env = getEnv();
    const t1 = makeThread(chase("1.00", "Shop", "2024-02-20T12:00:00Z"));
    const t2 = makeThread(chase("2.00", "Shop2", "2024-02-21T12:00:00Z"));
    env.GmailApp.threads = [t1, t2];
    const { getAndRecordReceipts } = await load();
    getAndRecordReceipts(null, null, true);

    for (const t of [t1, t2]) {
      expect(t.labels.map((l) => l.getName())).toEqual([
        "Receipts/Scripted",
        "Receipts",
      ]);
      expect(t.archived).toBe(true);
    }
    // Labels are created once and memoized
    expect(env.GmailApp.createdLabels).toEqual([
      "Receipts/Scripted",
      "Receipts",
    ]);
  });

  it("reuses existing Gmail labels instead of creating them", async () => {
    const env = getEnv();
    env.GmailApp.createLabel("Receipts/Scripted");
    env.GmailApp.createLabel("Receipts");
    env.GmailApp.createdLabels.length = 0;
    env.GmailApp.threads = [
      makeThread(chase("1.00", "Shop", "2024-02-20T12:00:00Z")),
    ];
    const { getAndRecordReceipts } = await load();
    getAndRecordReceipts(null, null, true);
    expect(env.GmailApp.createdLabels).toEqual([]);
  });

  it("highlights empty costs", async () => {
    const env = getEnv();
    env.GmailApp.threads = [
      makeThread({
        from: "Chase <no.reply.alerts@chase.com>",
        subject: "You used your card at a gas station",
        body: "x\nMerchant   Shell  \n\ny\n",
        date: new Date("2024-02-20T12:00:00Z"),
      }),
    ];
    const { getAndRecordReceipts } = await load();
    getAndRecordReceipts(null, null, false);
    expect(sheet.getRange("A5:C5").getValues()[0].slice(1)).toEqual([
      "Shell   (Gas)",
      "",
    ]);
    expect(sheet.getRange("E5:F5").getValues()).toEqual([["Gas", "Credit"]]);
    expect(sheet.getRange("C5").getBackgrounds()).toEqual([["#E8A9CA"]]);
    expect(sheet.getRange("B5").getBackgrounds()).toEqual([[null]]);
  });

  it("notes errors in red on the name cell, does not mark errored threads, and alerts the user", async () => {
    const env = getEnv();
    const bad = makeThread({
      from: "Weird <weird@example.org>",
      subject: "huh",
      body: "",
      date: new Date("2024-02-20T12:00:00Z"),
    });
    const good = makeThread(chase("1.00", "Shop", "2024-02-21T12:00:00Z"));
    env.GmailApp.threads = [bad, good];
    const { getAndRecordReceipts } = await load();
    getAndRecordReceipts(null, null, true);

    expect(sheet.getRange("B5").getBackgrounds()).toEqual([["#FF0000"]]);
    expect(sheet.getRange("B5").getValue()).toBe("");
    expect(sheet.getRange("B5").getNotes()[0][0]).toContain(
      "weird@example.org",
    );
    expect(sheet.getRange("C5").getBackgrounds()).toEqual([["#E8A9CA"]]);
    expect(bad.archived).toBe(false);
    expect(good.archived).toBe(true);
    expect(env.SpreadsheetApp.ui.alerts).toEqual([
      "There were 1 errors while processing. Please review.",
    ]);
  });

  it("combines error and note text in the cell note", async () => {
    const env = getEnv();
    env.GmailApp.threads = [
      makeThread(
        chase("1.00", "Shop", "2024-02-20T12:00:00Z"),
        {
          from: "Chase <no.reply.alerts@chase.com>",
          subject: "Promo",
          body: "hi",
          date: new Date("2024-02-20T12:00:00Z"),
        },
        {
          from: "z@example.org",
          subject: "z",
          body: "",
          date: new Date("2024-02-20T12:00:00Z"),
        },
      ),
    ];
    const { getAndRecordReceipts } = await load();
    getAndRecordReceipts(null, null, false);
    const note = sheet.getRange("B5").getNotes()[0][0];
    expect(note).toContain(
      "No function to handle receipt message from z@example.org",
    );
    expect(note).toContain("\n\n>>>>>>>>>> NOTES <<<<<<<<<<\n\n");
    expect(note).toContain("Message is not a receipt:");
  });

  it("throws (after alerting is skipped) with the original UI error when alerting fails in a trigger", async () => {
    const env = getEnv();
    env.SpreadsheetApp.uiAvailable = false;
    env.GmailApp.threads = [
      makeThread({
        from: "a@example.org",
        subject: "s",
        body: "",
        date: new Date("2024-02-20T12:00:00Z"),
      }),
    ];
    const { getAndRecordReceipts } = await load();
    expect(() => getAndRecordReceipts(null, null, false)).toThrow(
      /Error while trying to notify user via UI of 1 errors while processing!/,
    );
  });

  it("throws when a receipt date fits no transaction sheet", async () => {
    getEnv().GmailApp.threads = [
      makeThread(chase("1.00", "Shop", "2025-01-01T12:00:00Z")),
    ];
    const { getAndRecordReceipts } = await load();
    expect(() => getAndRecordReceipts(null, null, false)).toThrow(
      /Could not find transaction sheet for receipt/,
    );
  });

  it("throws when the sheet has too few empty rows", async () => {
    const env = getEnv();
    sheet.seed("A5", [
      [new Date(), "a", 1],
      [new Date(), "b", 1],
      [new Date(), "c", 1],
      [new Date(), "d", 1],
    ]);
    env.GmailApp.threads = [
      makeThread(chase("1.00", "S1", "2024-02-20T12:00:00Z")),
      makeThread(chase("1.00", "S2", "2024-02-20T12:00:00Z")),
    ];
    const { getAndRecordReceipts } = await load();
    expect(() => getAndRecordReceipts(null, null, false)).toThrow(
      /There are not enough empty transaction rows in sheet 2\/18\/24 - 3\/2\/24 to get 2 transaction rows! First empty transaction row: 9\. TransactionsMax: 5\. Last transaction row: 9/,
    );
  });

  it("does nothing for a non-receipt-only thread: it is neither recorded nor marked", async () => {
    const env = getEnv();
    const t = makeThread({
      from: "Chase <no.reply.alerts@chase.com>",
      subject: "Statement ready",
      body: "hi",
    });
    env.GmailApp.threads = [t];
    const { getAndRecordReceipts } = await load();
    expect(getAndRecordReceipts(null, null, true)).toEqual({});
    // NOTE: possible bug: threads with only non-receipt messages are dropped without being
    // labeled/archived, so they are re-scanned on every run (AGENTS.md describes a
    // Receipts/Ignored label for this, which does not exist in the code)
    expect(t.archived).toBe(false);
    expect(t.labels).toEqual([]);
  });
});

describe("splitTransaction", () => {
  const D = new Date("2024-02-20T12:00:00Z");
  let sheet: AnySheet;

  beforeEach(() => {
    sheet = addTransactionSheet(getEnv(), "2/18/24 - 3/2/24", { max: 6 });
  });

  /** Values (A:C) of rows 5..n */
  const rows = (n: number) => sheet.getRange(5, 1, n, 3).getValues();

  it("splits a plain-cost transaction", async () => {
    sheet.seed("A5", [[D, "Store", 5.5]]);
    sheet.seed("E5", [["Food", "Credit"]]);
    const { splitTransaction } = await load();
    const range = splitTransaction(sheet, 0);

    expect(rows(2)).toEqual([
      [D, "Store | ", "=5.5-R[1]C[0]"],
      [D, "Store | ", "=1.0825*(0)"],
    ]);
    expect(sheet.getRange("C5").getFormulas()).toEqual([["=5.5-R[1]C[0]"]]);
    // New row copies the type but not the category
    expect(sheet.getRange("E5:F6").getValues()).toEqual([
      ["Food", "Credit"],
      ["", "Credit"],
    ]);
    // Returned range spans the whole new group (date..cost)
    expect([
      range.getRow(),
      range.getColumn(),
      range.getNumRows(),
      range.getNumColumns(),
    ]).toEqual([5, 1, 2, 3]);
  });

  it("keeps an existing formula cost and does not double the ' | ' in the name", async () => {
    sheet.seed("A5", [[D, "Store | milk", "=3+2"]]);
    const { splitTransaction } = await load();
    splitTransaction(sheet, 0);
    expect(rows(2)).toEqual([
      [D, "Store | milk", "=3+2-R[1]C[0]"],
      [D, "Store | ", "=1.0825*(0)"],
    ]);
  });

  it("splitting a group again appends another subtraction and another child row", async () => {
    sheet.seed("A5", [[D, "Store", 5.5]]);
    const { splitTransaction } = await load();
    splitTransaction(sheet, 0);
    // Pretend the user edited the child row
    sheet.seed("B6", [["Store | bread"]]);
    splitTransaction(sheet, 0);
    expect(rows(3)).toEqual([
      [D, "Store | ", "=5.5-R[1]C[0]-R[2]C[0]"],
      [D, "Store | bread", "=1.0825*(0)"],
      [D, "Store | ", "=1.0825*(0)"],
    ]);
  });

  it("pushes following transactions down one row", async () => {
    const D2 = new Date("2024-02-21T12:00:00Z");
    sheet.seed("A5", [
      [D, "First", 10],
      [D2, "Second", 20],
      [D2, "Third", 30],
    ]);
    sheet.seed("E5", [
      ["c1", "t1"],
      ["c2", "t2"],
      ["c3", "t3"],
    ]);
    const { splitTransaction } = await load();
    splitTransaction(sheet, 0);
    expect(rows(5)).toEqual([
      [D, "First | ", "=10-R[1]C[0]"],
      [D, "First | ", "=1.0825*(0)"],
      [D2, "Second", 20],
      [D2, "Third", 30],
      ["", "", ""],
    ]);
    expect(sheet.getRange("E5:F8").getValues()).toEqual([
      ["c1", "t1"],
      ["", "t1"],
      ["c2", "t2"],
      ["c3", "t3"],
    ]);
  });

  it("splits a transaction in the middle of the list (index > 0)", async () => {
    const D2 = new Date("2024-02-21T12:00:00Z");
    sheet.seed("A5", [
      [D, "First", 10],
      [D2, "Second", 20],
    ]);
    const { splitTransaction } = await load();
    splitTransaction(sheet, 1);
    expect(rows(3)).toEqual([
      [D, "First", 10],
      [D2, "Second | ", "=20-R[1]C[0]"],
      [D2, "Second | ", "=1.0825*(0)"],
    ]);
  });

  it("rejects splitting an empty row", async () => {
    const { splitTransaction } = await load();
    expect(() => splitTransaction(sheet, 0)).toThrow(
      "Checked row 5 on sheet 2/18/24 - 3/2/24, but the transaction row has no content!",
    );
  });

  it("rejects splitting a row in the middle of a group", async () => {
    sheet.seed("A5", [
      [D, "Store | ", "=5-R[1]C[0]"],
      [D, "Store | ", "=1.0825*(0)"],
    ]);
    const { splitTransaction } = await load();
    expect(() => splitTransaction(sheet, 1)).toThrow(
      "Transaction in sheet 2/18/24 - 3/2/24 at row 6 seems to be in a transaction group already!",
    );
  });

  it("rejects when there is no room left on the sheet", async () => {
    sheet.seed(
      "A5",
      Array.from({ length: 6 }, () => [D, "Same", 1]),
    );
    const { splitTransaction } = await load();
    expect(() => splitTransaction(sheet, 0)).toThrow(
      "No room on sheet 2/18/24 - 3/2/24 to split group of 6 starting at row 5!",
    );
  });

  it("groups rows whose names share the prefix before ' | ' or the first third", async () => {
    sheet.seed("A5", [
      [D, "Costco | milk", 5],
      [D, "Costco | ", 1],
      [D, "Costco", 2],
    ]);
    const { splitTransaction } = await load();
    splitTransaction(sheet, 0);
    // All three rows were treated as one group (parent subtracts R[3]); the new row is the 4th
    expect(sheet.getRange("C5").getFormulas()).toEqual([["=5-R[3]C[0]"]]);
    expect(rows(4)[3]).toEqual([D, "Costco | ", "=1.0825*(0)"]);
  });

  it("getTransactionSplitCheckboxesRange spans TransactionsMax rows in the checkbox column", async () => {
    const { getTransactionSplitCheckboxesRange } = await load();
    const r = getTransactionSplitCheckboxesRange(sheet);
    expect([
      r.getRow(),
      r.getColumn(),
      r.getNumRows(),
      r.getNumColumns(),
    ]).toEqual([5, 7, 6, 1]);
  });
});
