import { describe, expect, it } from "vitest";
import {
  assignRowsToSheets,
  findTransactionSheet,
  joinThreadErrors,
  joinThreadNotes,
  messageErrorText,
  planThreadImport,
  sortRowsByDate,
  threadErrorText,
  type ImportMessageInput,
  type RuleOutcome,
} from "../src/index.ts";

function outcome(
  receipt: Partial<RuleOutcome["receipt"]>,
  reason: RuleOutcome["reason"] = null,
): RuleOutcome {
  const full = {
    cost: undefined,
    name: undefined,
    category: undefined,
    type: undefined,
    ...receipt,
  };
  return {
    senderAddress: null,
    forwarded: false,
    ruleId: null,
    reason,
    receipt: full,
    isReceipt: !!full.cost || !!full.name,
    captures: { subject: null, body: null },
    warnings: [],
  };
}

function message(over: Partial<ImportMessageInput> = {}): ImportMessageInput {
  return {
    messageId: "m1",
    subject: "Subj",
    fromAddress: "Shop@Example.com",
    date: new Date("2025-03-05T12:00:00Z"),
    dateText: "DATE-TEXT",
    plainBody: "body",
    outcome: outcome({ cost: 4.5, name: "Shop", type: "Credit" }),
    error: null,
    ...over,
  };
}

const LAST = new Date("2025-03-06T00:00:00Z");
const last = () => LAST;

describe("planThreadImport", () => {
  it("makes one row per receipt message", () => {
    const plan = planThreadImport({
      threadId: "t1",
      lastMessageDate: last,
      messages: [message()],
    });
    expect(plan.hasInformation).toBe(true);
    expect(plan.rows).toEqual([
      {
        messageId: "m1",
        date: new Date("2025-03-05T12:00:00Z"),
        cost: 4.5,
        name: "Shop",
        category: undefined,
        type: "Credit",
        errorMessage: "",
        note: "",
      },
    ]);
  });

  it("notes a non-receipt message but ignores a thread with only those", () => {
    const plan = planThreadImport({
      threadId: "t1",
      lastMessageDate: last,
      messages: [message({ outcome: outcome({}, "no-rule-matched") })],
    });
    expect(plan.hasInformation).toBe(false);
    expect(plan.rows).toEqual([]);
    expect(plan.notes).toEqual([
      "Message is not a receipt:\nFrom: Shop@Example.com\nSubject: Subj\nDate: DATE-TEXT\nThread ID: t1\n280 Chars of Plain Body:\nbody",
    ]);
  });

  it("truncates the body in the note to 280 chars", () => {
    const plan = planThreadImport({
      threadId: "t1",
      lastMessageDate: last,
      messages: [message({ plainBody: "x".repeat(300), outcome: outcome({}) })],
    });
    expect(plan.notes[0].endsWith(`Body:\n${"x".repeat(280)}`)).toBe(true);
  });

  it("puts a message error on an empty row dated at the last message", () => {
    const plan = planThreadImport({
      threadId: "t1",
      lastMessageDate: last,
      messages: [message({ outcome: null, error: "Error: boom" })],
    });
    expect(plan.errors).toEqual([
      "Error while processing message m1 with subject Subj from date DATE-TEXT on thread t1. Skipping marking as processed. Error: boom",
    ]);
    expect(plan.hasInformation).toBe(true);
    expect(plan.rows).toHaveLength(1);
    expect(plan.rows[0]).toMatchObject({
      messageId: null,
      date: LAST,
      errorMessage: plan.errors[0],
      note: "",
    });
  });

  it("adds errors and notes to the first receipt, joined", () => {
    const plan = planThreadImport({
      threadId: "t1",
      lastMessageDate: last,
      messages: [
        message({ messageId: "a", outcome: outcome({}) }),
        message({ messageId: "b", outcome: outcome({}) }),
        message({ messageId: "c" }),
        message({ messageId: "d", outcome: null, error: "E1" }),
        message({ messageId: "e", outcome: null, error: "E2" }),
      ],
    });
    expect(plan.rows).toHaveLength(1);
    expect(plan.rows[0].messageId).toBe("c");
    expect(plan.rows[0].errorMessage).toBe(joinThreadErrors(plan.errors));
    expect(plan.rows[0].errorMessage).toContain(
      "\n\n~~~~~~~~~~~~~~~~~ NEXT ERROR ~~~~~~~~~~~~~~~~~\n\n",
    );
    expect(plan.rows[0].note).toBe(joinThreadNotes(plan.notes));
    expect(plan.rows[0].note).toContain(
      "\n\n~~~~~~~~~~~~~~~~~ NEXT NOTE ~~~~~~~~~~~~~~~~~\n\n",
    );
  });

  it("sorts rows by date ascending, stably", () => {
    const rows = [
      { id: "late", date: new Date("2025-03-09T00:00:00Z") },
      { id: "tie1", date: new Date("2025-03-01T00:00:00Z") },
      { id: "tie2", date: new Date("2025-03-01T00:00:00Z") },
    ];
    expect(sortRowsByDate(rows).map((r) => r.id)).toEqual([
      "tie1",
      "tie2",
      "late",
    ]);
  });

  it("formats thread-level errors", () => {
    expect(threadErrorText("t9", "Error: x")).toBe(
      "Error while processing thread with ID t9. Skipping marking as processed. Error: x",
    );
    expect(
      messageErrorText({
        messageId: "m",
        subject: "s",
        dateText: "d",
        threadId: "t",
        error: "e",
      }),
    ).toContain("message m with subject s from date d on thread t");
  });
});

describe("findTransactionSheet", () => {
  const sheets = [
    {
      name: "A",
      startDate: new Date("2025-03-01T06:00:00Z"),
      endDate: new Date("2025-03-15T05:59:59Z"),
    },
    {
      name: "B",
      startDate: new Date("2025-03-15T06:00:00Z"),
      endDate: new Date("2025-03-31T05:59:59Z"),
    },
  ];
  it("finds the sheet whose range holds the date", () => {
    expect(findTransactionSheet(new Date("2025-03-20T00:00:00Z"), sheets)).toBe(
      sheets[1],
    );
  });
  it("includes the start and excludes the end", () => {
    expect(findTransactionSheet(sheets[0].startDate, sheets)).toBe(sheets[0]);
    expect(findTransactionSheet(sheets[0].endDate, sheets)).toBeUndefined();
  });
});

describe("assignRowsToSheets", () => {
  const sheets = [
    {
      name: "A",
      startDate: new Date("2025-03-01T06:00:00Z"),
      endDate: new Date("2025-03-15T05:59:59Z"),
    },
  ];
  const inA = new Date("2025-03-05T00:00:00Z");
  const outside = new Date("2025-04-05T00:00:00Z");
  it("records every row of threads whose rows all fit a sheet", () => {
    const rows = [
      { id: 1, thread: "t1", date: inA },
      { id: 2, thread: "t2", date: inA },
    ];
    const { recorded, skipped } = assignRowsToSheets(
      rows,
      sheets,
      (r) => r.thread,
    );
    expect(recorded).toEqual([
      { row: rows[0], sheet: sheets[0] },
      { row: rows[1], sheet: sheets[0] },
    ]);
    expect(skipped).toEqual([]);
  });
  it("skips a whole thread with a row that fits no sheet, keeping input order", () => {
    const rows = [
      { id: 1, thread: "t1", date: inA },
      { id: 2, thread: "t2", date: inA },
      { id: 3, thread: "t1", date: outside },
    ];
    const { recorded, skipped } = assignRowsToSheets(
      rows,
      sheets,
      (r) => r.thread,
    );
    expect(recorded).toEqual([{ row: rows[1], sheet: sheets[0] }]);
    expect(skipped).toEqual([
      { row: rows[0], sheet: sheets[0] },
      { row: rows[2], sheet: undefined },
    ]);
  });
});
