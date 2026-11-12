import { beforeEach, describe, expect, it } from "vitest";
import {
  addVariablesSheet,
  getEnv,
  makeThread,
  MessageInit,
} from "./gas-mocks";
import {
  CHASE,
  FWD,
  receiptFixtures,
  type ReceiptFixture,
} from "./fixtures/receipts";

beforeEach(() => {
  addVariablesSheet(getEnv());
});

/** Run a single message through ThreadList and return the thread info */
async function parse(...messages: MessageInit[]) {
  const { ThreadList } = await import("../src/models/thread-list.model");
  const list = new ThreadList([makeThread(...messages)]);
  return list;
}

async function parseOne(message: MessageInit) {
  const list = await parse(message);
  expect(list.threadInfos).toHaveLength(1);
  const info = list.threadInfos[0];
  expect(info.errors).toEqual([]);
  expect(info.receiptInfos).toHaveLength(1);
  return info.receiptInfos[0];
}

const date = new Date("2024-03-01T18:30:00Z");

// Sample emails and expected results live in test/fixtures/receipts.ts
describe.each([
  ["Chase receipts", /^chase/],
  ["PayPal receipts", /^paypal/],
  ["Venmo receipts", /^venmo/],
])("%s", (_title, namePattern) => {
  const fixtures = Object.entries(receiptFixtures)
    .filter(([name]) => namePattern.test(name))
    .map(([, fixture]) => fixture as ReceiptFixture);

  it.each(fixtures)("$description", async ({ email, expected }) => {
    if (expected === null) {
      // Notes alone are not "information": the thread is dropped
      const list = await parse(email);
      expect(list.threadInfos).toEqual([]);
      expect(getEnv().logs.join("\n")).toContain("Ignoring thread");
      return;
    }
    const r = await parseOne(email);
    for (const [key, value] of Object.entries(expected)) {
      if (key === "date") expect(r.date).toBe(email.date);
      else expect(r[key as keyof typeof r], key).toBe(value);
    }
  });
});

describe("ThreadList error handling and aggregation", () => {
  it("records an error for an unknown sender", async () => {
    const list = await parse({
      id: "m1",
      from: "Spam <spam@example.org>",
      subject: "Hello",
      body: "b",
    });
    expect(list.threadInfos).toHaveLength(1);
    expect(list.threadInfos[0].errors[0]).toContain(
      "No function to handle receipt message from spam@example.org",
    );
    expect(list.threadInfos[0].errors[0]).toContain(
      "Skipping marking as processed",
    );
  });

  it("records an error when the sender cannot be parsed", async () => {
    const list = await parse({ from: "nobody", subject: "s", body: "b" });
    expect(list.threadInfos[0].errors[0]).toContain(
      "Could not get email address from nobody",
    );
  });

  it("records a thread-level error if getMessages throws", async () => {
    const { ThreadList } = await import("../src/models/thread-list.model");
    const thread = makeThread();
    thread.messagesError = new Error("boom");
    const list = new ThreadList([thread]);
    expect(list.threadInfos[0].errors[0]).toContain(
      `Error while processing thread with ID ${thread.getId()}`,
    );
  });

  it("records an error when a forwarded email's original sender cannot be determined", async () => {
    const list = await parse({
      from: FWD,
      subject: "Fwd: hi",
      body: "no forward header here",
    });
    expect(list.threadInfos[0].errors[0]).toContain(
      "Could not get 'from' email address from forwarded email with subject 'Fwd: hi'",
    );
  });

  it("forwards from an unhandled sender produce a blank receipt, which is only a note", async () => {
    const list = await parse({
      from: FWD,
      subject: "Fwd: Hello",
      body: "---------- Forwarded message ---------\nFrom: Someone <x@example.com>\nDate: d\nSubject: s\nTo: t\n\nBody\n",
    });
    expect(list.threadInfos).toEqual([]);
  });

  it("getAllReceiptInfos sorts by date, adds errors/notes to the first receipt, and adds EmptyReceipts", async () => {
    const early = new Date("2024-03-01T00:00:00Z");
    const late = new Date("2024-03-05T00:00:00Z");
    const { ThreadList } = await import("../src/models/thread-list.model");
    const t1 = makeThread({
      from: CHASE,
      subject: "You made a $5.00 transaction with Late Shop",
      body: "",
      date: late,
    });
    // thread with a receipt, a non-receipt message (note) and an error message
    const t2 = makeThread(
      {
        from: CHASE,
        subject: "You made a $1.00 transaction with Early Shop",
        body: "",
        date: early,
      },
      { from: CHASE, subject: "Marketing", body: "buy things", date: early },
      { from: "x@example.org", subject: "weird", body: "", date: early },
    );
    // thread with only an error -> EmptyReceipt dated at its last message
    const t3 = makeThread({
      from: "y@example.org",
      subject: "bad",
      body: "",
      date: new Date("2024-03-03T00:00:00Z"),
    });
    const list = new ThreadList([t1, t2, t3]);
    const all = list.getAllReceiptInfos();

    expect(all.map((r) => r.name ?? "(empty)")).toEqual([
      "Early Shop",
      "(empty)",
      "Late Shop",
    ]);
    expect(all[0].errorMessage).toContain(
      "No function to handle receipt message from x@example.org",
    );
    expect(all[0].note).toContain(
      "Message is not a receipt:\nFrom: no.reply.alerts@chase.com\nSubject: Marketing",
    );
    expect(all[1].errorMessage).toContain("y@example.org");
    expect(all[1].cost).toBeUndefined();
    // cached
    expect(list.getAllReceiptInfos()).toBe(all);
  });

  it("joins multiple errors with a separator", async () => {
    const { ThreadList } = await import("../src/models/thread-list.model");
    const t = makeThread(
      { from: "a@example.org", subject: "1", body: "" },
      { from: "b@example.org", subject: "2", body: "" },
    );
    const [first] = new ThreadList([t]).getAllReceiptInfos();
    expect(first.errorMessage).toContain(
      "\n\n~~~~~~~~~~~~~~~~~ NEXT ERROR ~~~~~~~~~~~~~~~~~\n\n",
    );
  });

  it("serializes to JSON without Gmail objects", async () => {
    const r = await parseOne({
      id: "msg-x",
      from: CHASE,
      subject: "You made a $2.00 transaction with Shop",
      body: "",
      date,
    });
    const json = JSON.parse(JSON.stringify(r));
    expect(json).toMatchObject({
      messageId: "msg-x",
      cost: 2,
      name: "Shop",
      type: "Credit",
    });
    expect(json.message).toBeUndefined();
    expect(json.errorMessage).toBeUndefined();
    expect(json.note).toBeUndefined();
  });
});
