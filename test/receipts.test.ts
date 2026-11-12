import { beforeEach, describe, expect, it } from "vitest";
import { addVariablesSheet, getEnv, makeThread, MessageInit } from "./gas-mocks";

const CHASE = "Chase <no.reply.alerts@chase.com>";
const PAYPAL = "PayPal <service@paypal.com>";
const VENMO = "Venmo <venmo@venmo.com>";
const FWD = "Sam Example <fwd@example.com>";

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

describe("Chase receipts", () => {
  it("parses a normal purchase from the subject", async () => {
    const r = await parseOne({
      from: CHASE,
      subject: "You made a $12.34 transaction with Trader Joe's",
      body: "irrelevant",
      date,
    });
    expect(r.cost).toBe(12.34);
    expect(r.name).toBe("Trader Joe's");
    expect(r.type).toBe("Credit");
    expect(r.category).toBeUndefined();
    expect(r.date).toBe(date);
  });

  it("parses a pending credit as a negative cost with merchant from the body", async () => {
    const r = await parseOne({
      from: CHASE,
      subject: "You have a $45.00 credit pending on your credit card",
      body: "Account ending in 1234\r\nMerchant   Target Store  \r\nDate Mar 1\r\n",
    });
    expect(r.cost).toBe(-45);
    // NOTE: possible bug: greedy `.+` before `s+` keeps trailing spaces in the merchant name
    expect(r.name).toBe("Target Store  ");
    expect(r.type).toBe("Credit");
  });

  it("parses a gas station receipt with the Gas category", async () => {
    const r = await parseOne({
      from: CHASE,
      subject: "You used your card at a gas station",
      body: "Hello\nMerchant   SHELL OIL 123  \n\nMore text\n",
    });
    expect(r.cost).toBeUndefined();
    // NOTE: possible bug: trailing spaces from the body are kept before " (Gas)"
    expect(r.name).toBe("SHELL OIL 123   (Gas)");
    expect(r.category).toBe("Gas");
    expect(r.type).toBe("Credit");
  });

  it("treats an unrecognized Chase subject as a non-receipt with a note", async () => {
    const list = await parse({
      from: CHASE,
      subject: "Your statement is ready",
      body: "x".repeat(500),
      date,
    });
    // Notes alone are not "information": the thread is dropped
    expect(list.threadInfos).toEqual([]);
    expect(getEnv().logs.join("\n")).toContain("Ignoring thread");
  });
});

describe("PayPal receipts", () => {
  const paymentBody =
    "Hello,\n\nYou sent $25.50 to Jane Doe\n\nTransaction ID\nABC123\n\nLunch money\nThanks\n";

  it("parses a payment sent", async () => {
    const r = await parseOne({
      from: PAYPAL,
      to: "someone@example.com",
      subject: "You sent a $25.50 USD payment",
      body: paymentBody,
    });
    expect(r.cost).toBe(25.5);
    expect(r.name).toBe("Jane Doe - Lunch money");
    // Default prefix is TJ when 'to' is not the forward address
    expect(r.type).toBe("TJ Paypal");
  });

  it("uses ForwardName as type prefix when sent to the forward address", async () => {
    const r = await parseOne({
      from: PAYPAL,
      to: "Sam <fwd@example.com>",
      subject: "You sent a $25.50 USD payment",
      body: paymentBody,
    });
    expect(r.type).toBe("Sam Paypal");
  });

  it("accepts a non-breaking space before USD", async () => {
    const r = await parseOne({
      from: PAYPAL,
      subject: "You sent a $9.99 USD payment",
      body: paymentBody,
    });
    expect(r.cost).toBe(9.99);
  });

  it("parses money received as a negative cost", async () => {
    const r = await parseOne({
      from: PAYPAL,
      subject: "Money is waiting for you",
      body: "Hi\n\nAccept your $20.00 USD from Bob Smith\n\nline\nline\n\nBirthday gift\nbye\n",
    });
    expect(r.cost).toBe(-20);
    expect(r.name).toBe("Bob Smith - Birthday gift");
  });

  it("returns an empty receipt (no cost, no name) for unknown PayPal subjects", async () => {
    const list = await parse({
      from: PAYPAL,
      subject: "Update to your account",
      body: "nothing",
    });
    expect(list.threadInfos).toEqual([]);
  });

  it("parses a payment forwarded from the forward address", async () => {
    const r = await parseOne({
      from: FWD,
      subject: "Fwd: You sent a $25.50 USD payment",
      body:
        "FYI\n\n---------- Forwarded message ---------\n" +
        "From: PayPal <service@paypal.com>\nDate: Mon, Nov 27, 2023 at 9:52 PM\n" +
        "Subject: You sent a $25.50 USD payment\nTo: Sam <fwd@example.com>\n\n" +
        paymentBody,
    });
    expect(r.cost).toBe(25.5);
    expect(r.name).toBe("Jane Doe - Lunch money");
    expect(r.type).toBe("Sam Paypal");
  });

  it("parses a forwarded email with the Japanese forwarding header", async () => {
    const r = await parseOne({
      from: FWD,
      subject: "Fwd: You sent a $25.50 USD payment",
      body:
        "メモ\n\n転送されたメッセージ:\n\n*差出人:* PayPal <service@paypal.com>\n" +
        "*日時:* 2024年3月3日\n*宛先:* fwd@example.com\n*件名:* *You sent*\n" +
        paymentBody,
    });
    expect(r.cost).toBe(25.5);
    expect(r.name).toBe("Jane Doe - Lunch money");
  });

  it("parses a forwarded email with the 'Begin forwarded message:' header", async () => {
    const r = await parseOne({
      from: FWD,
      subject: "Fwd: You sent a $25.50 USD payment",
      body:
        "Hi\n\nBegin forwarded message:\n\n*From:* PayPal <service@paypal.com>\n" +
        "*Date:* December 11, 2023\n*To:* Sam <fwd@example.com>\n*Subject:* *You sent*\n\n" +
        paymentBody,
    });
    expect(r.cost).toBe(25.5);
    expect(r.name).toBe("Jane Doe - Lunch money");
  });
});

describe("Venmo receipts", () => {
  it("parses a payment with a note", async () => {
    const r = await parseOne({
      from: VENMO,
      subject: "You paid Jane Doe $15.00",
      body: "You paid\nJane Doe\n$15.00\nPizza night\nTransaction details\n",
    });
    expect(r.cost).toBe(15);
    expect(r.name).toBe("Jane Doe - Pizza night");
    expect(r.type).toBe("TJ Venmo");
  });

  it("parses a completed charge request", async () => {
    const r = await parseOne({
      from: VENMO,
      subject: "You completed Bob's $30.25 charge request",
      body: "You charged\nBob\n$30.25\nRent share\nmore\n",
    });
    expect(r.cost).toBe(30.25);
    expect(r.name).toBe("Bob - Rent share");
  });

  it("parses money received as negative cost; name is NOT taken from the subject", async () => {
    const r = await parseOne({
      from: VENMO,
      subject: "Alice paid you $40.00",
      body: "Alice paid\nyou\n$40.00\nConcert tickets\nmore\n",
    });
    expect(r.cost).toBe(-40);
    // NOTE: possible bug: for received payments `name` is never assigned from the subject, so the
    // recorded name is "undefined - <note>"
    expect(r.name).toBe("undefined - Concert tickets");
  });

  it("parses a forwarded Venmo payment", async () => {
    const r = await parseOne({
      from: FWD,
      subject: "Fwd: You paid Jane Doe $15.00",
      body:
        "---------- Forwarded message ---------\nFrom: Venmo <venmo@venmo.com>\n" +
        "Date: Mon\nSubject: You paid Jane Doe $15.00\nTo: Sam <fwd@example.com>\n\n" +
        "You paid\nJane Doe\n$15.00\nPizza night\nmore\n",
    });
    expect(r.cost).toBe(15);
    expect(r.name).toBe("Jane Doe - Pizza night");
    expect(r.type).toBe("Sam Venmo");
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
      "No function to handle receipt message from spam@example.org"
    );
    expect(list.threadInfos[0].errors[0]).toContain("Skipping marking as processed");
  });

  it("records an error when the sender cannot be parsed", async () => {
    const list = await parse({ from: "nobody", subject: "s", body: "b" });
    expect(list.threadInfos[0].errors[0]).toContain("Could not get email address from nobody");
  });

  it("records a thread-level error if getMessages throws", async () => {
    const { ThreadList } = await import("../src/models/thread-list.model");
    const thread = makeThread();
    thread.messagesError = new Error("boom");
    const list = new ThreadList([thread]);
    expect(list.threadInfos[0].errors[0]).toContain(
      `Error while processing thread with ID ${thread.getId()}`
    );
  });

  it("records an error when a forwarded email's original sender cannot be determined", async () => {
    const list = await parse({
      from: FWD,
      subject: "Fwd: hi",
      body: "no forward header here",
    });
    expect(list.threadInfos[0].errors[0]).toContain(
      "Could not get 'from' email address from forwarded email with subject 'Fwd: hi'"
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
      { from: CHASE, subject: "You made a $1.00 transaction with Early Shop", body: "", date: early },
      { from: CHASE, subject: "Marketing", body: "buy things", date: early },
      { from: "x@example.org", subject: "weird", body: "", date: early }
    );
    // thread with only an error -> EmptyReceipt dated at its last message
    const t3 = makeThread({ from: "y@example.org", subject: "bad", body: "", date: new Date("2024-03-03T00:00:00Z") });
    const list = new ThreadList([t1, t2, t3]);
    const all = list.getAllReceiptInfos();

    expect(all.map((r) => r.name ?? "(empty)")).toEqual(["Early Shop", "(empty)", "Late Shop"]);
    expect(all[0].errorMessage).toContain("No function to handle receipt message from x@example.org");
    expect(all[0].note).toContain("Message is not a receipt:\nFrom: no.reply.alerts@chase.com\nSubject: Marketing");
    expect(all[1].errorMessage).toContain("y@example.org");
    expect(all[1].cost).toBeUndefined();
    // cached
    expect(list.getAllReceiptInfos()).toBe(all);
  });

  it("joins multiple errors with a separator", async () => {
    const { ThreadList } = await import("../src/models/thread-list.model");
    const t = makeThread(
      { from: "a@example.org", subject: "1", body: "" },
      { from: "b@example.org", subject: "2", body: "" }
    );
    const [first] = new ThreadList([t]).getAllReceiptInfos();
    expect(first.errorMessage).toContain(
      "\n\n~~~~~~~~~~~~~~~~~ NEXT ERROR ~~~~~~~~~~~~~~~~~\n\n"
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
