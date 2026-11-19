/* eslint-disable @typescript-eslint/no-explicit-any */
import { beforeEach, describe, expect, it } from "vitest";
import {
  addTransactionSheet,
  addVariablesSheet,
  DEFAULT_VARIABLES,
  GasEnv,
  getEnv,
  makeThread,
  MockLabel,
  MockTextOutput,
} from "./gas-mocks";
import { chasePurchaseEmail } from "./fixtures/receipts";

const load = () => import("../src/index");

const KEY = "test-api-key-0123456789abcdef";
const EXCLUSIONS =
  "NOT label:receipts NOT label:receipts-cru-reimburse NOT label:receipts-tax-deductible NOT label:receipts-scripted";
const DEFAULT_QUERY = `(from:(no.reply.alerts@chase.com OR service@paypal.com OR venmo@venmo.com) OR (from:fwd@example.com subject:Fwd)) in:inbox ${EXCLUSIONS} `;

let env: GasEnv;

function setup(extraVariables: unknown[][] = []) {
  addVariablesSheet(env, [...DEFAULT_VARIABLES, ...extraVariables]);
}

beforeEach(() => {
  env = getEnv();
  env.PropertiesService.scriptProperties.setProperty("API_KEY", KEY);
});

async function call(request: Record<string, unknown>): Promise<any> {
  const { doPost } = await load();
  const output = doPost({
    postData: { contents: JSON.stringify(request) },
  } as unknown as GoogleAppsScript.Events.DoPost) as unknown as MockTextOutput;
  expect(output.getMimeType()).toBe("JSON");
  return JSON.parse(output.getContent());
}

const api = (action: string, fields: Record<string, unknown> = {}) =>
  call({ apiVersion: 1, apiKey: KEY, action, ...fields });

/** Gmail state that must stay untouched by read-only calls */
function expectGmailUntouched(threads: any[]) {
  expect(env.GmailApp.createdLabels).toEqual([]);
  expect(env.GmailApp.moveThreadsToArchiveCalls).toEqual([]);
  for (const t of threads) expect(t.archived).toBe(false);
}

describe("fetchReceiptEmails", () => {
  beforeEach(() => setup());

  it("rejects a wrong API key", async () => {
    const res = await call({
      apiVersion: 1,
      apiKey: "wrong",
      action: "fetchReceiptEmails",
    });
    expect(res.error.code).toBe("unauthorized");
    expect(env.GmailApp.searches).toEqual([]);
  });

  it("searches with the default senders and returns each message read-only", async () => {
    const thread = makeThread(
      {
        ...chasePurchaseEmail("4.50", "Demo Cafe", "2024-03-01T18:30:00Z"),
        id: "m1",
        to: "me@example.com",
        date: new Date("2024-03-01T18:30:00Z"),
      },
      {
        from: "Sam Example <fwd@example.com>",
        subject: "Re: thanks",
        body: "second",
        id: "m2",
      },
    );
    const existing = new MockLabel("Receipts/Error");
    thread.labels.push(existing);
    env.GmailApp.threads = [thread];

    const res = await api("fetchReceiptEmails");

    expect(res.ok).toBe(true);
    expect(env.GmailApp.searches).toEqual([
      { query: DEFAULT_QUERY, start: 0, max: 10 },
    ]);
    expect(res.data.query).toBe(DEFAULT_QUERY);
    expect(res.data.threadCount).toBe(1);
    expect(res.data.truncated).toBe(false);
    expect(res.data.emails).toHaveLength(2);
    expect(res.data.emails[0]).toEqual({
      threadId: thread.getId(),
      messageId: "m1",
      from: thread.messages[0].getFrom(),
      to: "me@example.com",
      subject: thread.messages[0].getSubject(),
      date: "2024-03-01T18:30:00.000Z",
      plainBody: thread.messages[0].getPlainBody(),
      plainBodyTruncated: false,
      labels: ["Receipts/Error"],
    });
    expect(res.data.emails[1].messageId).toBe("m2");
    expect(thread.labels).toEqual([existing]);
    expectGmailUntouched([thread]);
    expect(env.LockService.scriptLock.calls).toEqual([]);
  });

  it("lists the transaction sheets (name and date range) and changes nothing on the sheets or in Gmail", async () => {
    const jan = addTransactionSheet(env, "3/1/2024 - 3/15/2024");
    addTransactionSheet(env, "3/16/2024 - 3/31/2024");
    jan.seed("A5", [["3/2/2024", "Existing", 1]]);
    const thread = makeThread({
      ...chasePurchaseEmail("4.50", "Demo Cafe", "2024-03-01T18:30:00Z"),
      id: "m1",
    });
    env.GmailApp.threads = [thread];
    const sheetsBefore = JSON.stringify(
      env.spreadsheet
        .getSheets()
        .map((sh: any) => [sh.getName(), sh.getRange("A1:J12").getValues()]),
    );

    const res = await api("fetchReceiptEmails");

    expect(res.ok).toBe(true);
    expect(res.data.transactionSheets).toEqual([
      {
        name: "3/1/2024 - 3/15/2024",
        startDate: "2024-03-01T06:00:00.000Z",
        endDate: "2024-03-16T05:59:59.999Z",
      },
      {
        name: "3/16/2024 - 3/31/2024",
        startDate: "2024-03-16T06:00:00.000Z",
        endDate: "2024-04-01T05:59:59.999Z",
      },
    ]);
    expect(
      JSON.stringify(
        env.spreadsheet
          .getSheets()
          .map((sh: any) => [sh.getName(), sh.getRange("A1:J12").getValues()]),
      ),
    ).toBe(sheetsBefore);
    expectGmailUntouched([thread]);
    expect(thread.labels).toEqual([]);
    expect(env.LockService.scriptLock.calls).toEqual([]);
  });

  it("uses the senders, subjects, includeProcessed, newerThanDays and maxThreads given", async () => {
    const res = await api("fetchReceiptEmails", {
      senders: ["a@shop.example.com", "B@Other.example.com"],
      subjects: ["Your order", "Receipt"],
      includeProcessed: true,
      newerThanDays: 30,
      maxThreads: 5,
    });
    expect(res.ok).toBe(true);
    const query =
      '(from:(a@shop.example.com OR B@Other.example.com) OR (from:fwd@example.com subject:Fwd)) subject:("Your order" OR "Receipt") newer_than:30d ';
    expect(res.data.query).toBe(query);
    expect(env.GmailApp.searches).toEqual([{ query, start: 0, max: 5 }]);
  });

  it("uses the forward subject prefixes given (the candidate rules' forwarding.subjectPrefixes)", async () => {
    const res = await api("fetchReceiptEmails", {
      senders: ["a@shop.example.com"],
      forwardSubjectPrefixes: ["Fwd: ", "FW: "],
    });
    expect(res.data.query).toBe(
      `(from:(a@shop.example.com) OR (from:fwd@example.com subject:(Fwd OR FW))) in:inbox ${EXCLUSIONS} `,
    );
  });

  it("clamps maxThreads to the server cap of 50", async () => {
    await api("fetchReceiptEmails", { maxThreads: 500 });
    expect(env.GmailApp.searches[0].max).toBe(50);
  });

  it.each([
    [{ senders: [] }, "senders must be an array of 1-30 email addresses"],
    [
      { senders: Array.from({ length: 31 }, (_, i) => `s${i}@example.com`) },
      "senders must be an array of 1-30 email addresses",
    ],
    [
      { senders: ["a@example.com", "b@example.com", "not an email"] },
      "senders[2] is not a valid email address",
    ],
    [
      { subjects: ["Order (new)"] },
      "subjects[0] contains characters that are not allowed",
    ],
    [
      { subjects: ["a", "b", "c", "d", "e", "f"] },
      "subjects must be an array of 0-5 subject filters",
    ],
    [{ maxThreads: 0 }, "maxThreads must be an integer from 1 to 50"],
    [{ maxThreads: 2.5 }, "maxThreads must be an integer from 1 to 50"],
    [
      { newerThanDays: 4000 },
      "newerThanDays must be an integer from 1 to 3650",
    ],
    [{ includeHtml: "yes" }, "includeHtml must be a boolean"],
    [{ includeProcessed: 1 }, "includeProcessed must be a boolean"],
    [
      { forwardSubjectPrefixes: [] },
      "forwardSubjectPrefixes must be an array of 1-5 forward subject prefixes",
    ],
    [
      { forwardSubjectPrefixes: ["Fwd: ", "[FW] "] },
      'forwardSubjectPrefixes[1] must be letters or digits, optionally followed by ":" and spaces',
    ],
  ])("rejects %j", async (fields, message) => {
    const res = await api("fetchReceiptEmails", fields);
    expect(res).toEqual({ ok: false, error: { code: "bad_request", message } });
    expect(env.GmailApp.searches).toEqual([]);
  });

  it("adds htmlBody only when includeHtml is true", async () => {
    env.GmailApp.threads = [
      makeThread({
        from: "x@example.com",
        subject: "s",
        body: "plain",
        html: "<p>hi</p>",
      }),
    ];
    const plain = await api("fetchReceiptEmails");
    expect(plain.data.emails[0]).not.toHaveProperty("htmlBody");
    const html = await api("fetchReceiptEmails", { includeHtml: true });
    expect(html.data.emails[0].htmlBody).toBe("<p>hi</p>");
    expect(html.data.emails[0].htmlBodyTruncated).toBe(false);
  });

  it("cuts long bodies and flags them", async () => {
    env.GmailApp.threads = [
      makeThread({
        from: "x@example.com",
        subject: "s",
        body: "p".repeat(25_000),
        html: "h".repeat(70_000),
      }),
    ];
    const res = await api("fetchReceiptEmails", { includeHtml: true });
    const [email] = res.data.emails;
    expect(email.plainBody).toBe("p".repeat(20_000));
    expect(email.plainBodyTruncated).toBe(true);
    expect(email.htmlBody).toBe("h".repeat(60_000));
    expect(email.htmlBodyTruncated).toBe(true);
    expect(res.data.truncated).toBe(false);
  });

  it("stops at 100 messages and reports truncated", async () => {
    env.GmailApp.threads = Array.from({ length: 50 }, (_, t) =>
      makeThread(
        ...Array.from({ length: 3 }, (_, m) => ({
          from: "x@example.com",
          subject: `t${t} m${m}`,
          body: "b",
        })),
      ),
    );
    const res = await api("fetchReceiptEmails", { maxThreads: 50 });
    expect(res.data.emails).toHaveLength(100);
    expect(res.data.threadCount).toBe(50);
    expect(res.data.truncated).toBe(true);
    expectGmailUntouched(env.GmailApp.threads);
  });

  it("stops when the total body size passes 3,000,000 characters", async () => {
    env.GmailApp.threads = Array.from({ length: 50 }, () =>
      makeThread({
        from: "x@example.com",
        subject: "s",
        body: "p".repeat(20_000),
        html: "h".repeat(60_000),
      }),
    );
    const res = await api("fetchReceiptEmails", {
      includeHtml: true,
      maxThreads: 50,
    });
    // 80,000 characters per message: the 38th would pass the cap
    expect(res.data.emails).toHaveLength(37);
    expect(res.data.truncated).toBe(true);
  });

  it("reports an invalid ReceiptRules variable as invalid_config", async () => {
    addVariablesSheetReplacing([["ReceiptRules", "{not json"]]);
    const res = await api("fetchReceiptEmails");
    expect(res.error.code).toBe("invalid_config");
    expect(res.error.message).toMatch(/^ReceiptRules variable is invalid/);
    expect(env.GmailApp.searches).toEqual([]);
  });

  it("does not read ReceiptRules when senders are given", async () => {
    addVariablesSheetReplacing([["ReceiptRules", "{not json"]]);
    const res = await api("fetchReceiptEmails", {
      senders: ["a@example.com"],
    });
    expect(res.ok).toBe(true);
  });
});

/** Replace the Variables sheet with the defaults plus `extra` rows */
function addVariablesSheetReplacing(extra: unknown[][]) {
  env.spreadsheet.getSheetByName("Variables")!.seed("A8", extra);
}

const shared = () => import("@budget/shared");
const variablesSheet = () => env.spreadsheet.getSheetByName("Variables")!;
const cell = (a1: string) =>
  variablesSheet().getRange(a1).getValues()[0][0] as unknown;

describe("getReceiptRules", () => {
  it("reports built-in defaults when the key is missing", async () => {
    setup();
    const { DEFAULT_RECEIPT_RULES } = await shared();
    const res = await api("getReceiptRules");
    expect(res).toEqual({
      ok: true,
      data: {
        raw: "",
        hasVariable: false,
        config: DEFAULT_RECEIPT_RULES,
        isDefault: true,
        issues: [],
        forwardEmailAddress: "fwd@example.com",
        forwardName: "Sam",
        query: DEFAULT_QUERY,
      },
    });
    expect(env.LockService.scriptLock.calls).toEqual([]);
  });

  it("reports an empty value cell as defaults with hasVariable true", async () => {
    setup([["ReceiptRules", ""]]);
    const res = await api("getReceiptRules");
    expect(res.data.hasVariable).toBe(true);
    expect(res.data.isDefault).toBe(true);
  });

  it("reports issues, no config and no query for an invalid cell", async () => {
    setup([["ReceiptRules", "{not json"]]);
    const res = await api("getReceiptRules");
    expect(res.ok).toBe(true);
    expect(res.data.raw).toBe("{not json");
    expect(res.data.hasVariable).toBe(true);
    expect(res.data.isDefault).toBe(false);
    expect(res.data.config).toBeNull();
    expect(res.data.query).toBeNull();
    expect(res.data.issues[0].path).toBe("(root)");
    expect(res.data.issues[0].message).toMatch(/^Invalid JSON/);
  });
});

describe("saveReceiptRules", () => {
  async function customRules() {
    const { DEFAULT_RECEIPT_RULES } = await shared();
    const rules = structuredClone(DEFAULT_RECEIPT_RULES);
    rules.senders[2].address = "pay@venmo.example.com";
    return rules;
  }

  it("adds the key, writes pretty JSON under the lock and busts the caches", async () => {
    setup();
    const rules = await customRules();
    const before = await api("getReceiptRules");
    expect(before.data.query).toBe(DEFAULT_QUERY);

    const res = await api("saveReceiptRules", { rules, expectedRaw: "" });

    const text = JSON.stringify(rules, null, 2);
    expect(cell("A8")).toBe("ReceiptRules");
    expect(cell("B8")).toBe(text);
    expect(res.ok).toBe(true);
    expect(res.data).toMatchObject({
      raw: text,
      hasVariable: true,
      config: rules,
      isDefault: false,
      issues: [],
    });
    const query = DEFAULT_QUERY.replace(
      "venmo@venmo.com",
      "pay@venmo.example.com",
    );
    expect(res.data.query).toBe(query);
    expect(env.LockService.scriptLock.calls).toEqual([
      "tryLock:20000",
      "release",
    ]);

    // Same module instance: a following read sees the new value
    const after = await api("getReceiptRules");
    expect(after.data.raw).toBe(text);
    expect(after.data.query).toBe(query);
  });

  it("overwrites an existing value in place", async () => {
    setup([["ReceiptRules", "{not json"]]);
    const rules = await customRules();
    const res = await api("saveReceiptRules", { rules });
    expect(res.ok).toBe(true);
    expect(cell("B8")).toBe(JSON.stringify(rules, null, 2));
    expect(cell("A9")).toBe("");
  });

  it("rejects invalid rules with their issues and writes nothing", async () => {
    setup([["ReceiptRules", ""]]);
    const rules: any = await customRules();
    rules.senders[0].rules[0].subject.pattern = "(";
    const res = await api("saveReceiptRules", { rules });
    expect(res.ok).toBe(false);
    expect(res.error.code).toBe("bad_request");
    expect(res.issues[0].path).toBe("senders[0].rules[0].subject.pattern");
    expect(res.error.message).toBe(
      `ReceiptRules is invalid: ${res.issues[0].path}: ${res.issues[0].message}`,
    );
    expect(cell("B8")).toBe("");
  });

  it("clears the cell for rules null", async () => {
    setup([["ReceiptRules", "{not json"]]);
    const res = await api("saveReceiptRules", { rules: null });
    expect(res.ok).toBe(true);
    expect(cell("B8")).toBe("");
    expect(res.data.isDefault).toBe(true);
    expect(res.data.hasVariable).toBe(true);
    expect(res.data.raw).toBe("");
  });

  it("returns conflict when expectedRaw does not match the cell", async () => {
    setup([["ReceiptRules", "{not json"]]);
    const res = await api("saveReceiptRules", {
      rules: null,
      expectedRaw: "",
    });
    expect(res).toEqual({
      ok: false,
      error: {
        code: "conflict",
        message:
          "The rules on the sheet changed since you loaded them. Reload and try again.",
      },
    });
    expect(cell("B8")).toBe("{not json");
  });

  it("returns busy without writing when the lock is held", async () => {
    setup([["ReceiptRules", "{not json"]]);
    env.LockService.scriptLock.contended = true;
    const res = await api("saveReceiptRules", { rules: null });
    expect(res.error.code).toBe("busy");
    expect(cell("B8")).toBe("{not json");
  });

  it.each([
    [{}, "rules must be an object or null"],
    [{ rules: "x" }, "rules must be an object or null"],
    [{ rules: [] }, "rules must be an object or null"],
    [{ rules: null, expectedRaw: 1 }, "expectedRaw must be a string"],
  ])("rejects %j", async (fields, message) => {
    setup();
    const res = await api("saveReceiptRules", fields);
    expect(res).toEqual({ ok: false, error: { code: "bad_request", message } });
    expect(env.LockService.scriptLock.calls).toEqual([]);
  });
});

describe("a malformed ForwardEmailAddress", () => {
  const MESSAGE =
    /^ForwardEmailAddress variable is invalid: .*Fix it on the Variables sheet/;
  beforeEach(() => {
    setup();
    getEnv()
      .spreadsheet.getSheetByName("Variables")!
      .getRange("B6")
      .setValue("Sam <sam@example.com>");
  });

  it("makes syncReceipts return invalid_config", async () => {
    const res = await api("syncReceipts");
    expect(res.error.code).toBe("invalid_config");
    expect(res.error.message).toMatch(MESSAGE);
  });

  it("makes fetchReceiptEmails (no senders) return invalid_config", async () => {
    const res = await api("fetchReceiptEmails");
    expect(res.error.code).toBe("invalid_config");
    expect(res.error.message).toMatch(MESSAGE);
    expect(env.GmailApp.searches).toEqual([]);
  });

  it("keeps getReceiptRules query null", async () => {
    const res = await api("getReceiptRules");
    expect(res.data.query).toBeNull();
  });
});

describe("fetchReceiptEmails sender limit", () => {
  it("accepts 30 senders and rejects 31", async () => {
    setup();
    const senders = (n: number) =>
      Array.from({ length: n }, (_, i) => `s${i}@example.com`);
    expect((await api("fetchReceiptEmails", { senders: senders(30) })).ok).toBe(
      true,
    );
    const res = await api("fetchReceiptEmails", { senders: senders(31) });
    expect(res.error.message).toBe(
      "senders must be an array of 1-30 email addresses",
    );
  });
});

describe("saveReceiptRules failure paths", () => {
  it("rejects an oversize config as bad_request at (root)", async () => {
    setup();
    const { DEFAULT_RECEIPT_RULES } = await shared();
    const rules = structuredClone(DEFAULT_RECEIPT_RULES);
    const first = rules.senders[0];
    first.rules[0].subject.pattern = `^${"a".repeat(400)}(?<cost>.+) transaction with (?<name>.+)$`;
    for (let i = 0; i < 22; i++)
      rules.senders.push({
        ...structuredClone(first),
        address: `s${i}@example.com`,
        rules: Array.from({ length: 5 }, (_, j) => ({
          ...structuredClone(first.rules[0]),
          id: `r${j}`,
        })),
      });
    const res = await api("saveReceiptRules", { rules });
    expect(res.error.code).toBe("bad_request");
    expect(res.issues[0].path).toBe("(root)");
  });

  it("returns a readable error when Variables!A1:B30 has no free row", async () => {
    setup(Array.from({ length: 23 }, (_, i) => [`x${i}`, i]));
    const { DEFAULT_RECEIPT_RULES } = await shared();
    const res = await api("saveReceiptRules", { rules: DEFAULT_RECEIPT_RULES });
    expect(res.ok).toBe(false);
    expect(res.error.message).toContain("No empty row left");
  });
});

describe("syncReceipts with an invalid ReceiptRules variable", () => {
  it("returns invalid_config", async () => {
    setup([["ReceiptRules", "{not json"]]);
    const res = await api("syncReceipts");
    expect(res.error.code).toBe("invalid_config");
    expect(res.error.message).toMatch(/^ReceiptRules variable is invalid/);
    expect(env.LockService.scriptLock.calls).toEqual([
      "tryLock:20000",
      "release",
    ]);
  });
});
