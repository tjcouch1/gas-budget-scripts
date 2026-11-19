import { describe, expect, it } from "vitest";
import { addVariablesSheet, getEnv, makeThread } from "./gas-mocks";
import {
  applyReceiptRules,
  buildReceiptSearchQuery,
  compileReceiptRules,
  DEFAULT_RECEIPT_RULES,
  type RuleEmail,
} from "@budget/shared";
import {
  FWD,
  receiptEdgeFixtures,
  receiptFixtures,
  type EmailFixture,
  type ReceiptFixture,
} from "./fixtures/receipts";

// The shared engine with the built-in rules must reproduce every fixture's expected parse
// (receipt-rules-plan.md section 7.2). Pure: no GAS mocks needed.

const context = { forwardEmailAddress: "fwd@example.com", forwardName: "Sam" };
const rules = compileReceiptRules(DEFAULT_RECEIPT_RULES, {
  forwardEmailAddress: context.forwardEmailAddress,
});

/** What GAS's GmailMessage would give the engine (the mock's default To is me@example.com) */
function toRuleEmail(email: EmailFixture): RuleEmail {
  return {
    from: email.from,
    to: email.to ?? "me@example.com",
    subject: email.subject,
    plainBody: email.body,
  };
}

const allFixtures = Object.entries({
  ...receiptFixtures,
  ...receiptEdgeFixtures,
}).map(([key, fixture]) => ({ key, ...(fixture as ReceiptFixture) }));

describe("DEFAULT_RECEIPT_RULES reproduce the receipt fixtures", () => {
  it.each(allFixtures)("$key: $description", ({ email, expected }) => {
    const outcome = applyReceiptRules(rules, toRuleEmail(email), context);
    if (expected === null) {
      expect(outcome.isReceipt).toBe(false);
      return;
    }
    expect(outcome.isReceipt).toBe(true);
    expect(outcome.ruleId).not.toBeNull();
    expect(outcome.captures.subject).not.toBeNull();
    for (const [key, value] of Object.entries(expected)) {
      if (key === "costIsNaN")
        expect(Number.isNaN(outcome.receipt.cost)).toBe(true);
      else if (key !== "date")
        expect(outcome.receipt[key as keyof typeof outcome.receipt], key).toBe(
          value,
        );
    }
  });
});

describe("engine error messages match the old parser", () => {
  const run = (email: EmailFixture) =>
    applyReceiptRules(rules, toRuleEmail(email), context);

  it("unknown sender is not a receipt (intentional change; previously threw)", () => {
    const outcome = run({
      from: "Spam <spam@example.org>",
      subject: "Hello",
      body: "b",
    });
    expect(outcome.isReceipt).toBe(false);
    expect(outcome.reason).toBe("unknown-sender");
  });

  it("unparsable sender", () => {
    expect(() => run({ from: "nobody", subject: "s", body: "b" })).toThrow(
      "Could not get email address from nobody",
    );
  });

  it("forwarded email without an original sender", () => {
    expect(() =>
      run({ from: FWD, subject: "Fwd: hi", body: "no forward header here" }),
    ).toThrow(
      "Could not get 'from' email address from forwarded email with subject 'Fwd: hi'",
    );
  });
});

it("the defaults' sender addresses build the pinned Gmail query", () => {
  expect(
    buildReceiptSearchQuery(rules.senderAddresses, "fwd@example.com"),
  ).toBe(
    "(from:(no.reply.alerts@chase.com OR service@paypal.com OR venmo@venmo.com) OR (from:fwd@example.com subject:Fwd)) in:inbox NOT label:receipts NOT label:receipts-cru-reimburse NOT label:receipts-tax-deductible NOT label:receipts-scripted ",
  );
});

describe("ThreadList is an adapter over the engine (parity)", () => {
  it.each(allFixtures)("$key", async ({ email }) => {
    addVariablesSheet(getEnv());
    const { ThreadList } = await import("../src/models/thread-list.model");
    const outcome = applyReceiptRules(rules, toRuleEmail(email), context);
    const list = new ThreadList([makeThread(email)]);
    if (!outcome.isReceipt) {
      expect(list.threadInfos).toEqual([]);
      return;
    }
    expect(list.threadInfos).toHaveLength(1);
    const [info] = list.threadInfos;
    expect(info.errors).toEqual([]);
    expect(info.receiptInfos).toHaveLength(1);
    const [r] = info.receiptInfos;
    expect({
      cost: r.cost,
      name: r.name,
      category: r.category,
      type: r.type,
    }).toEqual(outcome.receipt);
  });

  it("uses the injected rules and context instead of the sheet's", async () => {
    addVariablesSheet(getEnv());
    const { ThreadList } = await import("../src/models/thread-list.model");
    // Same module instance as ThreadList (setup.ts resets modules before each test)
    const shared = await import("@budget/shared");
    const list = new ThreadList(
      [makeThread(receiptFixtures.venmoForwarded.email)],
      shared.compileReceiptRules(shared.DEFAULT_RECEIPT_RULES),
      { forwardEmailAddress: "fwd@example.com", forwardName: "Pat" },
    );
    expect(list.threadInfos[0].receiptInfos[0].type).toBe("Pat Venmo");
  });

  it("matches the sender address case-insensitively (intentional change)", async () => {
    addVariablesSheet(getEnv());
    const { ThreadList } = await import("../src/models/thread-list.model");
    const list = new ThreadList([
      makeThread({
        from: "Chase <No.Reply.Alerts@Chase.com>",
        subject: "You made a $1.00 transaction with Shop",
        body: "",
      }),
    ]);
    expect(list.threadInfos[0].errors).toEqual([]);
    expect(list.threadInfos[0].receiptInfos[0].name).toBe("Shop");
  });

  it("throws out of the constructor before reading any thread when ReceiptRules is invalid", async () => {
    const sheet = addVariablesSheet(getEnv());
    sheet.seed("A8", [["ReceiptRules", "{ nope"]]);
    const { ThreadList } = await import("../src/models/thread-list.model");
    const thread = makeThread(receiptFixtures.chasePurchase.email);
    let read = false;
    thread.getMessages = () => {
      read = true;
      return [];
    };
    expect(() => new ThreadList([thread])).toThrow(
      /^ReceiptRules variable is invalid: /,
    );
    expect(() => new ThreadList([])).toThrow(
      /^ReceiptRules variable is invalid: /,
    );
    expect(read).toBe(false);
  });
});
