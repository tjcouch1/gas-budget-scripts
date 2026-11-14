import { describe, expect, it } from "vitest";
import { DEFAULT_RECEIPT_RULES } from "@budget/shared";
import { addVariablesSheet, getEnv, MockMessage } from "./gas-mocks";

const load = () => import("../src/util/receipt-rules.util");

const PINNED_QUERY =
  "from:(no.reply.alerts@chase.com OR service@paypal.com OR venmo@venmo.com OR fwd@example.com) in:inbox NOT label:receipts NOT label:receipts-cru-reimburse NOT label:receipts-tax-deductible NOT label:receipts-scripted ";

/** Variables sheet with ReceiptRules in row 8 */
function withRules(value: unknown) {
  const sheet = addVariablesSheet(getEnv());
  sheet.seed("A8", [["ReceiptRules", value]]);
  return sheet;
}

const oneSender = {
  ...DEFAULT_RECEIPT_RULES,
  senders: [DEFAULT_RECEIPT_RULES.senders[0]],
};

describe("getCompiledReceiptRules", () => {
  it("uses the built-in defaults when the key is missing", async () => {
    addVariablesSheet(getEnv());
    const { getCompiledReceiptRules, getReceiptRulesRaw } = await load();
    expect(getReceiptRulesRaw()).toBe("");
    expect(getCompiledReceiptRules().config).toEqual(DEFAULT_RECEIPT_RULES);
  });

  it("uses the built-in defaults when the cell is empty or whitespace", async () => {
    withRules("  \n ");
    const { getCompiledReceiptRules } = await load();
    expect(getCompiledReceiptRules().config).toEqual(DEFAULT_RECEIPT_RULES);
  });

  it("compiles valid JSON from the cell", async () => {
    withRules(JSON.stringify(oneSender, null, 2));
    const { getCompiledReceiptRules, getReceiptSenderAddresses } = await load();
    expect(getCompiledReceiptRules().config).toEqual(oneSender);
    expect(getReceiptSenderAddresses()).toEqual(["no.reply.alerts@chase.com"]);
  });

  it("throws the invalid-config message for invalid JSON", async () => {
    withRules("{ nope");
    const { getCompiledReceiptRules } = await load();
    expect(() => getCompiledReceiptRules()).toThrow(
      /^ReceiptRules variable is invalid: .+Fix it in the web app \(Settings > Customize variables\) or clear the cell to use the built-in defaults\.$/,
    );
  });

  it("throws for a number cell", async () => {
    withRules(42);
    const { getCompiledReceiptRules } = await load();
    expect(() => getCompiledReceiptRules()).toThrow(
      "ReceiptRules variable is invalid: ReceiptRules: must be text (JSON).",
    );
  });

  it("memoizes by raw text and recompiles after bustVariablesCache + a cell change", async () => {
    const sheet = withRules("");
    const { getCompiledReceiptRules } = await load();
    const { bustVariablesCache } = await import("../src/util/variables.util");
    const first = getCompiledReceiptRules();
    expect(getCompiledReceiptRules()).toBe(first);

    sheet.getRange("B8").setValue(JSON.stringify(oneSender));
    expect(getCompiledReceiptRules()).toBe(first);
    bustVariablesCache();
    const second = getCompiledReceiptRules();
    expect(second).not.toBe(first);
    expect(second.config).toEqual(oneSender);
  });
});

describe("getCompiledReceiptRules memo and ForwardEmailAddress", () => {
  it("re-validates when ForwardEmailAddress changes with the same raw text", async () => {
    const sheet = withRules("");
    const { getCompiledReceiptRules } = await load();
    const { bustVariablesCache } = await import("../src/util/variables.util");
    expect(() => getCompiledReceiptRules()).not.toThrow();
    sheet.getRange("B6").setValue("no.reply.alerts@chase.com");
    bustVariablesCache();
    expect(() => getCompiledReceiptRules()).toThrow(
      /ReceiptRules variable is invalid: senders\[0\]\.address/,
    );
  });
});

describe("getReceiptSearchQuery with a malformed ForwardEmailAddress", () => {
  it("names the variable and points at the Variables sheet", async () => {
    const sheet = addVariablesSheet(getEnv());
    sheet.getRange("B6").setValue("Sam <sam@example.com>");
    const { getReceiptSearchQuery } = await load();
    expect(() => getReceiptSearchQuery()).toThrow(
      /^ForwardEmailAddress variable is invalid: .*Fix it on the Variables sheet/,
    );
  });
});

describe("getReceiptSearchQuery / getRuleContext", () => {
  it("builds the pinned query from the defaults", async () => {
    addVariablesSheet(getEnv());
    const { getReceiptSearchQuery } = await load();
    expect(getReceiptSearchQuery()).toBe(PINNED_QUERY);
  });

  it("reads the forward address and name from Variables", async () => {
    addVariablesSheet(getEnv());
    const { getRuleContext } = await load();
    expect(getRuleContext()).toEqual({
      forwardEmailAddress: "fwd@example.com",
      forwardName: "Sam",
    });
  });
});

describe("toRuleEmail", () => {
  const message = () => {
    const m = new MockMessage({
      from: "A <a@example.com>",
      to: "b@example.com",
      subject: "S",
      body: "plain",
      html: "<p>html</p>",
    });
    let bodyCalls = 0;
    const getBody = m.getBody.bind(m);
    m.getBody = () => {
      bodyCalls++;
      return getBody();
    };
    return { m, calls: () => bodyCalls };
  };

  it("does not call getBody unless asked", async () => {
    const { toRuleEmail } = await load();
    const { m, calls } = message();
    expect(
      toRuleEmail(m as unknown as GoogleAppsScript.Gmail.GmailMessage, false),
    ).toEqual({
      from: "A <a@example.com>",
      to: "b@example.com",
      subject: "S",
      plainBody: "plain",
    });
    expect(calls()).toBe(0);
  });

  it("adds htmlBody when asked", async () => {
    const { toRuleEmail } = await load();
    const { m, calls } = message();
    expect(
      toRuleEmail(m as unknown as GoogleAppsScript.Gmail.GmailMessage, true)
        .htmlBody,
    ).toBe("<p>html</p>");
    expect(calls()).toBe(1);
  });
});
