import { describe, expect, it } from "vitest";
import {
  applyReceiptRules,
  compileReceiptRules,
  type ReceiptRulesConfig,
  type RuleContext,
  type RuleEmail,
} from "../src/receipt-rules.ts";
import { DEFAULT_RECEIPT_RULES } from "../src/receipt-rules-defaults.ts";

const FWD = "fwd@example.com";
const context: RuleContext = { forwardEmailAddress: FWD, forwardName: "Sam" };
const defaults = compileReceiptRules(DEFAULT_RECEIPT_RULES, {
  forwardEmailAddress: FWD,
});

function email(init: Partial<RuleEmail>): RuleEmail {
  return {
    from: "venmo@venmo.com",
    to: "me@example.com",
    subject: "",
    plainBody: "",
    ...init,
  };
}

function apply(init: Partial<RuleEmail>, ctx = context, rules = defaults) {
  return applyReceiptRules(rules, email(init), ctx);
}

/** A tiny config with one sender, for focused rule tests */
function single(
  sender: Partial<ReceiptRulesConfig["senders"][number]>,
  forwarding: Partial<ReceiptRulesConfig["forwarding"]> = {},
) {
  return compileReceiptRules({
    ...DEFAULT_RECEIPT_RULES,
    forwarding: { ...DEFAULT_RECEIPT_RULES.forwarding, ...forwarding },
    senders: [
      {
        address: "shop@example.com",
        type: "Card",
        typePrefix: "fixed",
        forwardable: true,
        rules: [],
        ...sender,
      },
    ],
  });
}

const VENMO_PAID_BODY = "You paid\nJane Doe\n$15.00\nPizza night\nmore\n";

function forwardedVenmo(header: string) {
  return email({
    from: `Sam Example <${FWD}>`,
    subject: "Fwd: You paid Jane Doe $15.00",
    plainBody: header + VENMO_PAID_BODY,
  });
}

describe("applyReceiptRules: direct emails", () => {
  it("returns the full outcome with captures", () => {
    const outcome = apply({
      from: "Chase <no.reply.alerts@chase.com>",
      subject: "You made a $12.34 transaction with Shop",
    });
    expect(outcome).toEqual({
      senderAddress: "no.reply.alerts@chase.com",
      forwarded: false,
      ruleId: "chase-purchase",
      receipt: {
        cost: 12.34,
        name: "Shop",
        category: undefined,
        type: "Credit",
      },
      isReceipt: true,
      captures: { subject: { cost: "12.34", name: "Shop" }, body: null },
      warnings: [],
    });
  });

  it("compares the sender address case-insensitively (intentional change)", () => {
    const outcome = apply({
      from: "Chase <No.Reply.Alerts@CHASE.com>",
      subject: "You made a $1.00 transaction with Shop",
    });
    expect(outcome.senderAddress).toBe("no.reply.alerts@chase.com");
    expect(outcome.receipt.name).toBe("Shop");
  });

  it("throws for an unknown sender with the lower-cased address", () => {
    expect(() => apply({ from: "Spam <Spam@Example.org>" })).toThrow(
      new Error("No function to handle receipt message from spam@example.org"),
    );
  });

  it("throws when the sender cannot be parsed", () => {
    expect(() => apply({ from: "nobody" })).toThrow(
      new Error("Could not get email address from nobody"),
    );
  });

  it("known sender with no matching subject: ruleId null, type still set", () => {
    const outcome = apply({
      from: "service@paypal.com",
      subject: "Update to your account",
    });
    expect(outcome.ruleId).toBeNull();
    expect(outcome.isReceipt).toBe(false);
    expect(outcome.receipt).toEqual({
      cost: undefined,
      name: undefined,
      category: undefined,
      type: "TJ Paypal",
    });
    expect(outcome.captures).toEqual({ subject: null, body: null });
  });

  it("uses nameFallback when the body does not match, and a failed body gives null captures", () => {
    const outcome = apply({
      subject: "You paid Jane Doe $15.00",
      plainBody: "",
    });
    expect(outcome.receipt.name).toBe("Jane Doe");
    expect(outcome.captures.body).toBeNull();
  });

  it("negates the cost", () => {
    expect(
      apply({ subject: "Alice paid you $40.00", plainBody: "" }).receipt.cost,
    ).toBe(-40);
  });

  it("keeps parseFloat semantics: '1,234.56' => 1, 'abc' => NaN (not a receipt without a name)", () => {
    const rules = single({
      rules: [
        {
          id: "r",
          subject: { pattern: "^Paid \\$(?<cost>.*)$" },
          cost: { source: "subject", group: "cost" },
        },
      ],
    });
    const run = (subject: string) =>
      applyReceiptRules(
        rules,
        email({ from: "shop@example.com", subject }),
        context,
      );
    expect(run("Paid $1,234.56").receipt.cost).toBe(1);
    const nan = run("Paid $abc");
    expect(Number.isNaN(nan.receipt.cost)).toBe(true);
    expect(nan.isReceipt).toBe(false);
    // an empty group gives no cost
    expect(run("Paid $").receipt.cost).toBeUndefined();
  });

  it("zero cost (-0 after negate) with no name is not a receipt", () => {
    const outcome = apply({
      from: "no.reply.alerts@chase.com",
      subject: "You have a $0.00 credit pending on your credit card",
      plainBody: "no merchant line",
    });
    expect(Object.is(outcome.receipt.cost, -0)).toBe(true);
    expect(outcome.isReceipt).toBe(false);
  });

  it("sets category only when the outcome is a receipt", () => {
    const gas = "You used your card at a gas station";
    const hit = apply({
      from: "no.reply.alerts@chase.com",
      subject: gas,
      plainBody: "Hi\nMerchant   SHELL  \n",
    });
    expect(hit.receipt).toMatchObject({ name: "SHELL (Gas)", category: "Gas" });
    const miss = apply({
      from: "no.reply.alerts@chase.com",
      subject: gas,
      plainBody: "nothing",
    });
    expect(miss.ruleId).toBe("chase-gas");
    expect(miss.isReceipt).toBe(false);
    expect(miss.receipt.category).toBeUndefined();
  });
});

describe("applyReceiptRules: type prefix", () => {
  const paypal = {
    from: "service@paypal.com",
    subject: "Update",
  };

  it("uses ForwardName when sent To the forward address (case-insensitive)", () => {
    expect(apply({ ...paypal, to: "Sam <FWD@example.com>" }).receipt.type).toBe(
      "Sam Paypal",
    );
  });

  it("falls back to defaultTypePrefix for other or unparsable To addresses", () => {
    expect(apply({ ...paypal, to: "other@example.com" }).receipt.type).toBe(
      "TJ Paypal",
    );
    expect(
      apply({ ...paypal, to: "undisclosed-recipients" }).receipt.type,
    ).toBe("TJ Paypal");
  });

  it("falls through to defaultTypePrefix when ForwardName is empty", () => {
    const ctx = { forwardEmailAddress: FWD, forwardName: "" };
    expect(apply({ ...paypal, to: FWD }, ctx).receipt.type).toBe("TJ Paypal");
    // forwarded email with empty ForwardName and a To that is not the forwarder
    const outcome = applyReceiptRules(
      defaults,
      forwardedVenmo(
        "---------- Forwarded message ---------\nFrom: venmo@venmo.com\nDate: d\nSubject: s\nTo: t\n\n",
      ),
      ctx,
    );
    expect(outcome.receipt.type).toBe("TJ Venmo");
  });
});

describe("applyReceiptRules: forwarded emails", () => {
  it.each([
    [
      "Forwarded message marker (skip 5) + English From",
      "FYI\n\n---------- Forwarded message ---------\nFrom: Venmo <venmo@venmo.com>\nDate: Mon\nSubject: s\nTo: Sam <fwd@example.com>\n\n",
    ],
    [
      "Begin forwarded message marker (skip 6) + *From:*",
      "Hi\n\nBegin forwarded message:\n\n*From:* Venmo <venmo@venmo.com>\n*Date:* d\n*To:* t\n*Subject:* s\n\n",
    ],
    [
      "Japanese marker (skip 6) + Japanese From",
      "メモ\n\n転送されたメッセージ:\n\n*差出人:* Venmo <venmo@venmo.com>\n*日時:* d\n*宛先:* t\n*件名:* s\n\n",
    ],
  ])("%s", (_title, header) => {
    const outcome = applyReceiptRules(
      defaults,
      forwardedVenmo(header),
      context,
    );
    expect(outcome).toMatchObject({
      senderAddress: "venmo@venmo.com",
      forwarded: true,
      ruleId: "venmo-paid",
      receipt: { cost: 15, name: "Jane Doe - Pizza night", type: "Sam Venmo" },
      isReceipt: true,
    });
  });

  it("uses a bare original address as is, lower-cased", () => {
    const outcome = applyReceiptRules(
      defaults,
      forwardedVenmo(
        "---------- Forwarded message ---------\nFrom: Venmo@Venmo.com\nDate: d\nSubject: s\nTo: t\n\n",
      ),
      context,
    );
    expect(outcome.senderAddress).toBe("venmo@venmo.com");
    expect(outcome.isReceipt).toBe(true);
  });

  it("detects the forwarder case-insensitively", () => {
    const e = forwardedVenmo(
      "---------- Forwarded message ---------\nFrom: venmo@venmo.com\nDate: d\nSubject: s\nTo: t\n\n",
    );
    const outcome = applyReceiptRules(
      defaults,
      { ...e, from: "FWD@Example.com" },
      context,
    );
    expect(outcome.forwarded).toBe(true);
  });

  it("requires the subject prefix", () => {
    const e = forwardedVenmo(
      "---------- Forwarded message ---------\nFrom: venmo@venmo.com\nDate: d\nSubject: s\nTo: t\n\n",
    );
    const outcome = applyReceiptRules(
      defaults,
      { ...e, subject: "You paid Jane Doe $15.00" },
      context,
    );
    expect(outcome.ruleId).toBeNull();
    expect(outcome.receipt.type).toBe("Sam Venmo");
  });

  it("supports configured subject prefixes, escaped as literal text", () => {
    const rules = single(
      {
        typePrefix: "recipient",
        rules: [
          {
            id: "r",
            subject: { pattern: "^Paid (?<name>.+)$" },
            name: "${subject.name}",
          },
        ],
      },
      { subjectPrefixes: ["Fwd: ", "[FW] (x)? "] },
    );
    const run = (subject: string) =>
      applyReceiptRules(
        rules,
        email({
          from: FWD,
          subject,
          plainBody:
            "---------- Forwarded message ---------\nFrom: shop@example.com\n",
        }),
        context,
      );
    expect(run("[FW] (x)? Paid Bob").receipt).toMatchObject({
      name: "Bob",
      type: "Sam Card",
    });
    expect(run("Fwd: Paid Bob").receipt.name).toBe("Bob");
    // "?" and "()" are literal, not regex syntax
    expect(run("[FW]  Paid Bob").ruleId).toBeNull();
    expect(run("F Paid Bob").ruleId).toBeNull();
  });

  it("throws when the original sender cannot be found", () => {
    expect(() =>
      apply({ from: FWD, subject: "Fwd: hi", plainBody: "no header" }),
    ).toThrow(
      new Error(
        "Could not get 'from' email address from forwarded email with subject 'Fwd: hi'",
      ),
    );
  });

  it("not-forwardable and unknown original senders give an empty non-receipt", () => {
    const empty = {
      senderAddress: null,
      forwarded: true,
      ruleId: null,
      receipt: {
        cost: undefined,
        name: undefined,
        category: undefined,
        type: undefined,
      },
      isReceipt: false,
      captures: { subject: null, body: null },
      warnings: [],
    };
    expect(
      apply({
        from: FWD,
        subject: "Fwd: You made a $1.00 transaction with X",
        plainBody:
          "---------- Forwarded message ---------\nFrom: Chase <no.reply.alerts@chase.com>\n",
      }),
    ).toEqual(empty);
    expect(
      apply({
        from: FWD,
        subject: "Fwd: Hello",
        plainBody:
          "---------- Forwarded message ---------\nFrom: Someone <x@example.com>\n",
      }),
    ).toEqual(empty);
  });

  it("uses the whole body when no marker is found", () => {
    const rules = single(
      {
        rules: [
          {
            id: "r",
            subject: { pattern: "^S$" },
            body: { source: "plainBody", pattern: "^FYI$", flags: "m" },
            name: "found",
          },
        ],
      },
      { bodyMarkers: [] },
    );
    const outcome = applyReceiptRules(
      rules,
      email({
        from: FWD,
        subject: "Fwd: S",
        plainBody: "FYI\nForwarded message\nFrom: shop@example.com\n",
      }),
      context,
    );
    expect(outcome.receipt.name).toBe("found");
  });
});

describe("applyReceiptRules: htmlBody", () => {
  const rules = single({
    rules: [
      {
        id: "html",
        subject: { pattern: "^Order$" },
        body: { source: "htmlBody", pattern: "<b>(?<name>[^<]+)</b>" },
        name: "${body.name}",
      },
    ],
  });

  it("matches the HTML body when provided", () => {
    const outcome = applyReceiptRules(
      rules,
      email({
        from: "shop@example.com",
        subject: "Order",
        htmlBody: "<p><b>Widget</b></p>",
      }),
      context,
    );
    expect(outcome.receipt.name).toBe("Widget");
    expect(outcome.captures.body).toEqual({ name: "Widget" });
    expect(outcome.warnings).toEqual([]);
  });

  it("warns when the HTML body is missing", () => {
    const outcome = applyReceiptRules(
      rules,
      email({ from: "shop@example.com", subject: "Order" }),
      context,
    );
    expect(outcome.receipt.name).toBeUndefined();
    expect(outcome.captures.body).toBeNull();
    expect(outcome.warnings).toEqual([
      "Rule html needs the HTML body but it was not provided",
    ]);
  });

  it("warns that forwarded HTML bodies are not unwrapped", () => {
    const outcome = applyReceiptRules(
      rules,
      email({
        from: FWD,
        subject: "Fwd: Order",
        plainBody:
          "---------- Forwarded message ---------\nFrom: shop@example.com\n",
        htmlBody: "<b>Widget</b>",
      }),
      context,
    );
    expect(outcome.receipt.name).toBe("Widget");
    expect(outcome.warnings).toEqual([
      "forwarded HTML bodies are not unwrapped",
    ]);
  });
});

describe("forward subject with top-level alternation", () => {
  it("requires the forward prefix for every alternative", () => {
    const config = structuredClone(DEFAULT_RECEIPT_RULES);
    config.senders[2].rules = [
      { id: "alt", subject: { pattern: "^a|b" }, name: "Alt" },
    ];
    const rules = compileReceiptRules(config, { forwardEmailAddress: FWD });
    const fwd = (subject: string) =>
      apply(
        {
          from: FWD,
          subject,
          plainBody:
            "---------- Forwarded message ---------\nFrom: Venmo <venmo@venmo.com>\nDate: x\nSubject: y\nTo: z\n\nbody",
        },
        context,
        rules,
      );
    expect(fwd("Fwd: b").receipt.name).toBe("Alt");
    expect(fwd("b").receipt.name).toBeUndefined();
  });
});
