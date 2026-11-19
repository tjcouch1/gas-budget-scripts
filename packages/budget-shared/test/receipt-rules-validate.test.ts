import { describe, expect, it } from "vitest";
import {
  compileReceiptRules,
  expandTemplate,
  extractEmailAddress,
  parseReceiptRulesText,
  validateReceiptRulesConfig,
  type ReceiptRulesConfig,
} from "../src/receipt-rules.ts";
import {
  DEFAULT_RECEIPT_RULES,
  MAX_RECEIPT_RULES_CHARS,
  RECEIPT_RULES_NOTE,
} from "../src/receipt-rules-defaults.ts";

/** A mutable deep copy of the defaults, typed loosely so tests can break it */
// eslint-disable-next-line @typescript-eslint/no-explicit-any -- tests deliberately build invalid configs
function defaults(): any {
  return structuredClone(DEFAULT_RECEIPT_RULES);
}

/** The issue paths a value produces */
function issuePaths(value: unknown, forwardEmailAddress?: string): string[] {
  const result = validateReceiptRulesConfig(value, { forwardEmailAddress });
  return result.ok ? [] : result.issues.map((i) => i.path);
}

function issueFor(value: unknown, path: string, forwardEmailAddress?: string) {
  const result = validateReceiptRulesConfig(value, { forwardEmailAddress });
  if (result.ok) throw new Error("expected issues");
  return result.issues.find((i) => i.path === path);
}

describe("DEFAULT_RECEIPT_RULES", () => {
  it("deep-equals the plan's JSON (section 2.4, plus forwarding.subjectPrefixes)", () => {
    const expected = {
      version: 1,
      defaultTypePrefix: "TJ",
      forwarding: {
        subjectPrefixes: ["Fwd: "],
        fromPatterns: [
          {
            pattern:
              "[Ff]orwarded message.*\\s*\\r?\\n\\r?\\n?\\*?From:\\*?[ \\u00a0](?<email>.+)\\r?\\n",
          },
          {
            pattern:
              "転送されたメッセージ.*\\s*\\r?\\n\\r?\\n?\\*?差出人:\\*?[ \\u00a0](?<email>.+)\\r?\\n",
          },
        ],
        bodyMarkers: [
          { marker: "---------- Forwarded message ---------", skipLines: 5 },
          { marker: "Begin forwarded message:", skipLines: 6 },
          { marker: "転送されたメッセージ:", skipLines: 6 },
        ],
      },
      senders: [
        {
          address: "no.reply.alerts@chase.com",
          type: "Credit",
          typePrefix: "fixed",
          forwardable: false,
          rules: [
            {
              id: "chase-purchase",
              subject: {
                pattern:
                  "^You made a \\$(?<cost>.+) transaction with (?<name>.+)$",
              },
              name: "${subject.name}",
              cost: { source: "subject", group: "cost" },
            },
            {
              id: "chase-refund",
              subject: {
                pattern:
                  "^You have a \\$(?<cost>.+) credit pending on your credit card$",
              },
              body: {
                source: "plainBody",
                pattern: "\\r?\\nMerchant\\s+(?<name>.*\\S)\\s*\\r?\\n",
              },
              name: "${body.name}",
              cost: { source: "subject", group: "cost", negate: true },
            },
            {
              id: "chase-gas",
              subject: { pattern: "^You used your card at a gas station$" },
              body: {
                source: "plainBody",
                pattern: "\\r?\\nMerchant\\s+(?<name>.*\\S)\\s*\\r?\\n",
              },
              name: "${body.name} (Gas)",
              category: "Gas",
            },
          ],
        },
        {
          address: "service@paypal.com",
          type: "Paypal",
          typePrefix: "recipient",
          forwardable: true,
          rules: [
            {
              id: "paypal-sent",
              subject: {
                pattern: "^You sent a \\$(?<cost>.+)[\\s\\u00a0]USD payment$",
              },
              body: {
                source: "plainBody",
                pattern:
                  "You sent \\$.+ to (?<name>.+)\\r?\\n\\r?\\n.*\\r?\\n.*\\r?\\n\\r?\\n(?<details>.+)\\r?\\n",
              },
              name: "${body.name} - ${body.details}",
              cost: { source: "subject", group: "cost" },
            },
            {
              id: "paypal-received",
              subject: { pattern: "^Money is waiting for you$" },
              body: {
                source: "plainBody",
                pattern:
                  "Accept your \\$(?<cost>.+)[\\s\\u00a0]USD from (?<name>.+)\\r?\\n\\r?\\n.*\\r?\\n.*\\r?\\n\\r?\\n(?<details>.+)\\r?\\n",
              },
              name: "${body.name} - ${body.details}",
              cost: { source: "body", group: "cost", negate: true },
            },
            {
              id: "paypal-sent-you",
              subject: {
                pattern: "^(?<name>.+) sent you \\$(?<cost>.+)[\\s\\u00a0]USD$",
              },
              body: {
                source: "plainBody",
                pattern:
                  "[Nn]ote from [^\\r\\n]+\\r?\\n(?:[ \\t]*\\r?\\n)*(?<details>[^\\r\\n]*\\S)",
              },
              name: "${subject.name} - ${body.details}",
              nameFallback: "${subject.name}",
              cost: { source: "subject", group: "cost", negate: true },
            },
          ],
        },
        {
          address: "venmo@venmo.com",
          type: "Venmo",
          typePrefix: "recipient",
          forwardable: true,
          rules: [
            {
              id: "venmo-paid",
              subject: { pattern: "^You paid (?<name>.+) \\$(?<cost>.+)$" },
              body: {
                source: "plainBody",
                pattern: "paid\\s*.*\\s*\\s*.*\\s*\\r?\\n(?<details>.+)\\r?\\n",
              },
              name: "${subject.name} - ${body.details}",
              nameFallback: "${subject.name}",
              cost: { source: "subject", group: "cost" },
            },
            {
              id: "venmo-charge",
              subject: {
                pattern:
                  "^You completed (?<name>.+)'s \\$(?<cost>.+) charge request$",
              },
              body: {
                source: "plainBody",
                pattern:
                  "charged\\s*.*\\s*\\s*.*\\s*\\r?\\n(?<details>.+)\\r?\\n",
              },
              name: "${subject.name} - ${body.details}",
              nameFallback: "${subject.name}",
              cost: { source: "subject", group: "cost" },
            },
            {
              id: "venmo-received",
              subject: { pattern: "^(?<name>.+) paid you \\$(?<cost>.+)$" },
              body: {
                source: "plainBody",
                pattern: "paid\\s*.*\\s*\\s*.*\\s*\\r?\\n(?<details>.+)\\r?\\n",
              },
              name: "${subject.name} - ${body.details}",
              nameFallback: "${subject.name}",
              cost: { source: "subject", group: "cost", negate: true },
            },
          ],
        },
      ],
    };
    expect(DEFAULT_RECEIPT_RULES).toEqual(expected);
  });

  it("validates ok, also against the test ForwardEmailAddress", () => {
    expect(validateReceiptRulesConfig(DEFAULT_RECEIPT_RULES)).toEqual({
      ok: true,
      config: DEFAULT_RECEIPT_RULES,
    });
    expect(
      validateReceiptRulesConfig(DEFAULT_RECEIPT_RULES, {
        forwardEmailAddress: "fwd@example.com",
      }).ok,
    ).toBe(true);
  });

  it("serializes well under the size limit", () => {
    expect(JSON.stringify(DEFAULT_RECEIPT_RULES, null, 2).length).toBeLessThan(
      MAX_RECEIPT_RULES_CHARS,
    );
    expect(MAX_RECEIPT_RULES_CHARS).toBe(40_000);
  });

  it("has a plain-text cell note of at most 2,500 characters", () => {
    expect(RECEIPT_RULES_NOTE.length).toBeGreaterThan(200);
    expect(RECEIPT_RULES_NOTE.length).toBeLessThanOrEqual(2500);
    expect(RECEIPT_RULES_NOTE).toContain("Customize variables");
    expect(RECEIPT_RULES_NOTE).toContain("subjectPrefixes");
    expect(RECEIPT_RULES_NOTE).toContain("Reset to built-in defaults");
    expect(RECEIPT_RULES_NOTE).toContain("unknown senders are not receipts");
  });
});

describe("parseReceiptRulesText", () => {
  it.each([[""], ["  \n "], [null], [undefined]])(
    "returns the defaults for %j",
    (text) => {
      expect(parseReceiptRulesText(text)).toEqual({
        ok: true,
        config: DEFAULT_RECEIPT_RULES,
        isDefault: true,
      });
    },
  );

  it("parses and validates JSON text", () => {
    const text = JSON.stringify(DEFAULT_RECEIPT_RULES, null, 2);
    const result = parseReceiptRulesText(text);
    expect(result).toEqual({ ok: true, config: DEFAULT_RECEIPT_RULES });
  });

  it("reports a JSON syntax error at the root", () => {
    const result = parseReceiptRulesText("{ nope");
    expect(result.ok).toBe(false);
    if (result.ok) return;
    expect(result.issues).toHaveLength(1);
    expect(result.issues[0].path).toBe("(root)");
    expect(result.issues[0].message).toMatch(/^Invalid JSON: /);
  });

  it("passes forwardEmailAddress to validation", () => {
    const result = parseReceiptRulesText(
      JSON.stringify(DEFAULT_RECEIPT_RULES),
      { forwardEmailAddress: "Service@PayPal.com" },
    );
    expect(result.ok).toBe(false);
  });
});

describe("validateReceiptRulesConfig issues (exact paths)", () => {
  it("rejects a non-object root", () => {
    expect(issuePaths(null)).toEqual(["(root)"]);
    expect(issuePaths([])).toEqual(["(root)"]);
  });

  it("bad version", () => {
    const c = defaults();
    c.version = 2;
    expect(issuePaths(c)).toEqual(["version"]);
  });

  it("bad flags 'gi'", () => {
    const c = defaults();
    c.senders[0].rules[0].subject.flags = "gi";
    expect(issuePaths(c)).toEqual(["senders[0].rules[0].subject.flags"]);
  });

  it("repeated flag 'ii' and accepts 'imsu'", () => {
    const c = defaults();
    c.senders[0].rules[0].subject.flags = "ii";
    expect(issuePaths(c)).toEqual(["senders[0].rules[0].subject.flags"]);
    c.senders[0].rules[0].subject.flags = "imsu";
    expect(issuePaths(c)).toEqual([]);
  });

  it("invalid regex '(' reports the RegExp error message", () => {
    const c = defaults();
    c.senders[1].rules[0].body.pattern = "(";
    const issue = issueFor(c, "senders[1].rules[0].body.pattern");
    expect(issue?.message).toMatch(/Invalid regular expression/);
  });

  it("missing (?<email> in a from pattern", () => {
    const c = defaults();
    c.forwarding.fromPatterns[1].pattern = "From: (.+)";
    expect(issuePaths(c)).toEqual(["forwarding.fromPatterns[1].pattern"]);
  });

  it("template referencing ${body.x} without a body", () => {
    const c = defaults();
    c.senders[0].rules[0].name = "${body.name}";
    expect(issuePaths(c)).toEqual(["senders[0].rules[0].name"]);
  });

  it("template referencing an undeclared group or a bad placeholder", () => {
    const c = defaults();
    c.senders[2].rules[0].nameFallback = "${subject.nope}";
    c.senders[2].rules[1].name = "${name}";
    expect(issuePaths(c)).toEqual([
      "senders[2].rules[0].nameFallback",
      "senders[2].rules[1].name",
    ]);
  });

  it("template whose literal start is a formula character", () => {
    const c = defaults();
    c.senders[0].rules[0].name = "=${subject.name}";
    expect(issuePaths(c)).toEqual(["senders[0].rules[0].name"]);
  });

  it("cost group not declared, and cost from a missing body", () => {
    const c = defaults();
    c.senders[0].rules[0].cost.group = "amount";
    c.senders[0].rules[0].cost.source = "subject";
    c.senders[0].rules[0].name = undefined;
    delete c.senders[0].rules[0].name;
    expect(issuePaths(c)).toEqual(["senders[0].rules[0].cost.group"]);
    c.senders[0].rules[0].cost = { source: "body", group: "cost" };
    expect(issuePaths(c)).toEqual(["senders[0].rules[0].cost.source"]);
  });

  it("duplicate address (case-insensitive) is reported on the later sender", () => {
    const c = defaults();
    c.senders[2].address = "Service@PayPal.com";
    expect(issuePaths(c)).toEqual(["senders[2].address"]);
  });

  it("address equal to ForwardEmailAddress (case-insensitive)", () => {
    expect(issuePaths(DEFAULT_RECEIPT_RULES, "VENMO@venmo.com")).toEqual([
      "senders[2].address",
    ]);
  });

  it("address that is not a strict email address", () => {
    const c = defaults();
    c.senders[0].address = "a@b.com OR x@y.com";
    expect(issuePaths(c)).toEqual(["senders[0].address"]);
  });

  it("unknown keys anywhere", () => {
    const c = defaults();
    c.extra = 1;
    c.forwarding.bodyMarkers[0].skip = 1;
    c.senders[0].rules[0].nmae = "typo";
    c.senders[1].rules[0].body.sorce = "x";
    const result = validateReceiptRulesConfig(c);
    expect(result.ok).toBe(false);
    if (result.ok) return;
    expect(result.issues).toEqual([
      { path: "extra", message: "unknown key" },
      { path: "forwarding.bodyMarkers[0].skip", message: "unknown key" },
      { path: "senders[0].rules[0].nmae", message: "unknown key" },
      { path: "senders[1].rules[0].body.sorce", message: "unknown key" },
    ]);
  });

  it("oversize JSON", () => {
    const c = defaults();
    c.senders[0].rules[0].subject.pattern = `^${"a".repeat(400)}(?<cost>.+) transaction with (?<name>.+)$`;
    for (let i = 0; i < 25 - 3; i++)
      c.senders.push({
        ...structuredClone(c.senders[0]),
        address: `s${i}@example.com`,
        rules: Array.from({ length: 5 }, (_, j) => ({
          ...structuredClone(c.senders[0].rules[0]),
          id: `r${j}`,
        })),
      });
    expect(JSON.stringify(c, null, 2).length).toBeGreaterThan(
      MAX_RECEIPT_RULES_CHARS,
    );
    expect(issuePaths(c)).toEqual(["(root)"]);
  });

  it("unsafe type '=x', a '|' in type, and an unsafe defaultTypePrefix", () => {
    const c = defaults();
    c.senders[0].type = "=x";
    c.senders[1].type = "Pay|pal";
    c.defaultTypePrefix = "+TJ";
    expect(issuePaths(c)).toEqual([
      "defaultTypePrefix",
      "senders[0].type",
      "senders[1].type",
    ]);
  });

  it("typePrefix, forwardable, rule id and duplicate id, category", () => {
    const c = defaults();
    c.senders[0].typePrefix = "other";
    c.senders[0].forwardable = "yes";
    c.senders[0].rules[1].id = "bad id";
    c.senders[0].rules[2].id = "chase-purchase";
    c.senders[0].rules[2].category = "@Gas";
    expect(issuePaths(c)).toEqual([
      "senders[0].typePrefix",
      "senders[0].forwardable",
      "senders[0].rules[1].id",
      "senders[0].rules[2].id",
      "senders[0].rules[2].category",
    ]);
  });

  it("body source, cost negate, missing subject", () => {
    const c = defaults();
    c.senders[0].rules[1].body.source = "text";
    c.senders[0].rules[1].cost.negate = "yes";
    delete c.senders[0].rules[2].subject;
    expect(issuePaths(c)).toEqual([
      "senders[0].rules[1].body.source",
      "senders[0].rules[1].cost.negate",
      "senders[0].rules[2].subject",
    ]);
  });

  it("array sizes: senders, rules, fromPatterns, bodyMarkers", () => {
    const c = defaults();
    c.senders = [];
    c.forwarding.fromPatterns = [];
    c.forwarding.bodyMarkers = Array.from({ length: 11 }, () => ({
      marker: "m",
      skipLines: 0,
    }));
    expect(issuePaths(c)).toEqual([
      "forwarding.fromPatterns",
      "forwarding.bodyMarkers",
      "senders",
    ]);
    const d = defaults();
    d.senders[0].rules = Array.from({ length: 26 }, (_, i) => ({
      id: `r${i}`,
      subject: { pattern: "x" },
    }));
    expect(issuePaths(d)).toEqual(["senders[0].rules"]);
  });

  it("body markers: empty marker and bad skipLines", () => {
    const c = defaults();
    c.forwarding.bodyMarkers[0].marker = " ";
    c.forwarding.bodyMarkers[1].skipLines = 1.5;
    c.forwarding.bodyMarkers[2].skipLines = 21;
    expect(issuePaths(c)).toEqual([
      "forwarding.bodyMarkers[0].marker",
      "forwarding.bodyMarkers[1].skipLines",
      "forwarding.bodyMarkers[2].skipLines",
    ]);
  });

  it("forwarding.subjectPrefixes: optional, 1-5 entries of 1-20 characters", () => {
    const c = defaults();
    delete c.forwarding.subjectPrefixes;
    expect(issuePaths(c)).toEqual([]);
    c.forwarding.subjectPrefixes = [];
    expect(issuePaths(c)).toEqual(["forwarding.subjectPrefixes"]);
    c.forwarding.subjectPrefixes = ["a", "b", "c", "d", "e", "f"];
    expect(issuePaths(c)).toEqual(["forwarding.subjectPrefixes"]);
    c.forwarding.subjectPrefixes = ["Fwd: ", "", "x".repeat(21), 3];
    expect(issuePaths(c)).toEqual([
      "forwarding.subjectPrefixes[1]",
      "forwarding.subjectPrefixes[2]",
      "forwarding.subjectPrefixes[3]",
    ]);
  });

  it("forwarding.subjectPrefixes: each must be letters/digits plus an optional ':' and spaces (Gmail query token)", () => {
    const c = defaults();
    c.forwarding.subjectPrefixes = ["Fwd: ", "FW:", "転送: ", "TR"];
    expect(issuePaths(c)).toEqual([]);
    c.forwarding.subjectPrefixes = ["[FW] ", "Re Fwd: ", ": "];
    const result = validateReceiptRulesConfig(c);
    expect(result.ok ? [] : result.issues).toEqual(
      [0, 1, 2].map((i) => ({
        path: `forwarding.subjectPrefixes[${i}]`,
        message:
          'must be letters or digits, optionally followed by ":" and spaces (e.g. "Fwd: ")',
      })),
    );
  });

  it("pattern length limit (no cascading group issues) and template length limit", () => {
    const c = defaults();
    c.senders[0].rules[0].subject.pattern = "a".repeat(501);
    c.senders[0].rules[1].name = `x${"${body.name}".repeat(17)}`;
    expect(issuePaths(c)).toEqual([
      "senders[0].rules[0].subject.pattern",
      "senders[0].rules[1].name",
    ]);
  });

  it("collects at most 50 issues", () => {
    const c = defaults();
    for (let i = 0; i < 60; i++) c[`k${i}`] = i;
    const result = validateReceiptRulesConfig(c);
    expect(result.ok).toBe(false);
    if (!result.ok) expect(result.issues).toHaveLength(50);
  });
});

describe("compileReceiptRules", () => {
  it("compiles the defaults", () => {
    const compiled = compileReceiptRules(DEFAULT_RECEIPT_RULES, {
      forwardEmailAddress: "fwd@example.com",
    });
    expect(compiled.config).toBe(DEFAULT_RECEIPT_RULES);
    expect(compiled.usesHtmlBody).toBe(false);
    expect(compiled.senderAddresses).toEqual([
      "no.reply.alerts@chase.com",
      "service@paypal.com",
      "venmo@venmo.com",
    ]);
  });

  it("lower-cases sender addresses and detects htmlBody use", () => {
    const c = defaults();
    c.senders[0].address = "No.Reply.Alerts@Chase.com";
    c.senders[0].rules[1].body.source = "htmlBody";
    const compiled = compileReceiptRules(c as ReceiptRulesConfig);
    expect(compiled.senderAddresses[0]).toBe("no.reply.alerts@chase.com");
    expect(compiled.usesHtmlBody).toBe(true);
  });

  it("throws the exact invalid-config message for the first issue", () => {
    expect(() =>
      compileReceiptRules(DEFAULT_RECEIPT_RULES, {
        forwardEmailAddress: "service@paypal.com",
      }),
    ).toThrow(
      new Error(
        "ReceiptRules variable is invalid: senders[1].address: must not be the ForwardEmailAddress (forwarded emails are handled by the forwarding settings). Fix it in the web app (Settings > Customize variables) or clear the cell to use the built-in defaults.",
      ),
    );
  });
});

describe("expandTemplate", () => {
  const groups = {
    subject: { name: "Shop", cost: "1.00", none: undefined },
    body: null,
  };

  it("expands placeholders and leaves literal $ untouched", () => {
    expect(expandTemplate("$5 ${subject.name} $& $1 ${x", groups)).toBe(
      "$5 Shop $& $1 ${x",
    );
  });

  it("does not interpret replacement patterns in group text", () => {
    expect(
      expandTemplate("${subject.name}", {
        subject: { name: "$& $1 $$" },
        body: null,
      }),
    ).toBe("$& $1 $$");
  });

  it("returns undefined when any referenced group is missing", () => {
    expect(expandTemplate("${subject.none}", groups)).toBeUndefined();
    expect(expandTemplate("${subject.nope}", groups)).toBeUndefined();
    expect(
      expandTemplate("${subject.name} - ${body.details}", groups),
    ).toBeUndefined();
  });

  it("keeps an empty group (participated, matched nothing)", () => {
    expect(
      expandTemplate("[${subject.name}]", {
        subject: { name: "" },
        body: null,
      }),
    ).toBe("[]");
  });
});

describe("extractEmailAddress", () => {
  it("extracts from a formal address and returns a bare address", () => {
    expect(extractEmailAddress("Chase <no.reply.alerts@chase.com>")).toBe(
      "no.reply.alerts@chase.com",
    );
    expect(extractEmailAddress("venmo@venmo.com")).toBe("venmo@venmo.com");
  });

  it("prefers the address inside <...> and strips quotes", () => {
    for (const [from, address] of [
      ['"service@paypal.com" <service@paypal.com>', "service@paypal.com"],
      ['"service@paypal.com"', "service@paypal.com"],
      ["'service@paypal.com'", "service@paypal.com"],
      ["<service@paypal.com>", "service@paypal.com"],
      ["Pat <pat@example.com> ", "pat@example.com"],
      ["old@example.com <new@example.com>", "new@example.com"],
      ['"Pat Fake" <pat@example.com>', "pat@example.com"],
    ])
      expect(extractEmailAddress(from), from).toBe(address);
  });

  it("throws the pinned message", () => {
    expect(() => extractEmailAddress("nobody")).toThrow(
      new Error("Could not get email address from nobody"),
    );
  });
});

describe("template and defaults hardening", () => {
  it("rejects an unclosed ${ placeholder", () => {
    const config = defaults();
    config.senders[0].rules[0].name = "${subject.name";
    expect(issueFor(config, "senders[0].rules[0].name")?.message).toMatch(
      /unclosed \$\{ placeholder/,
    );
  });

  it("does not let a parse result mutate DEFAULT_RECEIPT_RULES", () => {
    const parsed = parseReceiptRulesText("");
    if (!parsed.ok) throw new Error("expected ok");
    parsed.config.senders[0].address = "changed@example.com";
    parsed.config.forwarding?.subjectPrefixes?.push("X");
    expect(DEFAULT_RECEIPT_RULES.senders[0].address).toBe(
      "no.reply.alerts@chase.com",
    );
    expect(DEFAULT_RECEIPT_RULES.forwarding?.subjectPrefixes).toEqual([
      "Fwd: ",
    ]);
  });
});
