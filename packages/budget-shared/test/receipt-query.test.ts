import { describe, expect, it } from "vitest";
import {
  buildReceiptSearchQuery,
  isSafeSubjectFilter,
  isStrictEmailAddress,
  RECEIPT_LABEL_EXCLUSIONS,
} from "../src/receipt-query.ts";

const DEFAULT_SENDERS = [
  "no.reply.alerts@chase.com",
  "service@paypal.com",
  "venmo@venmo.com",
];

describe("buildReceiptSearchQuery", () => {
  it("matches the pinned query for the default senders", () => {
    expect(buildReceiptSearchQuery(DEFAULT_SENDERS, "fwd@example.com")).toBe(
      "from:(no.reply.alerts@chase.com OR service@paypal.com OR venmo@venmo.com OR fwd@example.com) in:inbox NOT label:receipts NOT label:receipts-cru-reimburse NOT label:receipts-tax-deductible NOT label:receipts-scripted ",
    );
  });

  it("exposes the fixed label exclusions", () => {
    expect(RECEIPT_LABEL_EXCLUSIONS).toBe(
      "NOT label:receipts NOT label:receipts-cru-reimburse NOT label:receipts-tax-deductible NOT label:receipts-scripted",
    );
  });

  it("omits the forward address when empty", () => {
    expect(buildReceiptSearchQuery(["a@b.co"], "")).toBe(
      `from:(a@b.co) in:inbox ${RECEIPT_LABEL_EXCLUSIONS} `,
    );
  });

  it("throws for an invalid sender or forward address", () => {
    expect(() => buildReceiptSearchQuery(["x"], "")).toThrow(
      "Invalid sender address: x",
    );
    expect(() => buildReceiptSearchQuery(["a@b.co"], "bad addr")).toThrow(
      "Invalid sender address: bad addr",
    );
  });

  it("emits quoted subjects and newer_than between from and in:inbox", () => {
    expect(
      buildReceiptSearchQuery(["a@b.co"], "f@b.co", {
        subjects: ["Receipt", "Your order"],
        newerThanDays: 30,
      }),
    ).toBe(
      `from:(a@b.co OR f@b.co) subject:("Receipt" OR "Your order") newer_than:30d in:inbox ${RECEIPT_LABEL_EXCLUSIONS} `,
    );
  });

  it("includeProcessed omits in:inbox and exclusions but keeps the trailing space", () => {
    expect(
      buildReceiptSearchQuery(["a@b.co"], "", { includeProcessed: true }),
    ).toBe("from:(a@b.co) ");
    expect(
      buildReceiptSearchQuery(["a@b.co"], "", {
        includeProcessed: true,
        newerThanDays: 7,
      }),
    ).toBe("from:(a@b.co) newer_than:7d ");
  });

  it("rejects unsafe subjects and out-of-range days", () => {
    expect(() =>
      buildReceiptSearchQuery(["a@b.co"], "", { subjects: ['a"b'] }),
    ).toThrow("Invalid subject filter");
    for (const n of [0, 3651, 1.5, NaN]) {
      expect(() =>
        buildReceiptSearchQuery(["a@b.co"], "", { newerThanDays: n }),
      ).toThrow("Invalid newerThanDays");
    }
  });
});

describe("isSafeSubjectFilter", () => {
  it("accepts plain text", () => {
    expect(isSafeSubjectFilter("Your receipt #1")).toBe(true);
    expect(isSafeSubjectFilter("x".repeat(100))).toBe(true);
  });
  it("rejects quotes, parens, braces, colon, backslash, newline, blank, long, non-string", () => {
    for (const bad of [
      'a"b',
      "a(b",
      "a)b",
      "a{b",
      "a}b",
      "a:b",
      "a\\b",
      "a\nb",
      "   ",
      "",
      "x".repeat(101),
      5,
      null,
    ])
      expect(isSafeSubjectFilter(bad), String(bad)).toBe(false);
  });
});

describe("isStrictEmailAddress", () => {
  it("accepts normal addresses", () => {
    expect(isStrictEmailAddress("no.reply.alerts@chase.com")).toBe(true);
    expect(isStrictEmailAddress("a+b@sub.example.org")).toBe(true);
  });
  it("rejects spaces, parens, OR, missing parts, long, non-string", () => {
    for (const bad of [
      "a b@c.com",
      "a@c.com OR b@c.com",
      "(a@c.com)",
      "a@c",
      "@c.com",
      "OR",
      `${"a".repeat(250)}@c.com`,
      `${"a".repeat(65)}@c.com`,
      undefined,
      12,
    ])
      expect(isStrictEmailAddress(bad), String(bad)).toBe(false);
  });
});
