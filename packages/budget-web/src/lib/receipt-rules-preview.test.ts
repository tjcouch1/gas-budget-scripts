import { describe, expect, it } from "vitest";
import {
  compileReceiptRules,
  DEFAULT_RECEIPT_RULES,
  type ReceiptEmailDto,
} from "@budget/shared";
import {
  demoInbox,
  receiptEdgeFixtures,
} from "../../../budget-scripts-gas/test/fixtures/receipts.ts";
import {
  formatOutcome,
  previewReceiptRules,
  toRuleEmail,
  type PreviewItem,
} from "./receipt-rules-preview.ts";

const context = { forwardEmailAddress: "fwd@example.com", forwardName: "Sam" };
const rules = compileReceiptRules(DEFAULT_RECEIPT_RULES, {
  forwardEmailAddress: context.forwardEmailAddress,
});

function dto(
  fixture: (typeof demoInbox)[string],
  extra: Partial<ReceiptEmailDto> = {},
): ReceiptEmailDto {
  return {
    threadId: "t1",
    messageId: "m1",
    from: fixture.from,
    to: fixture.to ?? "me@example.com",
    subject: fixture.subject,
    date: "2026-01-02T03:04:05.000Z",
    plainBody: fixture.body,
    plainBodyTruncated: false,
    labels: [],
    ...extra,
  };
}

describe("toRuleEmail", () => {
  it("maps the DTO fields and only includes htmlBody when present", () => {
    const email = dto(demoInbox.chasePurchase!);
    expect(toRuleEmail(email)).toEqual({
      from: email.from,
      to: email.to,
      subject: email.subject,
      plainBody: email.plainBody,
    });
    expect(toRuleEmail({ ...email, htmlBody: "<p>x</p>" }).htmlBody).toBe(
      "<p>x</p>",
    );
  });
});

describe("previewReceiptRules", () => {
  it("applies the rules to each email", () => {
    const items = previewReceiptRules(
      rules,
      [dto(demoInbox.chasePurchase!), dto(demoInbox.chaseStatement!)],
      context,
    );
    expect(items).toHaveLength(2);
    expect(items[0]!.error).toBeNull();
    expect(items[0]!.outcome?.isReceipt).toBe(true);
    expect(items[0]!.outcome?.receipt.cost).toBe(18.42);
    expect(items[1]!.outcome?.isReceipt).toBe(false);
  });

  it("turns a thrown engine error into the item error", () => {
    const bad = { config: DEFAULT_RECEIPT_RULES } as never;
    const [item] = previewReceiptRules(
      bad,
      [dto(demoInbox.chasePurchase!)],
      context,
    );
    expect(item!.outcome).toBeNull();
    expect(item!.error).toMatch(/compileReceiptRules/);
  });

  it("reports invalid rules when none are compiled", () => {
    const [item] = previewReceiptRules(
      null,
      [dto(demoInbox.chasePurchase!)],
      context,
    );
    expect(item).toMatchObject({ outcome: null, error: "Rules are invalid" });
  });
});

describe("formatOutcome", () => {
  const [receipt, other] = previewReceiptRules(
    rules,
    [dto(demoInbox.chasePurchase!), dto(demoInbox.chaseStatement!)],
    context,
  ) as [PreviewItem, PreviewItem];

  it("labels receipts, non-receipts and errors", () => {
    expect(formatOutcome(receipt)).toEqual({ label: "Receipt", tone: "ok" });
    expect(formatOutcome(other)).toEqual({
      label: "Not a receipt",
      tone: "muted",
      detail: "no rule matched",
    });
    expect(
      formatOutcome({ email: receipt.email, outcome: null, error: "x" }),
    ).toEqual({ label: "Error", tone: "error" });
  });

  it("explains why an email is not a receipt", () => {
    const [unknown, reply] = previewReceiptRules(
      rules,
      [
        dto(receiptEdgeFixtures.edgeUnknownSender.email),
        dto(receiptEdgeFixtures.edgeForwarderReply.email),
      ],
      context,
    ) as [PreviewItem, PreviewItem];
    expect(formatOutcome(unknown)).toEqual({
      label: "Not a receipt",
      tone: "muted",
      detail: "unknown sender",
    });
    expect(formatOutcome(reply).detail).toBe(
      "from the forward address but not a forward",
    );
  });
});
