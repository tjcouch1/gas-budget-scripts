import { describe, expect, it } from "vitest";
import {
  compileReceiptRules,
  DEFAULT_RECEIPT_RULES,
  planThreadImport,
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
      formatOutcome({
        email: receipt.email,
        outcome: null,
        error: "x",
        import: receipt.import,
      }),
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

describe("import preview (what the import would write)", () => {
  const sheets = [
    {
      name: "12/22/2025 - 1/4/2026",
      startDate: "2025-12-22T06:00:00.000Z",
      endDate: "2026-01-05T05:59:59.999Z",
    },
  ];

  it("shows the transaction row the import would write, with its sheet", () => {
    const [item] = previewReceiptRules(
      rules,
      [dto(demoInbox.chasePurchase!)],
      context,
      sheets,
    );
    expect(item!.import.skip).toBeNull();
    expect(item!.import.rows).toEqual([
      {
        date: new Date("2026-01-02T03:04:05.000Z"),
        name: "Demo Cafe",
        cost: 18.42,
        category: undefined,
        type: item!.outcome!.receipt.type,
        errorMessage: "",
        note: "",
        sheet: "12/22/2025 - 1/4/2026",
      },
    ]);
  });

  it("is built by the shared planner the GAS import uses", () => {
    const email = dto(demoInbox.chasePurchase!);
    const [item] = previewReceiptRules(rules, [email], context, sheets);
    const plan = planThreadImport({
      threadId: email.threadId,
      lastMessageDate: () => new Date(email.date),
      messages: [
        {
          messageId: email.messageId,
          subject: email.subject,
          fromAddress: email.from,
          date: new Date(email.date),
          dateText: String(new Date(email.date)),
          plainBody: email.plainBody,
          outcome: item!.outcome,
          error: null,
        },
      ],
    });
    expect(item!.import.rows).toMatchObject(
      plan.rows.map(({ messageId: _id, ...row }) => row),
    );
  });

  it("explains why a non-receipt would not be imported", () => {
    const [item] = previewReceiptRules(
      rules,
      [dto(demoInbox.chaseStatement!)],
      context,
      sheets,
    );
    expect(item!.import.rows).toEqual([]);
    expect(item!.import.skip).toBe(
      "Not a receipt (no rule matched): the thread is only labeled Receipts/Ignored",
    );
  });

  it("flags a receipt dated outside every transaction sheet", () => {
    const [item] = previewReceiptRules(
      rules,
      [dto(demoInbox.chasePurchase!, { date: "2030-01-01T00:00:00.000Z" })],
      context,
      sheets,
    );
    expect(item!.import.rows[0]!.sheet).toBeNull();
  });

  it("leaves the sheet undefined when sheets are unknown", () => {
    const [item] = previewReceiptRules(
      rules,
      [dto(demoInbox.chasePurchase!)],
      context,
    );
    expect(item!.import.rows[0]!.sheet).toBeUndefined();
  });

  it("notes emails already labeled Receipts/Scripted", () => {
    const [done, fresh] = previewReceiptRules(
      rules,
      [
        dto(demoInbox.chasePurchase!, { labels: ["Receipts/Scripted"] }),
        dto(demoInbox.chasePurchase!),
      ],
      context,
      sheets,
    );
    expect(done!.import.alreadyProcessed).toBe(true);
    expect(fresh!.import.alreadyProcessed).toBe(false);
  });

  it("shows an errored email as the red empty row the import writes", () => {
    const bad = { config: DEFAULT_RECEIPT_RULES } as never;
    const [item] = previewReceiptRules(
      bad,
      [dto(demoInbox.chasePurchase!)],
      context,
      sheets,
    );
    expect(item!.import.skip).toBeNull();
    expect(item!.import.rows).toHaveLength(1);
    expect(item!.import.rows[0]).toMatchObject({
      name: undefined,
      cost: undefined,
      sheet: "12/22/2025 - 1/4/2026",
    });
    expect(item!.import.rows[0]!.errorMessage).toContain(
      "Error while processing message m1",
    );
  });

  it("puts a thread's note on its first receipt and none on the non-receipt message", () => {
    const [first, second] = previewReceiptRules(
      rules,
      [
        dto(demoInbox.chasePurchase!, { messageId: "a" }),
        dto(demoInbox.chaseStatement!, { messageId: "b" }),
      ],
      context,
      sheets,
    );
    expect(first!.import.rows[0]!.note).toContain("Message is not a receipt");
    expect(second!.import.rows).toEqual([]);
    expect(second!.import.skip).toMatch(/^Not a receipt/);
  });

  it("with invalid rules, has nothing to import", () => {
    const [item] = previewReceiptRules(
      null,
      [dto(demoInbox.chasePurchase!)],
      context,
      sheets,
    );
    expect(item!.import.rows).toEqual([]);
    expect(item!.import.skip).toBe("Rules are invalid");
  });
});
