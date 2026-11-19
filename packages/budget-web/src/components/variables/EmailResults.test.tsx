import { describe, expect, it } from "vitest";
import { render, screen, within } from "@testing-library/react";
import type { ReceiptEmailDto, RuleOutcome } from "@budget/shared";
import type {
  ImportPreview,
  PreviewItem,
} from "../../lib/receipt-rules-preview.ts";
import { EmailResults } from "./EmailResults.tsx";

const email: ReceiptEmailDto = {
  threadId: "t1",
  messageId: "m1",
  from: "Chase <no.reply.alerts@chase.com>",
  to: "me@example.com",
  subject: "You made a $4.50 transaction with Demo Cafe",
  date: "2026-01-02T03:04:05.000Z",
  plainBody: "Hello",
  plainBodyTruncated: false,
  labels: ["Receipts"],
};

const receipt: RuleOutcome = {
  senderAddress: "no.reply.alerts@chase.com",
  forwarded: false,
  ruleId: "chase-purchase",
  reason: null,
  receipt: {
    cost: 4.5,
    name: "Demo Cafe",
    category: "Food",
    type: "Credit",
  },
  isReceipt: true,
  captures: { subject: { cost: "4.50", name: "Demo Cafe" }, body: null },
  warnings: ["a warning"],
};

const notReceipt: RuleOutcome = {
  ...receipt,
  ruleId: null,
  reason: "no-rule-matched",
  receipt: {
    cost: undefined,
    name: undefined,
    category: undefined,
    type: "Credit",
  },
  isReceipt: false,
  captures: { subject: null, body: null },
  warnings: [],
};

const NO_IMPORT: ImportPreview = {
  rows: [],
  skip: null,
  alreadyProcessed: false,
};

const items: PreviewItem[] = [
  { email, outcome: receipt, error: null, import: NO_IMPORT },
  {
    email: { ...email, messageId: "m2", subject: "Your statement is ready" },
    outcome: notReceipt,
    error: null,
    import: NO_IMPORT,
  },
  {
    email: { ...email, messageId: "m3", subject: "Broken" },
    outcome: null,
    error: "Could not get email address from nobody",
    import: NO_IMPORT,
  },
  {
    email: { ...email, messageId: "m4", subject: "Re: Dinner Friday?" },
    outcome: {
      ...notReceipt,
      senderAddress: null,
      reason: "unknown-sender",
      receipt: { ...notReceipt.receipt, type: undefined },
    },
    error: null,
    import: NO_IMPORT,
  },
];

const card = (subject: string) =>
  screen.getByRole("article", { name: subject });

describe("EmailResults", () => {
  it("shows the empty state", () => {
    render(<EmailResults items={[]} />);
    expect(
      screen.getByText("Fetch emails to test the rules."),
    ).toBeInTheDocument();
  });

  it("shows a badge per outcome", () => {
    render(<EmailResults items={items} />);
    expect(
      within(card(email.subject)).getByTestId("outcome-badge"),
    ).toHaveTextContent("Receipt");
    expect(
      within(card("Your statement is ready")).getByTestId("outcome-badge"),
    ).toHaveTextContent("Not a receipt");
    const broken = within(card("Broken"));
    expect(broken.getByTestId("outcome-badge")).toHaveTextContent("Error");
    expect(broken.getByText(/Could not get email address/)).toBeInTheDocument();
    expect(
      within(card("Re: Dinner Friday?")).getByTestId("outcome-badge"),
    ).toHaveTextContent(/^Not a receipt: unknown sender$/);
  });

  it("shows captures, the receipt, rule, and warnings", () => {
    render(<EmailResults items={items} />);
    const c = within(card(email.subject));
    const captures = c.getByRole("list", { name: "Captures" });
    expect(captures.tagName).toBe("DL");
    expect(captures).toHaveTextContent("subject.name");
    expect(captures).toHaveTextContent("Demo Cafe");
    expect(c.getByText("chase-purchase")).toBeInTheDocument();
    expect(c.getByTestId("receipt-name")).toHaveTextContent("Demo Cafe");
    expect(c.getByTestId("receipt-cost")).toHaveTextContent("4.50");
    expect(c.getByText("Food")).toBeInTheDocument();
    expect(c.getByText("a warning")).toBeInTheDocument();
  });

  it("renders bodies as escaped text, never as HTML", () => {
    const evil = '<script>alert(1)</script><img src=x onerror="alert(2)">';
    render(
      <EmailResults
        items={[
          {
            email: { ...email, plainBody: evil, htmlBody: evil },
            outcome: receipt,
            error: null,
            import: NO_IMPORT,
          },
        ]}
      />,
    );
    expect(document.querySelector("script")).toBeNull();
    expect(document.querySelector("img")).toBeNull();
    const pres = document.querySelectorAll("pre");
    expect(pres).toHaveLength(2);
    for (const pre of pres) expect(pre.textContent).toBe(evil);
  });

  it("notes truncated bodies", () => {
    render(
      <EmailResults
        items={[
          {
            email: { ...email, plainBodyTruncated: true },
            outcome: receipt,
            error: null,
            import: NO_IMPORT,
          },
        ]}
      />,
    );
    expect(
      screen.getByText(/Body truncated by the server/),
    ).toBeInTheDocument();
    expect(
      screen.getByText(/Plain body \(truncated by the server\)/),
    ).toBeInTheDocument();
  });

  it("shows a visible warning outside the details when a body is truncated", () => {
    render(
      <EmailResults
        items={[
          {
            email: { ...email, htmlBody: "<p>x</p>", htmlBodyTruncated: true },
            outcome: receipt,
            error: null,
            import: NO_IMPORT,
          },
        ]}
      />,
    );
    const warning = screen.getByText(/Body truncated by the server/);
    expect(warning.closest("details")).toBeNull();
    expect(warning).toHaveTextContent(/may differ/);
  });

  describe("import preview", () => {
    const row = {
      date: new Date("2026-01-02T03:04:05.000Z"),
      name: "Demo Cafe",
      cost: 4.5,
      category: "Food",
      type: "Credit",
      errorMessage: "",
      note: "",
      sheet: "12/22/2025 - 1/4/2026" as string | null | undefined,
      skippedWithThread: false,
    };
    const withImport = (imp: Partial<ImportPreview>): PreviewItem => ({
      email,
      outcome: receipt,
      error: null,
      import: { ...NO_IMPORT, ...imp },
    });
    const section = () =>
      within(card(email.subject)).getByRole("region", {
        name: "Import preview",
      });

    it("shows each row's date, name, cost, envelope, paid with and sheet", () => {
      render(<EmailResults items={[withImport({ rows: [row] })]} />);
      const s = within(section());
      expect(s.getByTestId("import-name")).toHaveTextContent("Demo Cafe");
      expect(s.getByTestId("import-cost")).toHaveTextContent("4.50");
      expect(s.getByTestId("import-date")).toHaveTextContent(
        row.date.toLocaleDateString(),
      );
      expect(s.getByTestId("import-envelope")).toHaveTextContent("Food");
      expect(s.getByTestId("import-paid-with")).toHaveTextContent("Credit");
      expect(s.getByTestId("import-sheet")).toHaveTextContent(
        "12/22/2025 - 1/4/2026",
      );
    });

    it("shows an error row in red and a note row in pink, with the text", () => {
      const { rerender } = render(
        <EmailResults
          items={[withImport({ rows: [{ ...row, errorMessage: "it broke" }] })]}
        />,
      );
      expect(within(section()).getByTestId("import-name")).toHaveClass(
        "bg-red-600",
      );
      expect(section()).toHaveTextContent("it broke");
      rerender(
        <EmailResults
          items={[withImport({ rows: [{ ...row, note: "a note" }] })]}
        />,
      );
      expect(within(section()).getByTestId("import-name")).toHaveClass(
        "bg-pink-200",
      );
      expect(section()).toHaveTextContent("a note");
    });

    it("flags a missing transaction sheet and a missing cost", () => {
      render(
        <EmailResults
          items={[
            withImport({ rows: [{ ...row, sheet: null, cost: undefined }] }),
          ]}
        />,
      );
      expect(section()).toHaveTextContent(
        /No matching transaction sheet: the import skips this email thread/,
      );
      expect(within(section()).getByTestId("import-cost")).toHaveClass(
        "bg-pink-200",
      );
    });

    it("flags a receipt skipped because its thread has one with no sheet", () => {
      render(
        <EmailResults
          items={[withImport({ rows: [{ ...row, skippedWithThread: true }] })]}
        />,
      );
      expect(section()).toHaveTextContent(
        /Another receipt in this email thread has no matching transaction sheet/,
      );
    });

    it("says why nothing would be imported", () => {
      render(
        <EmailResults
          items={[withImport({ skip: "Not a receipt (no rule matched): x" })]}
        />,
      );
      expect(section()).toHaveTextContent(
        "Would not import: Not a receipt (no rule matched): x",
      );
    });

    it("notes an already processed thread", () => {
      render(<EmailResults items={[withImport({ alreadyProcessed: true })]} />);
      expect(section()).toHaveTextContent(/Already processed/);
    });
  });
});
