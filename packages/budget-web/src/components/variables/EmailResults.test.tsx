import { describe, expect, it } from "vitest";
import { render, screen, within } from "@testing-library/react";
import type { ReceiptEmailDto, RuleOutcome } from "@budget/shared";
import type { PreviewItem } from "../../lib/receipt-rules-preview.ts";
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

const items: PreviewItem[] = [
  { email, outcome: receipt, error: null },
  {
    email: { ...email, messageId: "m2", subject: "Your statement is ready" },
    outcome: notReceipt,
    error: null,
  },
  {
    email: { ...email, messageId: "m3", subject: "Broken" },
    outcome: null,
    error: "No function to handle receipt message from x@example.com",
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
    expect(broken.getByText(/No function to handle/)).toBeInTheDocument();
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
          },
        ]}
      />,
    );
    const warning = screen.getByText(/Body truncated by the server/);
    expect(warning.closest("details")).toBeNull();
    expect(warning).toHaveTextContent(/may differ/);
  });
});
