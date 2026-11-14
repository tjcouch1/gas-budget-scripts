import { beforeEach, describe, expect, it, vi } from "vitest";
import { render, screen, within } from "@testing-library/react";
import userEvent from "@testing-library/user-event";
import { TransactionSheet } from "./TransactionSheet.tsx";
import { makeRow, makeState } from "../test/fixtures.ts";
import type { ViewRow } from "../lib/rows.ts";
import { ENVELOPE_EXPANDED_KEY } from "../lib/preferences.ts";

const row = (i: number, over: Partial<ViewRow> = {}): ViewRow => ({
  ...makeRow(i),
  key: `row:${i}`,
  ...over,
});
const group = () => [
  row(4, { name: "Aldi | cart", cost: 18, splitGroupId: "4" }),
  row(5, { name: "Aldi | milk", cost: 10, splitGroupId: "4" }),
  row(6, { name: "Aldi | soap", cost: 2, splitGroupId: "4" }),
];

function setup(
  rows: ViewRow[] = [row(0, { name: "Aldi" })],
  props: Partial<Parameters<typeof TransactionSheet>[0]> = {},
) {
  const state = makeState([]);
  render(
    <TransactionSheet
      rows={rows}
      taxMultiplier={1.0825}
      envelopeOptions={["Groceries", "Gas"]}
      envelopes={state.envelopes}
      paidWithOptions={["Credit", "Pat Venmo"]}
      onUpdate={vi.fn()}
      onSplit={vi.fn()}
      onDelete={vi.fn()}
      onClose={vi.fn()}
      {...props}
    />,
  );
  return userEvent.setup();
}
const envelope = () => screen.getByRole("button", { name: "Envelope" });
const toggleOf = (card: HTMLElement) =>
  within(card).getByRole("button", { name: "Envelope" });

describe("TransactionSheet layout and preferences", () => {
  beforeEach(() => localStorage.clear());

  it("hides the chip scrollbars", () => {
    setup();
    for (const s of screen.getAllByTestId("chip-scroller"))
      expect(s).toHaveClass("no-scrollbar");
  });

  it("uses smaller chips", () => {
    setup();
    const chip = screen.getByRole("button", { name: /Gas/ });
    expect(chip).toHaveClass("min-h-9");
    expect(chip).not.toHaveClass("min-h-11");
  });

  it("keeps Paid with inside Edit details", () => {
    setup();
    const details = screen.getByText("Edit details").closest("details")!;
    expect(
      within(details).getByRole("group", { name: "Paid with" }),
    ).toBeInTheDocument();
  });

  it("puts Add split item and Delete on one justify-between row", () => {
    setup();
    const add = screen.getByRole("button", { name: "Add split item" });
    const del = screen.getByRole("button", { name: "Delete transaction" });
    expect(add.parentElement).toBe(del.parentElement);
    expect(add.parentElement).toHaveClass("justify-between");
  });

  it("renders an unsplit transaction without a card border", () => {
    setup();
    const card = screen.getByRole("group", { name: "Transaction" });
    expect(card).not.toHaveClass("border");
    expect(card.querySelector("legend")).toBeNull();
  });

  it("labels split cards with a legend on the border", async () => {
    const user = setup(group());
    const original = screen.getByRole("group", { name: "Original" });
    expect(original).toHaveClass("border");
    expect(original.tagName).toBe("FIELDSET");
    expect(original.querySelector("legend")).toHaveTextContent("Original");
    await user.click(screen.getByRole("button", { name: "Add split item" }));
    const item = screen.getByRole("group", { name: "Split item 3" });
    expect(item.querySelector("legend")).toHaveTextContent("Split item 3");
    expect(
      within(item).getByRole("button", { name: "Remove split item 3" }),
    ).toBeInTheDocument();
  });

  it("starts the Envelope chooser collapsed by default and saves toggles", async () => {
    const user = setup();
    const toggle = envelope();
    expect(toggle).toHaveAttribute("aria-expanded", "false");
    await user.click(toggle);
    expect(localStorage.getItem(ENVELOPE_EXPANDED_KEY)).toBe("1");
    await user.click(toggle);
    expect(localStorage.getItem(ENVELOPE_EXPANDED_KEY)).toBe("0");
  });

  it("starts an unsplit Envelope chooser in the saved state", () => {
    localStorage.setItem(ENVELOPE_EXPANDED_KEY, "1");
    setup();
    expect(envelope()).toHaveAttribute("aria-expanded", "true");
  });

  it("in a group only the clicked row's card follows the setting", async () => {
    localStorage.setItem(ENVELOPE_EXPANDED_KEY, "1");
    const user = setup(group(), { focusRowIndex: 5 });
    const open = (name: string) =>
      toggleOf(screen.getByRole("group", { name })).getAttribute(
        "aria-expanded",
      );
    expect(open("Original")).toBe("false");
    expect(open("Split item 1")).toBe("true");
    expect(open("Split item 2")).toBe("false");
    await user.click(toggleOf(screen.getByRole("group", { name: "Original" })));
    expect(localStorage.getItem(ENVELOPE_EXPANDED_KEY)).toBe("1");
    await user.click(
      toggleOf(screen.getByRole("group", { name: "Split item 1" })),
    );
    expect(localStorage.getItem(ENVELOPE_EXPANDED_KEY)).toBe("0");
  });

  it("starts Paid with collapsed even when the preference is expanded", () => {
    localStorage.setItem(ENVELOPE_EXPANDED_KEY, "1");
    setup();
    expect(screen.getByRole("button", { name: "Paid with" })).toHaveAttribute(
      "aria-expanded",
      "false",
    );
  });
});

describe("Edit details and Paid with preferences", () => {
  beforeEach(() => localStorage.clear());
  const detailsOf = (name: string) =>
    within(screen.getByRole("group", { name }))
      .getByText("Edit details")
      .closest("details")!;
  const paidOf = (name: string) =>
    within(screen.getByRole("group", { name })).getByRole("button", {
      name: "Paid with",
    });

  it("saves toggles of Edit details and Paid with on an unsplit row", async () => {
    const user = setup();
    expect(detailsOf("Transaction")).not.toHaveAttribute("open");
    await user.click(screen.getByText("Edit details"));
    await vi.waitFor(() =>
      expect(localStorage.getItem("budget.detailsOpen")).toBe("1"),
    );
    await user.click(paidOf("Transaction"));
    expect(localStorage.getItem("budget.paidWithExpanded")).toBe("1");
    await user.click(paidOf("Transaction"));
    expect(localStorage.getItem("budget.paidWithExpanded")).toBe("0");
  });

  it("opens an unsplit row in the saved states", () => {
    localStorage.setItem("budget.detailsOpen", "1");
    localStorage.setItem("budget.paidWithExpanded", "1");
    setup();
    expect(detailsOf("Transaction")).toHaveAttribute("open");
    expect(paidOf("Transaction")).toHaveAttribute("aria-expanded", "true");
  });

  it("in a group only the clicked card follows; new items never do", async () => {
    localStorage.setItem("budget.detailsOpen", "1");
    localStorage.setItem("budget.paidWithExpanded", "1");
    const user = setup(group(), { focusRowIndex: 5 });
    expect(detailsOf("Split item 1")).toHaveAttribute("open");
    expect(paidOf("Split item 1")).toHaveAttribute("aria-expanded", "true");
    expect(detailsOf("Original")).not.toHaveAttribute("open");
    expect(paidOf("Original")).toHaveAttribute("aria-expanded", "false");
    await user.click(screen.getByRole("button", { name: "Add split item" }));
    expect(paidOf("Split item 3")).toHaveAttribute("aria-expanded", "false");
    await user.click(paidOf("Split item 3"));
    await user.click(paidOf("Original"));
    expect(localStorage.getItem("budget.paidWithExpanded")).toBe("1");
    await user.click(paidOf("Split item 1"));
    expect(localStorage.getItem("budget.paidWithExpanded")).toBe("0");
  });
});
