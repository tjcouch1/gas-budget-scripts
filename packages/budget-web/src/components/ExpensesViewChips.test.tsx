import { afterEach, beforeEach, describe, expect, it, vi } from "vitest";
import { render, screen, within } from "@testing-library/react";
import userEvent from "@testing-library/user-event";
import type { PayPeriodStateDto } from "@budget/shared";
import { ExpensesView } from "./ExpensesView.tsx";
import type { SyncEngine } from "../sync/engine.ts";
import { buildViewRows } from "../sync/optimistic.ts";
import { makeRow, makeState } from "../test/fixtures.ts";
import {
  setEnvelopeExpanded,
  setInlineEnvelopeChips,
} from "../lib/preferences.ts";

const snapshot = {
  all: null,
  loadingAll: false,
  loadAllFailed: false,
  queue: [],
  conflicts: [],
  offline: false,
  recentlySaved: [] as { sheetName: string; rowIndex: number }[],
};
const sheetName = "3/3/26 - 3/16/26";

function setup(state: PayPeriodStateDto) {
  const engine = {
    loadAll: vi.fn(),
    enqueueUpdate: vi.fn(),
    enqueueSplit: vi.fn(),
    enqueueDelete: vi.fn(),
  };
  const view = (s: PayPeriodStateDto, snap = snapshot) => (
    <ExpensesView
      engine={engine as unknown as SyncEngine}
      sheetName={sheetName}
      state={s}
      rows={buildViewRows(s, [])}
      showAll={false}
      snapshot={snap}
    />
  );
  const utils = render(view(state));
  return { engine, user: userEvent.setup(), view, ...utils };
}

const rowOf = (name: string) =>
  screen.getByText(name).closest<HTMLElement>("[data-row-key]")!;

describe("uncategorized rows with the chip selector setting", () => {
  beforeEach(() => localStorage.clear());
  afterEach(() => localStorage.clear());

  const state = () =>
    makeState([
      makeRow(0, { name: "Hardware" }),
      makeRow(1, { name: "Done", envelope: "Gas" }),
    ]);

  it("keeps the dropdown by default", () => {
    setup(state());
    expect(screen.getByLabelText("Envelope for Hardware")).toBeInTheDocument();
    expect(
      within(rowOf("Hardware")).queryByRole("group", { name: "Envelope" }),
    ).toBeNull();
  });

  it("shows a full-width scrolling chip selector below the row content, with no title or expand button", async () => {
    setInlineEnvelopeChips(true);
    const { user } = setup(state());
    await user.click(screen.getByRole("tab", { name: /^All/ }));
    const row = rowOf("Hardware");
    expect(screen.queryByLabelText("Envelope for Hardware")).toBeNull();
    const group = within(row).getByRole("group", { name: "Envelope" });
    expect(within(group).queryByText("Envelope")).toBeNull();
    expect(within(group).queryByRole("button", { expanded: false })).toBeNull();
    expect(within(group).getByTestId("chip-scroller")).toHaveClass(
      "overflow-x-auto",
    );
    // Below all the existing content: it is the row's last child, after the name/cost row
    expect(row.lastElementChild).toBe(group.parentElement);
    expect(row.lastElementChild).toHaveClass("w-full");
    expect(
      row.firstElementChild!.compareDocumentPosition(group) &
        Node.DOCUMENT_POSITION_FOLLOWING,
    ).toBeTruthy();
    // Categorized rows get none
    expect(
      within(rowOf("Done")).queryByRole("group", { name: "Envelope" }),
    ).toBeNull();
  });

  it("stays collapsed even when the saved expand preference is on", () => {
    setEnvelopeExpanded(true);
    setInlineEnvelopeChips(true);
    setup(state());
    expect(screen.getByTestId("chip-scroller")).toHaveClass("flex-nowrap");
  });

  it("saves via enqueueUpdate without opening the row, and keeps showing the selection after the row is categorized", async () => {
    setInlineEnvelopeChips(true);
    const s = state();
    const { engine, user, view, rerender } = setup(s);
    await user.click(
      within(rowOf("Hardware")).getByRole("button", { name: /^Gas/ }),
    );
    expect(screen.queryByRole("dialog")).toBeNull();
    expect(engine.enqueueUpdate).toHaveBeenCalledWith(
      sheetName,
      s.transactions[0],
      { envelope: "Gas" },
    );
    // The save is in flight: the row is categorized but still listed
    const saved = makeState([
      makeRow(0, { name: "Hardware", envelope: "Gas" }),
      makeRow(1, { name: "Done", envelope: "Gas" }),
    ]);
    rerender(
      view(saved, { ...snapshot, recentlySaved: [{ sheetName, rowIndex: 0 }] }),
    );
    const group = within(rowOf("Hardware")).getByRole("group", {
      name: "Envelope",
    });
    expect(within(group).getByRole("button", { name: /^Gas/ })).toHaveAttribute(
      "aria-pressed",
      "true",
    );
  });
});
