import { describe, expect, it, vi } from "vitest";
import { render, screen, within } from "@testing-library/react";
import userEvent from "@testing-library/user-event";
import type { PayPeriodStateDto } from "@budget/shared";
import { NewExpenseSheet, todayLocalIso } from "./NewExpenseSheet.tsx";
import { ExpensesView } from "./ExpensesView.tsx";
import type { SyncEngine } from "../sync/engine.ts";
import { buildViewRows } from "../sync/optimistic.ts";
import { makeRow, makeState } from "../test/fixtures.ts";

const SHEET = "3/3/26 - 3/16/26";
const PERIOD = { startDate: "2026-03-03", endDate: "2026-03-16" };

describe("todayLocalIso", () => {
  it("formats the browser's local date as yyyy-MM-dd", () => {
    expect(todayLocalIso(new Date(2026, 2, 5, 23, 59))).toBe("2026-03-05");
    expect(todayLocalIso(new Date(2026, 11, 31, 0, 1))).toBe("2026-12-31");
  });
});

describe("NewExpenseSheet", () => {
  function setup(props: Partial<Parameters<typeof NewExpenseSheet>[0]> = {}) {
    const onSave = vi.fn();
    const onClose = vi.fn();
    render(
      <NewExpenseSheet
        period={PERIOD}
        envelopeOptions={["Groceries", "Gas"]}
        paidWithOptions={["Credit", "Pat Venmo"]}
        today="2026-03-07"
        onSave={onSave}
        onClose={onClose}
        {...props}
      />,
    );
    const dialog = within(screen.getByRole("dialog", { name: "New expense" }));
    return { onSave, onClose, dialog, user: userEvent.setup() };
  }

  it("defaults the date to today and saves every field", async () => {
    const { onSave, onClose, dialog, user } = setup();
    expect(dialog.getByLabelText("Date")).toHaveValue("2026-03-07");
    const save = dialog.getByRole("button", { name: "Save" });
    expect(save).toBeDisabled();
    await user.type(dialog.getByLabelText("Name"), "Farmers market");
    await user.type(dialog.getByLabelText("Cost"), "4.25*2");
    expect(dialog.getByText("$8.50")).toBeInTheDocument();
    await user.click(
      within(dialog.getByRole("group", { name: "Envelope" })).getByRole(
        "button",
        { name: "Groceries" },
      ),
    );
    await user.click(
      within(dialog.getByRole("group", { name: "Paid with" })).getByRole(
        "button",
        { name: "Pat Venmo" },
      ),
    );
    await user.click(dialog.getByLabelText("Tax deductible"));
    await user.type(dialog.getByLabelText("Note (optional)"), "cash");
    await user.click(save);
    expect(onSave).toHaveBeenCalledWith({
      date: "2026-03-07",
      name: "Farmers market",
      costExpression: "4.25*2",
      envelope: "Groceries",
      paidWith: "Pat Venmo",
      taxDeductible: "Yes",
      note: "cash",
    });
    expect(onClose).toHaveBeenCalled();
  });

  it("requires a safe name and a valid cost", async () => {
    const { onSave, dialog, user } = setup();
    const save = dialog.getByRole("button", { name: "Save" });
    await user.type(dialog.getByLabelText("Name"), "=HYPERLINK()");
    await user.type(dialog.getByLabelText("Cost"), "12");
    expect(dialog.getByText(/must not start with/)).toBeInTheDocument();
    expect(save).toBeDisabled();
    await user.clear(dialog.getByLabelText("Name"));
    await user.type(dialog.getByLabelText("Name"), "Cash");
    expect(save).toBeEnabled();
    await user.clear(dialog.getByLabelText("Cost"));
    await user.type(dialog.getByLabelText("Cost"), "4+");
    expect(dialog.getByLabelText("Cost")).toHaveAttribute(
      "aria-invalid",
      "true",
    );
    expect(save).toBeDisabled();
    await user.clear(dialog.getByLabelText("Cost"));
    await user.type(dialog.getByLabelText("Cost"), "=4+1");
    await user.click(save);
    expect(onSave).toHaveBeenCalledWith(
      expect.objectContaining({ name: "Cash", costExpression: "4+1" }),
    );
  });

  it("warns about a date outside the pay period and lets the user choose", async () => {
    const { onSave, dialog, user } = setup({ today: "2026-03-20" });
    expect(dialog.getByRole("alert")).toHaveTextContent(
      /outside this pay period/,
    );
    await user.type(dialog.getByLabelText("Name"), "Late");
    await user.type(dialog.getByLabelText("Cost"), "3");
    await user.click(dialog.getByRole("button", { name: "Save" }));
    // The first Save asks; nothing is saved yet
    expect(onSave).not.toHaveBeenCalled();
    await user.click(dialog.getByRole("button", { name: "Change date" }));
    expect(dialog.getByLabelText("Date")).toHaveFocus();
    await user.click(dialog.getByRole("button", { name: "Save" }));
    await user.click(dialog.getByRole("button", { name: "Save anyway" }));
    expect(onSave).toHaveBeenCalledWith(
      expect.objectContaining({ date: "2026-03-20", name: "Late" }),
    );
  });

  it("shows no warning for a date inside the period, including its last day", async () => {
    const { dialog, user } = setup({ today: "2026-03-16" });
    expect(dialog.queryByRole("alert")).toBeNull();
    const date = dialog.getByLabelText("Date");
    await user.clear(date);
    await user.type(date, "2026-03-02");
    expect(dialog.getByRole("alert")).toBeInTheDocument();
  });

  it("closes at once when nothing changed, and guards unsaved changes", async () => {
    const { onClose, onSave, dialog, user } = setup();
    await user.type(dialog.getByLabelText("Name"), "Cash");
    await user.click(dialog.getByRole("button", { name: "Close" }));
    expect(onClose).not.toHaveBeenCalled();
    const prompt = within(
      screen.getByRole("alertdialog", { name: "Unsaved changes" }),
    );
    // Cost is still empty, so the prompt cannot save
    expect(prompt.getByRole("button", { name: "Save" })).toBeDisabled();
    await user.click(prompt.getByRole("button", { name: "Discard" }));
    expect(onClose).toHaveBeenCalled();
    expect(onSave).not.toHaveBeenCalled();
  });

  it("closes without a prompt when untouched", async () => {
    const { onClose, dialog, user } = setup();
    await user.click(dialog.getByRole("button", { name: "Close" }));
    expect(onClose).toHaveBeenCalled();
    expect(screen.queryByRole("alertdialog")).toBeNull();
  });

  it("hides the choosers when the sheet has no option lists", () => {
    const { dialog } = setup({ envelopeOptions: null, paidWithOptions: null });
    expect(dialog.queryByRole("group", { name: "Envelope" })).toBeNull();
    expect(dialog.queryByRole("group", { name: "Paid with" })).toBeNull();
  });
});

describe("ExpensesView New expense button", () => {
  const snapshot = {
    all: null,
    loadingAll: false,
    queue: [],
    conflicts: [],
    offline: false,
    recentlySaved: [],
  };
  function setup(state: PayPeriodStateDto | undefined, showAll = false) {
    const engine = { loadAll: vi.fn(), enqueueAdd: vi.fn() };
    render(
      <ExpensesView
        engine={engine as unknown as SyncEngine}
        sheetName={SHEET}
        state={state}
        rows={buildViewRows(state, [])}
        showAll={showAll}
        snapshot={snapshot}
      />,
    );
    return { engine, user: userEvent.setup() };
  }

  it("sits after the list and queues the new expense on the selected pay period", async () => {
    const { engine, user } = setup(makeState([makeRow(0, { name: "Aldi" })]));
    const button = screen.getByRole("button", { name: "New expense" });
    const list = screen.getByRole("list");
    // After the list in document order (the user scrolls down to it)
    expect(
      list.compareDocumentPosition(button) & Node.DOCUMENT_POSITION_FOLLOWING,
    ).toBeTruthy();
    await user.click(button);
    const dialog = within(screen.getByRole("dialog", { name: "New expense" }));
    await user.clear(dialog.getByLabelText("Date"));
    await user.type(dialog.getByLabelText("Date"), "2026-03-05");
    await user.type(dialog.getByLabelText("Name"), "Cash");
    await user.type(dialog.getByLabelText("Cost"), "5");
    await user.click(dialog.getByRole("button", { name: "Save" }));
    expect(engine.enqueueAdd).toHaveBeenCalledWith(SHEET, {
      date: "2026-03-05",
      name: "Cash",
      costExpression: "5",
      envelope: "",
      paidWith: "",
      taxDeductible: "",
      note: "",
    });
    expect(screen.queryByRole("dialog")).toBeNull();
  });

  it("is shown for an empty pay period", () => {
    setup(makeState([]));
    expect(
      screen.getByRole("button", { name: "New expense" }),
    ).toBeInTheDocument();
  });

  it("is hidden in the All view", () => {
    setup(makeState([makeRow(0)]), true);
    expect(screen.queryByRole("button", { name: "New expense" })).toBeNull();
  });

  it("is hidden while the pay period loads", () => {
    setup(undefined);
    expect(screen.queryByRole("button", { name: "New expense" })).toBeNull();
  });
});
