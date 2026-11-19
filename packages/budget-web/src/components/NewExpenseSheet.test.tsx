import { beforeEach, describe, expect, it, vi } from "vitest";
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
    expect(dialog.getByLabelText("Cost")).toHaveAttribute(
      "inputmode",
      "decimal",
    );
    for (const bad of ["4+", "1..2", "(4+1", "--4", "4a"]) {
      await user.clear(dialog.getByLabelText("Cost"));
      await user.type(dialog.getByLabelText("Cost"), bad);
      expect(dialog.getByLabelText("Cost")).toHaveAttribute(
        "aria-invalid",
        "true",
      );
      expect(dialog.getByRole("alert")).toHaveTextContent(
        "Enter a number or math, e.g. 4.23*2+3",
      );
      expect(save).toBeDisabled();
    }
    await user.clear(dialog.getByLabelText("Cost"));
    await user.type(dialog.getByLabelText("Cost"), "=4+1");
    expect(dialog.queryByRole("alert")).toBeNull();
    expect(dialog.getByText("$5.00")).toBeInTheDocument();
    await user.click(save);
    // Math is sent without its "=" (the server writes the formula =4+1)
    expect(onSave).toHaveBeenLastCalledWith(
      expect.objectContaining({ name: "Cash", costExpression: "4+1" }),
    );
  });

  it("saves a plain number exactly as typed", async () => {
    const { onSave, dialog, user } = setup();
    await user.type(dialog.getByLabelText("Name"), "Cash");
    await user.type(dialog.getByLabelText("Cost"), "-4.50");
    expect(dialog.queryByRole("alert")).toBeNull();
    await user.click(dialog.getByRole("button", { name: "Save" }));
    expect(onSave).toHaveBeenCalledWith(
      expect.objectContaining({ name: "Cash", costExpression: "-4.50" }),
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
    loadAllFailed: false,
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

describe("NewExpenseSheet choosers", () => {
  beforeEach(() => localStorage.clear());
  const envelopes = makeState([]).envelopes;
  const render_ = (
    props: Partial<Parameters<typeof NewExpenseSheet>[0]> = {},
  ) =>
    render(
      <NewExpenseSheet
        period={PERIOD}
        envelopeOptions={["Gas", "Groceries"]}
        envelopeLayoutOptions={["Groceries", "Gas"]}
        envelopes={envelopes}
        paidWithOptions={["Credit", "Pat Venmo"]}
        onSave={vi.fn()}
        onClose={vi.fn()}
        {...props}
      />,
    );
  const names = (group: HTMLElement) =>
    within(group)
      .getAllByRole("button", { pressed: false })
      .map((b) => b.textContent);

  it("shows the remaining amount under each envelope name, like the transaction sheet", () => {
    render_();
    const gas = screen.getByRole("button", { name: /^Gas/ });
    expect(gas).toHaveTextContent("-$30.00");
    expect(within(gas).getByText("-$30.00")).toHaveClass("text-2xs");
  });

  it("uses recent order collapsed and layout order expanded", async () => {
    const user = userEvent.setup();
    render_();
    const group = screen.getByRole("group", { name: "Envelope" });
    expect(names(group)[0]).toMatch(/^Gas/);
    await user.click(within(group).getByRole("button", { name: "Envelope" }));
    expect(names(group)[0]).toMatch(/^Groceries/);
  });

  it("starts expanded from the saved preferences and saves toggles", async () => {
    localStorage.setItem("budget.envelopeExpanded", "1");
    localStorage.setItem("budget.paidWithExpanded", "1");
    const user = userEvent.setup();
    render_();
    const env = screen.getByRole("group", { name: "Envelope" });
    const paid = screen.getByRole("group", { name: "Paid with" });
    expect(
      within(env).getByRole("button", { name: "Envelope" }),
    ).toHaveAttribute("aria-expanded", "true");
    expect(
      within(paid).getByRole("button", { name: "Paid with" }),
    ).toHaveAttribute("aria-expanded", "true");
    await user.click(within(paid).getByRole("button", { name: "Paid with" }));
    expect(localStorage.getItem("budget.paidWithExpanded")).toBe("0");
  });

  it("records a chosen envelope as recently used", async () => {
    const user = userEvent.setup();
    render_();
    await user.click(screen.getByRole("button", { name: /^Groceries/ }));
    expect(localStorage.getItem("budget.recentEnvelopes")).toBe(
      '["Groceries"]',
    );
  });
});

describe("ExpensesView inline envelope and Show in All", () => {
  beforeEach(() => localStorage.clear());
  const snapshot = {
    all: null,
    loadingAll: false,
    loadAllFailed: false,
    queue: [],
    conflicts: [],
    offline: false,
    recentlySaved: [],
  };
  function setup() {
    const state = makeState([
      makeRow(0, { name: "Aldi" }),
      makeRow(1, { name: "Shell", envelope: "Gas" }),
      makeRow(2, { name: "Aldi | milk", splitGroupId: "0" }),
    ]);
    const engine = { loadAll: vi.fn(), enqueueUpdate: vi.fn() };
    render(
      <ExpensesView
        engine={engine as unknown as SyncEngine}
        sheetName={SHEET}
        state={state}
        rows={buildViewRows(state, [])}
        showAll={false}
        snapshot={snapshot}
      />,
    );
    return { engine, state, user: userEvent.setup() };
  }

  it("sets the envelope of an uncategorized row from the list with the row's server copy", async () => {
    const { engine, state, user } = setup();
    expect(screen.getAllByLabelText(/^Envelope for/)).toHaveLength(2);
    await user.selectOptions(screen.getByLabelText("Envelope for Aldi"), "Gas");
    expect(engine.enqueueUpdate).toHaveBeenCalledWith(
      SHEET,
      state.transactions[0],
      { envelope: "Gas" },
    );
    // The row was not opened
    expect(screen.queryByRole("dialog")).toBeNull();
    expect(localStorage.getItem("budget.recentEnvelopes")).toBe('["Gas"]');
  });

  it("shows no dropdown on categorized rows", async () => {
    const { user } = setup();
    await user.click(screen.getByRole("tab", { name: /^All/ }));
    expect(screen.getAllByLabelText(/^Envelope for/)).toHaveLength(2);
    expect(screen.queryByLabelText("Envelope for Shell")).toBeNull();
  });

  it("jumps from a filter to the row in All; All has no jump buttons", async () => {
    const { user } = setup();
    const jumps = screen.getAllByRole("button", { name: "Show in All" });
    expect(jumps).toHaveLength(2);
    await user.click(jumps[0]);
    expect(screen.getByRole("tab", { name: /^All/ })).toHaveAttribute(
      "aria-selected",
      "true",
    );
    expect(screen.queryByRole("button", { name: "Show in All" })).toBeNull();
    const row = screen
      .getByText("Aldi")
      .closest<HTMLElement>("[data-row-key]")!;
    expect(row).toHaveClass("ring-teal-500");
  });

  it("does not indent split child rows", () => {
    setup();
    const child = screen
      .getByText("Aldi | milk")
      .closest<HTMLElement>("[data-row-key]")!;
    expect(child.innerHTML).not.toMatch(/pl-8/);
    expect(child.className).not.toMatch(/pl-8/);
  });
});
