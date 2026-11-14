import { describe, expect, it, vi } from "vitest";
import {
  fireEvent,
  render,
  screen,
  within,
  waitFor,
} from "@testing-library/react";
import userEvent from "@testing-library/user-event";
import { ConflictDialog } from "./ConflictDialog.tsx";
import { EnvelopesView } from "./EnvelopesView.tsx";
import { TransactionSheet } from "./TransactionSheet.tsx";
import { StatusPill, statusOf } from "./StatusPill.tsx";
import { makeRow, makeState } from "../test/fixtures.ts";
import type { EngineSnapshot, Conflict } from "../sync/types.ts";
import type { PayPeriodStateDto, TransactionRowDto } from "@budget/shared";
import { ExpensesView } from "./ExpensesView.tsx";
import type { SyncEngine } from "../sync/engine.ts";
import { buildViewRows } from "../sync/optimistic.ts";
import type { ViewRow } from "../lib/rows.ts";

describe("TransactionSheet", () => {
  const row = (i: number, over: Partial<TransactionRowDto> = {}): ViewRow => ({
    ...makeRow(i, over),
    key: `row:${i}`,
  });
  function setup(
    rowOver: Partial<TransactionRowDto> = {},
    props: Partial<Parameters<typeof TransactionSheet>[0]> = {},
  ) {
    const onUpdate = vi.fn();
    const onSplit = vi.fn();
    const onClose = vi.fn();
    const state = makeState([]);
    render(
      <TransactionSheet
        rows={[row(0, { name: "Aldi", ...rowOver })]}
        taxMultiplier={1.0825}
        envelopeOptions={["Groceries", "Gas"]}
        envelopes={state.envelopes}
        paidWithOptions={["Credit", "Pat Venmo"]}
        onUpdate={onUpdate}
        onSplit={onSplit}
        onClose={onClose}
        {...props}
      />,
    );
    return { onUpdate, onSplit, onClose, user: userEvent.setup() };
  }
  /** A group as the sheet has it after a split: Aldi, $30 in total. */
  const group = () => [
    row(4, {
      name: "Aldi | cart",
      cost: 18,
      costFormula: "=30-R[1]C[0]-R[2]C[0]",
      splitGroupId: "4",
    }),
    row(5, {
      name: "Aldi | milk",
      cost: 10.83,
      costFormula: "=1.0825*(10)",
      envelope: "Groceries",
      splitGroupId: "4",
      note: "for Pat",
    }),
    row(6, {
      name: "Aldi | soap",
      cost: 1.17,
      costFormula: null,
      splitGroupId: "4",
    }),
  ];
  const saveClose = () => screen.getByRole("button", { name: "Save & close" });
  const card = (name: string) => screen.getByRole("group", { name });
  const addItem = () => screen.getByRole("button", { name: "Add split item" });

  it("shows a single card with no tax toggle and no split items", () => {
    setup();
    expect(card("Transaction")).toBeInTheDocument();
    expect(screen.queryByLabelText(/Add tax/)).toBeNull();
    expect(screen.queryByTestId("remaining")).toBeNull();
    expect(addItem()).toBeInTheDocument();
  });

  it("assigns an envelope with one tap and closes", async () => {
    const { user, onUpdate, onClose } = setup();
    await user.click(screen.getByRole("button", { name: /Groceries/ }));
    expect(onUpdate).toHaveBeenCalledWith(
      expect.objectContaining({ rowIndex: 0 }),
      { envelope: "Groceries" },
    );
    expect(onClose).toHaveBeenCalled();
  });

  it("saves staged edits together with the envelope on a chip tap", async () => {
    const { user, onUpdate, onClose } = setup();
    await user.type(screen.getByLabelText("Name"), " Market");
    await user.click(screen.getByRole("button", { name: /Groceries/ }));
    expect(onUpdate).toHaveBeenCalledTimes(1);
    expect(onUpdate.mock.calls[0][1]).toEqual({
      name: "Aldi Market",
      envelope: "Groceries",
    });
    expect(onClose).toHaveBeenCalledTimes(1);
  });

  it("shows envelope balances on chips, negative ones in red", () => {
    setup();
    expect(screen.getByRole("button", { name: /Groceries/ })).toHaveTextContent(
      "$280.00",
    );
    expect(
      within(screen.getByRole("button", { name: /Gas/ })).getByText("-$30.00"),
    ).toHaveClass("text-red-600");
  });

  it("expands and collapses a chooser; collapsed it is one scrolling row", async () => {
    const { user } = setup({ paidWith: "Pat Venmo" });
    const chooser = within(screen.getByRole("group", { name: "Paid with" }));
    const toggle = chooser.getByRole("button", { name: "Paid with" });
    const scroller = chooser.getByTestId("chip-scroller");
    expect(toggle).toHaveAttribute("aria-expanded", "false");
    expect(scroller).toHaveClass("overflow-x-auto", "flex-nowrap");
    await user.click(toggle);
    expect(toggle).toHaveAttribute("aria-expanded", "true");
    expect(scroller).toHaveClass("flex-wrap");
    expect(scroller).not.toHaveClass("overflow-x-auto");
    await user.click(toggle);
    expect(toggle).toHaveAttribute("aria-expanded", "false");
    expect(chooser.getByRole("button", { name: "Pat Venmo" })).toHaveAttribute(
      "aria-pressed",
      "true",
    );
  });

  it("stages Paid with from its pills, including an unlisted current value", async () => {
    const { user, onUpdate } = setup({ paidWith: "Cash" });
    const chooser = within(screen.getByRole("group", { name: "Paid with" }));
    expect(chooser.getByRole("button", { name: "Cash" })).toHaveAttribute(
      "aria-pressed",
      "true",
    );
    await user.click(chooser.getByRole("button", { name: "Pat Venmo" }));
    expect(onUpdate).not.toHaveBeenCalled();
    await user.click(saveClose());
    expect(onUpdate.mock.calls[0][1]).toEqual({ paidWith: "Pat Venmo" });
  });

  it("clears Paid with", async () => {
    const { user, onUpdate } = setup();
    await user.click(
      within(screen.getByRole("group", { name: "Paid with" })).getByRole(
        "button",
        { name: "Clear" },
      ),
    );
    await user.click(saveClose());
    expect(onUpdate.mock.calls[0][1]).toEqual({ paidWith: "" });
  });

  it("degrades gracefully without envelope or paid-with options", () => {
    setup(
      {},
      { envelopeOptions: null, envelopes: null, paidWithOptions: null },
    );
    expect(screen.queryByRole("button", { name: /Groceries/ })).toBeNull();
    expect(screen.getByText(/no envelope list/)).toBeInTheDocument();
    expect(screen.queryByLabelText("Paid with")).toBeNull();
  });

  it("stages cost edits with inline math until Save & close", async () => {
    const { user, onUpdate, onClose } = setup();
    const input = screen.getByLabelText("Cost");
    await user.clear(input);
    await user.type(input, "2*3+1");
    expect(onUpdate).not.toHaveBeenCalled();
    expect(screen.getByTestId("dirty-indicator")).toBeInTheDocument();
    await user.click(saveClose());
    expect(onUpdate.mock.calls[0][1]).toEqual({ costExpression: "2*3+1" });
    expect(onClose).toHaveBeenCalled();
  });

  it("blocks an invalid cost", async () => {
    const { user } = setup();
    const input = screen.getByLabelText("Cost");
    await user.clear(input);
    await user.type(input, "2/0");
    expect(saveClose()).toBeDisabled();
  });

  it("edits the name in the title and saves it", async () => {
    const { user, onUpdate } = setup();
    const title = screen.getByLabelText("Name");
    expect(title).toHaveValue("Aldi");
    await user.clear(title);
    await user.type(title, "Aldi Foods");
    await user.click(saveClose());
    expect(onUpdate.mock.calls[0][1]).toEqual({ name: "Aldi Foods" });
  });

  it("edits only the item note of a split-named row, keeping the merchant prefix", async () => {
    const { user, onUpdate } = setup({ name: "Aldi | milk" });
    expect(screen.getByTestId("name-prefix")).toHaveTextContent("Aldi |");
    expect(screen.queryByLabelText("Name")).toBeNull();
    const note = screen.getByLabelText("Item note");
    expect(note).toHaveValue("milk");
    await user.clear(note);
    await user.type(note, "oat milk");
    await user.click(saveClose());
    expect(onUpdate.mock.calls[0][1]).toEqual({ name: "Aldi | oat milk" });
  });

  it("saves only changed details, with Tax deductible as Yes or blank", async () => {
    const { user, onUpdate } = setup();
    await user.click(screen.getByText("Edit details"));
    await user.click(screen.getByLabelText("Tax deductible"));
    await user.click(saveClose());
    expect(onUpdate.mock.calls[0][1]).toEqual({ taxDeductible: "Yes" });
  });

  it("adds a note from Edit details and focuses it", async () => {
    const { user, onUpdate } = setup();
    expect(screen.queryByLabelText("Note")).toBeNull();
    await user.click(screen.getByText("Edit details"));
    await user.click(screen.getByRole("button", { name: "Add note" }));
    const note = screen.getByLabelText("Note");
    expect(note).toHaveFocus();
    expect(screen.queryByRole("button", { name: "Add note" })).toBeNull();
    await user.type(note, "Reimbursed");
    await user.click(saveClose());
    expect(onUpdate.mock.calls[0][1]).toEqual({ note: "Reimbursed" });
  });

  it("shows an existing note in the box and saves its removal", async () => {
    const { user, onUpdate } = setup({ note: "Check this" });
    const note = screen.getByLabelText("Note");
    expect(note).toHaveValue("Check this");
    await user.click(screen.getByText("Edit details"));
    expect(screen.queryByRole("button", { name: "Add note" })).toBeNull();
    await user.clear(note);
    await user.click(saveClose());
    expect(onUpdate.mock.calls[0][1]).toEqual({ note: "" });
  });

  it("blocks a name that would be a formula", async () => {
    const { user } = setup();
    const name = screen.getByLabelText("Name");
    await user.clear(name);
    await user.type(name, "=cmd()");
    expect(saveClose()).toBeDisabled();
    expect(screen.getByRole("alert")).toHaveTextContent(/cannot start/);
    expect(screen.getByRole("button", { name: /Groceries/ })).toBeDisabled();
  });

  it("adds split items with Add tax checked; the original shows the live remainder", async () => {
    const { user, onUpdate, onSplit, onClose } = setup({ cost: 50 });
    await user.click(addItem());
    const original = within(card("Original"));
    expect(screen.queryByLabelText("Cost")).toBeNull();
    expect(original.queryByLabelText(/Add tax/)).toBeNull();
    expect(original.getByTestId("remaining")).toHaveTextContent("$50.00");
    expect(original.queryByText(/remainder from the split items/)).toBeNull();
    await user.click(
      original.getByRole("button", { name: "About this amount" }),
    );
    expect(
      original.getByText(
        "Calculated as the remainder from the split items below.",
      ),
    ).toBeInTheDocument();

    const item1 = within(card("Split item 1"));
    expect(item1.getByLabelText(/Add tax/)).toBeChecked();
    expect(item1.getByText(/can be set after saving/)).toBeInTheDocument();
    expect(item1.queryByText("Edit details")).toBeNull();
    await user.type(item1.getByLabelText("Amount"), "4.23*2+12.09");
    // 1.0825 * 20.55 = 22.245...
    expect(original.getByTestId("remaining")).toHaveTextContent("$27.75");
    await user.click(item1.getByLabelText(/Add tax/));
    expect(original.getByTestId("remaining")).toHaveTextContent("$29.45");
    await user.type(item1.getByLabelText("Item note"), "Napkins");
    await user.click(item1.getByRole("button", { name: /Groceries/ }));
    await user.click(item1.getByRole("button", { name: "Pat Venmo" }));

    await user.click(addItem());
    const item2 = within(card("Split item 2"));
    await user.type(item2.getByLabelText("Amount"), "10");
    expect(original.getByTestId("remaining")).toHaveTextContent("$18.63");
    await user.type(original.getByLabelText("Item note"), "Mine");
    // Not the fast path any more: the tap is staged
    await user.click(original.getByRole("button", { name: /Gas/ }));
    expect(onClose).not.toHaveBeenCalled();

    await user.click(saveClose());
    expect(onUpdate).toHaveBeenCalledTimes(1);
    expect(onUpdate.mock.calls[0][1]).toEqual({ envelope: "Gas" });
    expect(onSplit).toHaveBeenCalledTimes(1);
    expect(onSplit.mock.calls[0][0]).toMatchObject({ rowIndex: 0 });
    expect(onSplit.mock.calls[0][1]).toEqual([
      {
        itemNote: "Napkins",
        expression: "4.23*2+12.09",
        preTax: false,
        envelope: "Groceries",
        paidWith: "Pat Venmo",
      },
      { itemNote: "", expression: "10", preTax: true, envelope: "" },
    ]);
    expect(onSplit.mock.calls[0][2]).toEqual({ itemNote: "Mine" });
    expect(onUpdate.mock.invocationCallOrder[0]).toBeLessThan(
      onSplit.mock.invocationCallOrder[0],
    );
    expect(onClose).toHaveBeenCalledTimes(1);
  });

  it("sends no parent item note when it was left blank, and removes items", async () => {
    const { user, onSplit, onUpdate } = setup({ cost: 20 });
    await user.click(addItem());
    await user.click(addItem());
    await user.click(
      screen.getByRole("button", { name: "Remove split item 1" }),
    );
    expect(screen.getAllByTestId("split-item")).toHaveLength(1);
    await user.type(within(card("Split item 1")).getByLabelText("Amount"), "5");
    await user.click(saveClose());
    expect(onUpdate).not.toHaveBeenCalled();
    expect(onSplit.mock.calls[0][1]).toHaveLength(1);
    expect(onSplit.mock.calls[0][2]).toEqual({});
  });

  it("blocks over-allocation and invalid items", async () => {
    const { user } = setup({ cost: 10 });
    await user.click(addItem());
    expect(saveClose()).toBeDisabled(); // empty amount
    const item = within(card("Split item 1"));
    await user.click(item.getByLabelText(/Add tax/));
    await user.type(item.getByLabelText("Amount"), "10.01");
    expect(screen.getByTestId("remaining")).toHaveTextContent("-$0.01");
    expect(screen.getByTestId("remaining")).toHaveClass("text-red-600");
    expect(screen.getByRole("alert")).toHaveTextContent(/more than the total/);
    expect(saveClose()).toBeDisabled();
    await user.clear(item.getByLabelText("Amount"));
    await user.type(item.getByLabelText("Amount"), "10");
    expect(saveClose()).toBeEnabled();
    await user.type(item.getByLabelText("Item note"), "=x");
    expect(saveClose()).toBeDisabled();
  });

  it("splits an edited cost: the remainder starts from the new cost", async () => {
    const { user, onUpdate, onSplit } = setup({ cost: 10 });
    const cost = screen.getByLabelText("Cost");
    await user.clear(cost);
    await user.type(cost, "12");
    await user.click(addItem());
    expect(screen.getByTestId("remaining")).toHaveTextContent("$12.00");
    const item = within(card("Split item 1"));
    await user.click(item.getByLabelText(/Add tax/));
    await user.type(item.getByLabelText("Amount"), "2");
    expect(screen.getByTestId("remaining")).toHaveTextContent("$10.00");
    await user.click(saveClose());
    expect(onUpdate.mock.calls[0][1]).toEqual({ costExpression: "12" });
    expect(onSplit).toHaveBeenCalledTimes(1);
  });

  it("shows an existing group as cards and edits its rows", async () => {
    const { user, onUpdate, onSplit } = setup({}, { rows: group() });
    expect(screen.getByTestId("name-prefix")).toHaveTextContent("Aldi |");
    expect(screen.queryByLabelText("Name")).toBeNull();
    expect(screen.getByText(/Total \$30\.00/)).toBeInTheDocument();
    const original = within(card("Original"));
    expect(original.getByLabelText("Item note")).toHaveValue("cart");
    expect(original.getByTestId("remaining")).toHaveTextContent("$18.00");
    expect(original.queryByLabelText(/Add tax/)).toBeNull();

    const milk = within(card("Split item 1"));
    expect(milk.getByLabelText("Item note")).toHaveValue("milk");
    expect(milk.getByLabelText("Amount")).toHaveValue("10");
    expect(milk.getByLabelText(/Add tax/)).toBeChecked();
    expect(milk.getByLabelText("Note")).toHaveValue("for Pat");
    const soap = within(card("Split item 2"));
    expect(soap.getByLabelText("Amount")).toHaveValue("1.17");
    expect(soap.getByLabelText(/Add tax/)).not.toBeChecked();
    expect(soap.queryByLabelText("Note")).toBeNull();

    // Editing a child moves the remainder; envelope taps are staged
    await user.clear(milk.getByLabelText("Amount"));
    await user.type(milk.getByLabelText("Amount"), "12");
    expect(original.getByTestId("remaining")).toHaveTextContent("$15.84");
    await user.click(milk.getByRole("button", { name: /Gas/ }));
    await user.type(original.getByLabelText("Item note"), "s");
    await user.click(soap.getByText("Edit details"));
    await user.click(soap.getByLabelText("Tax deductible"));
    await user.click(soap.getByRole("button", { name: "Add note" }));
    await user.type(soap.getByLabelText("Note"), "gift");

    await user.click(saveClose());
    expect(onSplit).not.toHaveBeenCalled();
    expect(
      onUpdate.mock.calls.map((c) => [
        (c[0] as ViewRow).rowIndex,
        c[1] as unknown,
      ]),
    ).toEqual([
      [4, { name: "Aldi | carts" }],
      [5, { costExpression: "12", costPreTax: true, envelope: "Gas" }],
      [6, { taxDeductible: "Yes", note: "gift" }],
    ]);
  });

  it("appends new split items to an existing group", async () => {
    const { user, onUpdate, onSplit } = setup({}, { rows: group() });
    await user.click(addItem());
    const item = within(card("Split item 3"));
    await user.type(item.getByLabelText("Amount"), "2");
    expect(screen.getByTestId("remaining")).toHaveTextContent("$15.84");
    await user.click(saveClose());
    expect(onUpdate).not.toHaveBeenCalled();
    expect(onSplit.mock.calls[0][0]).toMatchObject({ rowIndex: 4 });
    expect(onSplit.mock.calls[0][1]).toEqual([
      { itemNote: "", expression: "2", preTax: true, envelope: "" },
    ]);
    expect(onSplit.mock.calls[0][2]).toEqual({});
  });

  it("disables adding items while the group's split is queued", () => {
    const rows = group();
    rows[0] = { ...rows[0], splitPending: true, pending: true };
    setup({}, { rows });
    expect(screen.queryByRole("button", { name: "Add split item" })).toBeNull();
    expect(screen.getByText("Split queued…")).toBeInTheDocument();
  });

  it("closes without a prompt when nothing changed", async () => {
    const { user, onClose, onUpdate } = setup();
    expect(screen.queryByTestId("dirty-indicator")).toBeNull();
    await user.click(screen.getByRole("button", { name: "Close" }));
    expect(onClose).toHaveBeenCalledTimes(1);
    expect(screen.queryByRole("alertdialog")).toBeNull();
    expect(onUpdate).not.toHaveBeenCalled();
  });

  it("asks Save / Discard / Go back when dismissed with unsaved changes", async () => {
    const { user, onClose, onUpdate } = setup();
    await user.type(screen.getByLabelText("Name"), "x");
    // Escape and the backdrop use the same path as the close button.
    await user.keyboard("{Escape}");
    expect(
      screen.getByRole("alertdialog", { name: "Unsaved changes" }),
    ).toBeInTheDocument();
    expect(onClose).not.toHaveBeenCalled();
    await user.click(screen.getByRole("button", { name: "Go back" }));
    expect(screen.queryByRole("alertdialog")).toBeNull();
    expect(screen.getByLabelText("Name")).toHaveValue("Aldix");

    await user.click(screen.getByRole("button", { name: "Close" }));
    await user.click(screen.getByRole("button", { name: "Discard" }));
    expect(onClose).toHaveBeenCalledTimes(1);
    expect(onUpdate).not.toHaveBeenCalled();
    await user.click(screen.getByRole("button", { name: "Go back" }));

    await user.click(screen.getByRole("button", { name: "Close" }));
    await user.click(screen.getByRole("button", { name: "Save" }));
    expect(onUpdate.mock.calls[0][1]).toEqual({ name: "Aldix" });
    expect(onClose).toHaveBeenCalledTimes(2);
  });

  it("treats added split items as unsaved changes", async () => {
    const { user, onClose } = setup();
    await user.click(addItem());
    expect(screen.getByTestId("dirty-indicator")).toBeInTheDocument();
    await user.keyboard("{Escape}");
    expect(screen.getByRole("alertdialog")).toBeInTheDocument();
    expect(onClose).not.toHaveBeenCalled();
  });

  describe("Delete", () => {
    const deleteButton = (
      name: string | RegExp = /^Delete (transaction|split)$/,
    ) => screen.getByRole("button", { name });

    it("has no Delete button without onDelete", () => {
      setup();
      expect(
        screen.queryByRole("button", { name: /^Delete (transaction|split)$/ }),
      ).toBeNull();
    });

    it("asks for confirmation in an alert dialog and deletes only on Delete", async () => {
      const onDelete = vi.fn();
      const { onClose, user } = setup(
        { cost: 12.5, costDisplay: "$12.50" },
        { onDelete },
      );
      await user.click(deleteButton("Delete transaction"));
      const confirm = screen.getByRole("alertdialog", {
        name: "Delete transaction?",
      });
      expect(confirm).toHaveTextContent("Aldi · $12.50");
      expect(confirm).toHaveTextContent("This deletes the row from the sheet.");
      // Cancel has focus, so Enter or a stray tap does not delete
      expect(
        within(confirm).getByRole("button", { name: "Cancel" }),
      ).toHaveFocus();
      expect(onDelete).not.toHaveBeenCalled();
      await user.click(within(confirm).getByRole("button", { name: "Delete" }));
      expect(onDelete).toHaveBeenCalledTimes(1);
      expect(onDelete.mock.calls[0][0].map((r: ViewRow) => r.rowIndex)).toEqual(
        [0],
      );
      expect(onClose).toHaveBeenCalledTimes(1);
    });

    it("Cancel and Escape close only the confirmation", async () => {
      const onDelete = vi.fn();
      const { onClose, user } = setup({}, { onDelete });
      await user.click(deleteButton("Delete transaction"));
      await user.click(screen.getByRole("button", { name: "Cancel" }));
      expect(screen.queryByRole("alertdialog")).toBeNull();
      await user.click(deleteButton("Delete transaction"));
      await user.keyboard("{Escape}");
      expect(screen.queryByRole("alertdialog")).toBeNull();
      // The drawer stays open and nothing was deleted
      expect(screen.getByRole("dialog")).toBeInTheDocument();
      expect(onClose).not.toHaveBeenCalled();
      expect(onDelete).not.toHaveBeenCalled();
    });

    it("deletes every row of a split group together", async () => {
      const onDelete = vi.fn();
      const { user } = setup({}, { rows: group(), onDelete });
      await user.click(deleteButton("Delete split"));
      const confirm = screen.getByRole("alertdialog", {
        name: "Delete split?",
      });
      expect(confirm).toHaveTextContent("Aldi · $30.00");
      expect(confirm).toHaveTextContent("all 3 rows of this split");
      await user.click(within(confirm).getByRole("button", { name: "Delete" }));
      expect(onDelete.mock.calls[0][0].map((r: ViewRow) => r.rowIndex)).toEqual(
        [4, 5, 6],
      );
    });

    it("is disabled while the group has a queued split", () => {
      setup({ splitPending: true } as Partial<ViewRow>, { onDelete: vi.fn() });
      expect(deleteButton()).toBeDisabled();
      expect(screen.getByText(/once the queued split is saved/)).toBeVisible();
    });
  });

  it("dismisses on dragging the handle down, closing straight away when clean", async () => {
    const { onClose } = setup();
    const handle = screen.getByTestId("drag-handle");
    fireEvent.pointerDown(handle, { clientY: 100, pointerId: 1 });
    fireEvent.pointerMove(handle, { clientY: 260, pointerId: 1 });
    fireEvent.pointerUp(handle, { clientY: 260, pointerId: 1 });
    expect(onClose).toHaveBeenCalledTimes(1);
  });
});

describe("ConflictDialog", () => {
  const base = makeRow(0, { name: "Aldi", envelope: "" });
  const current = makeRow(0, {
    name: "Aldi",
    envelope: "Gas",
    fingerprint: "remote",
  });
  const conflict: Conflict = {
    id: "c1",
    mutation: {
      id: "m1",
      action: "updateTransaction",
      sheetName: "S",
      rowIndex: 0,
      baseFingerprint: base.fingerprint,
      baseRow: base,
      createdAt: 0,
      payload: { changes: { envelope: "Groceries" } },
    },
    current,
  };

  it("shows Original, Discarded Local Change and Current Remote values", () => {
    render(
      <ConflictDialog
        conflict={conflict}
        taxMultiplier={1}
        onResolve={() => {}}
      />,
    );
    const original = screen.getByRole("region", { name: "Original Value" });
    const mine = screen.getByRole("region", { name: "Discarded Local Change" });
    const remote = screen.getByRole("region", { name: "Current Remote Value" });
    // Envelope, Tax deductible and Note are blank in the original
    expect(within(original).getAllByText("(blank)").length).toBe(3);
    expect(within(mine).getByText("Groceries")).toBeInTheDocument();
    expect(within(remote).getByText("Gas")).toBeInTheDocument();
  });

  it("requires an explicit choice", async () => {
    const onResolve = vi.fn();
    const user = userEvent.setup();
    render(
      <ConflictDialog
        conflict={conflict}
        taxMultiplier={1}
        onResolve={onResolve}
      />,
    );
    expect(onResolve).not.toHaveBeenCalled();
    await user.click(screen.getByRole("button", { name: "Accept Remote" }));
    expect(onResolve).toHaveBeenLastCalledWith("accept");
    await user.click(
      screen.getByRole("button", { name: "Overwrite with Mine" }),
    );
    expect(onResolve).toHaveBeenLastCalledWith("overwrite");
  });

  it("shows a delete as the discarded local change", async () => {
    const child = makeRow(1, { name: "Aldi | milk", splitGroupId: "0" });
    const remoteChild = { ...child, name: "Aldi | eggs", fingerprint: "r" };
    const parent = makeRow(0, { name: "Aldi | ", splitGroupId: "0" });
    const onResolve = vi.fn();
    render(
      <ConflictDialog
        conflict={{
          id: "c2",
          mutation: {
            id: "m2",
            action: "deleteTransaction",
            sheetName: "S",
            rowIndex: 0,
            baseFingerprint: parent.fingerprint,
            baseRow: parent,
            createdAt: 0,
            payload: {
              rows: [
                {
                  rowIndex: 0,
                  baseFingerprint: parent.fingerprint,
                  baseRow: parent,
                },
                {
                  rowIndex: 1,
                  baseFingerprint: child.fingerprint,
                  baseRow: child,
                },
              ],
            },
          },
          current: remoteChild,
        }}
        taxMultiplier={1}
        onResolve={onResolve}
      />,
    );
    expect(
      screen.getByText(/delete of all 2 rows of this split/),
    ).toBeVisible();
    // The original is the row that changed, not the first row of the group
    const original = screen.getByRole("region", { name: "Original Value" });
    expect(within(original).getByText("Aldi | milk")).toBeInTheDocument();
    const mine = screen.getByRole("region", { name: "Discarded Local Change" });
    expect(mine).toHaveTextContent("Delete (all 2 rows)");
    const remote = screen.getByRole("region", { name: "Current Remote Value" });
    expect(within(remote).getByText("Aldi | eggs")).toBeInTheDocument();
    await userEvent
      .setup()
      .click(screen.getByRole("button", { name: "Overwrite with Mine" }));
    expect(onResolve).toHaveBeenCalledWith("overwrite");
  });

  it("cannot overwrite a row that no longer exists", () => {
    render(
      <ConflictDialog
        conflict={{ ...conflict, current: null }}
        taxMultiplier={1}
        onResolve={() => {}}
      />,
    );
    expect(
      screen.getByRole("button", { name: "Overwrite with Mine" }),
    ).toBeDisabled();
  });
});

describe("ExpensesView", () => {
  const snapshot = {
    all: null,
    loadingAll: false,
    queue: [],
    conflicts: [],
    offline: false,
    recentlySaved: [],
  };
  function setup(state: PayPeriodStateDto | undefined) {
    const engine = {
      loadAll: vi.fn(),
      enqueueUpdate: vi.fn(),
      enqueueSplit: vi.fn(),
      enqueueDelete: vi.fn(),
    };
    const sheetName = "3/3/26 - 3/16/26";
    render(
      <ExpensesView
        engine={engine as unknown as SyncEngine}
        sheetName={sheetName}
        state={state}
        rows={buildViewRows(state, [])}
        showAll={false}
        snapshot={snapshot}
      />,
    );
    return { engine, user: userEvent.setup() };
  }

  it("keeps search and filters while the pay period loads", () => {
    setup(undefined);
    expect(screen.getByLabelText("Search")).toBeInTheDocument();
    expect(screen.getByRole("tab", { name: /Uncategorized/ })).toBeVisible();
    expect(screen.getByRole("status")).toHaveTextContent(
      "Loading this pay period…",
    );
  });

  it("opens a child row with its whole group and saves each row with its server copy", async () => {
    const state = makeState([
      makeRow(0, {
        name: "Aldi | ",
        cost: 8,
        costFormula: "=10-R[1]C[0]",
        splitGroupId: "0",
      }),
      makeRow(1, { name: "Aldi | milk", cost: 2, splitGroupId: "0" }),
    ]);
    const { engine, user } = setup(state);
    await user.click(screen.getByRole("button", { name: /Aldi \| milk/ }));
    const dialog = within(screen.getByRole("dialog"));
    expect(dialog.getByRole("group", { name: "Original" })).toBeInTheDocument();
    const child = within(dialog.getByRole("group", { name: "Split item 1" }));
    await user.click(child.getByRole("button", { name: /Groceries/ }));
    await user.click(dialog.getByRole("button", { name: "Add split item" }));
    await user.type(
      within(
        dialog.getByRole("group", { name: "Split item 2" }),
      ).getByLabelText("Amount"),
      "1",
    );
    await user.click(dialog.getByRole("button", { name: "Save & close" }));
    expect(engine.enqueueUpdate).toHaveBeenCalledWith(
      state.sheetName,
      state.transactions[1],
      { envelope: "Groceries" },
    );
    expect(engine.enqueueSplit).toHaveBeenCalledWith(
      state.sheetName,
      state.transactions[0],
      [{ itemNote: "", expression: "1", preTax: true, envelope: "" }],
      {},
    );
  });

  it("deletes a split group with the server copies of all its rows", async () => {
    const state = makeState([
      makeRow(0, { name: "Coffee" }),
      makeRow(1, {
        name: "Aldi | ",
        cost: 8,
        costFormula: "=10-R[1]C[0]",
        splitGroupId: "1",
      }),
      makeRow(2, { name: "Aldi | milk", cost: 2, splitGroupId: "1" }),
    ]);
    const { engine, user } = setup(state);
    await user.click(screen.getByRole("button", { name: /Aldi \| milk/ }));
    await user.click(screen.getByRole("button", { name: "Delete split" }));
    expect(engine.enqueueDelete).not.toHaveBeenCalled();
    await user.click(
      within(screen.getByRole("alertdialog")).getByRole("button", {
        name: "Delete",
      }),
    );
    expect(engine.enqueueDelete).toHaveBeenCalledWith(state.sheetName, [
      state.transactions[1],
      state.transactions[2],
    ]);
  });
});

describe("EnvelopesView", () => {
  it("shows the core four metrics with Remain emphasized and negatives in red", () => {
    render(<EnvelopesView state={makeState([])} warnings={[]} />);
    expect(screen.getByTestId("remain-Groceries")).toHaveTextContent("$280.00");
    expect(screen.getByTestId("remain-Gas")).toHaveClass("text-red-600");
    expect(screen.queryByText("Spent YTD")).toBeNull();
  });

  it("toggles the full table", async () => {
    const user = userEvent.setup();
    render(<EnvelopesView state={makeState([])} warnings={[]} />);
    await user.click(screen.getByLabelText("Show all columns"));
    expect(
      screen.getByRole("columnheader", { name: "Spent YTD" }),
    ).toBeInTheDocument();
  });

  it("explains when the sheet has no envelope variables", () => {
    render(
      <EnvelopesView
        state={makeState([], { envelopes: null })}
        warnings={["EnvelopesStart missing"]}
      />,
    );
    expect(screen.getByText(/No envelope data/)).toBeInTheDocument();
    expect(screen.getByText("EnvelopesStart missing")).toBeInTheDocument();
  });
});

describe("StatusPill", () => {
  const base: EngineSnapshot = {
    ready: true,
    currentSheet: null,
    states: {},
    all: null,
    loadingAll: false,
    queue: [],
    conflicts: [],
    recentlySaved: [],
    refreshing: false,
    flushing: false,
    syncingReceipts: false,
    offline: false,
    notice: null,
    warnings: [],
  };
  const queued = [{} as never, {} as never];

  it("reports each state", () => {
    expect(statusOf(base).label).toBe("Synced");
    expect(statusOf({ ...base, refreshing: true }).label).toBe("Refreshing…");
    expect(statusOf({ ...base, queue: queued }).label).toBe(
      "Syncing 2 changes…",
    );
    expect(statusOf({ ...base, queue: queued, offline: true }).label).toBe(
      "Offline (2 queued)",
    );
    expect(
      statusOf({ ...base, conflicts: [{} as never], offline: true }).label,
    ).toBe("Conflict");
  });

  it("renders", async () => {
    render(<StatusPill snapshot={{ ...base, offline: true }} />);
    await waitFor(() =>
      expect(screen.getByRole("status")).toHaveTextContent("Offline"),
    );
  });
});
