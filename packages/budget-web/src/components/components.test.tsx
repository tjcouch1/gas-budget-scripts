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
import { SplitEditor } from "./SplitEditor.tsx";
import { TransactionSheet } from "./TransactionSheet.tsx";
import { StatusPill, statusOf } from "./StatusPill.tsx";
import { makeRow, makeState } from "../test/fixtures.ts";
import type { EngineSnapshot, Conflict } from "../sync/types.ts";

describe("SplitEditor", () => {
  function setup(
    parentValue: number | null = 50,
    parentName = "Target",
    parentEnvelope = "",
  ) {
    const onSubmit = vi.fn();
    const onClose = vi.fn();
    render(
      <SplitEditor
        parentName={parentName}
        parentValue={parentValue}
        parentEnvelope={parentEnvelope}
        taxMultiplier={1.0825}
        envelopeOptions={["Groceries", "Gas"]}
        onSubmit={onSubmit}
        onClose={onClose}
      />,
    );
    return { onSubmit, onClose, user: userEvent.setup() };
  }
  const saveSplit = () => screen.getByRole("button", { name: "Save split" });

  it("shows live remaining-to-allocate from inline math", async () => {
    const { user } = setup();
    expect(screen.getByTestId("remaining")).toHaveValue("$50.00");
    await user.type(screen.getByLabelText("Item 1 amount"), "4.23*2+12.09");
    expect(screen.getByLabelText("Item 1 value")).toHaveTextContent("$20.55");
    expect(screen.getByTestId("remaining")).toHaveValue("$29.45");
  });

  it("applies the tax multiplier for pre-tax items", async () => {
    const { user } = setup(20);
    await user.type(screen.getByLabelText("Item 1 amount"), "10");
    // The first (original) card's tax toggle is disabled; item 1's is the second.
    await user.click(screen.getAllByLabelText(/Add tax/)[1]);
    expect(screen.getByLabelText("Item 1 value")).toHaveTextContent("$10.83");
    expect(screen.getByTestId("remaining")).toHaveValue("$9.18");
  });

  it("disables Save split for invalid expressions and over-allocation", async () => {
    const { user } = setup(10);
    expect(saveSplit()).toBeDisabled();
    await user.type(screen.getByLabelText("Item 1 amount"), "1+");
    expect(saveSplit()).toBeDisabled();
    await user.click(screen.getByRole("button", { name: "+ Add item" }));
    await user.type(screen.getByLabelText("Item 2 amount"), "2");
    expect(saveSplit()).toBeDisabled(); // item 1 still invalid
    await user.type(screen.getByLabelText("Item 1 amount"), "1");
    expect(saveSplit()).toBeEnabled();
    await user.clear(screen.getByLabelText("Item 2 amount"));
    await user.type(screen.getByLabelText("Item 2 amount"), "99");
    expect(saveSplit()).toBeDisabled();
    expect(screen.getByTestId("remaining")).toHaveClass("text-red-600");
    expect(screen.getByRole("alert")).toHaveTextContent(
      /more than the original/,
    );
  });

  it("starts with a locked-cost original card and one empty item card", () => {
    setup(50, "Target | cart", "Gas");
    const original = screen.getByTestId("original-entry");
    expect(original).toHaveTextContent("Remaining on original");
    // Cost and tax are read-only; note and envelope are editable.
    expect(within(original).getByLabelText("Original amount")).toBeDisabled();
    expect(within(original).getByLabelText(/Add tax/)).toBeDisabled();
    expect(within(original).getByLabelText("Original note")).toHaveValue(
      "cart",
    );
    expect(within(original).getByTestId("name-prefix")).toHaveTextContent(
      "Target |",
    );
    expect(within(original).getByLabelText("Original envelope")).toHaveValue(
      "Gas",
    );
    expect(
      within(original).queryByRole("button", { name: /Remove/ }),
    ).toBeNull();
    expect(screen.getByTestId("remaining")).toHaveValue("$50.00");
    expect(screen.getAllByTestId("split-item")).toHaveLength(1);
    expect(screen.getByLabelText("Item 1 amount")).toBeEnabled();
    expect(screen.queryByRole("button", { name: /Remove/ })).toBeNull(); // only one item
  });

  it("puts each item's remove button on the card, not in the note input", async () => {
    const { user } = setup(50);
    await user.click(screen.getByRole("button", { name: "+ Add item" }));
    const cards = screen.getAllByTestId("split-item");
    expect(cards).toHaveLength(2);
    for (const card of cards) {
      const remove = within(card).getByRole("button", { name: /Remove item/ });
      expect(remove.parentElement?.parentElement).toBe(card);
      expect(card.querySelector("input")?.contains(remove)).toBe(false);
    }
    expect(
      within(screen.getByTestId("original-entry")).queryByRole("button", {
        name: /Remove/,
      }),
    ).toBeNull();
  });

  it("adds and removes items, each reducing the original", async () => {
    const { user } = setup(50);
    await user.type(screen.getByLabelText("Item 1 amount"), "10");
    await user.click(screen.getByRole("button", { name: "+ Add item" }));
    await user.type(screen.getByLabelText("Item 2 amount"), "15");
    expect(screen.getByTestId("remaining")).toHaveValue("$25.00");
    await user.click(screen.getByRole("button", { name: "Remove item 1" }));
    expect(screen.getByTestId("remaining")).toHaveValue("$35.00");
    expect(screen.queryByLabelText("Item 2 amount")).toBeNull();
  });

  it("submits the split items with envelopes, sending no parent fields when unchanged", async () => {
    const { user, onSubmit } = setup(50);
    await user.type(screen.getByLabelText("Item 1 note"), "milk");
    await user.type(screen.getByLabelText("Item 1 amount"), "3.5");
    await user.selectOptions(
      screen.getByLabelText("Item 1 envelope"),
      "Groceries",
    );
    await user.click(screen.getByRole("button", { name: "+ Add item" }));
    await user.type(screen.getByLabelText("Item 2 amount"), "4*2");
    await user.click(screen.getAllByLabelText(/Add tax/)[2]);
    await user.click(saveSplit());
    expect(onSubmit).toHaveBeenCalledWith(
      [
        {
          itemNote: "milk",
          expression: "3.5",
          preTax: false,
          envelope: "Groceries",
        },
        { itemNote: "", expression: "4*2", preTax: true, envelope: "" },
      ],
      {},
    );
  });

  it("submits the original's edited note and envelope", async () => {
    const { user, onSubmit } = setup(50);
    await user.type(screen.getByLabelText("Original note"), "household");
    await user.selectOptions(screen.getByLabelText("Original envelope"), "Gas");
    await user.type(screen.getByLabelText("Item 1 amount"), "5");
    await user.click(saveSplit());
    expect(onSubmit).toHaveBeenCalledWith(expect.any(Array), {
      itemNote: "household",
      envelope: "Gas",
    });
  });

  it("rejects notes that would be formulas", async () => {
    const { user } = setup(50);
    await user.type(screen.getByLabelText("Item 1 note"), "=1+1");
    await user.type(screen.getByLabelText("Item 1 amount"), "1");
    expect(saveSplit()).toBeDisabled();
    await user.clear(screen.getByLabelText("Item 1 note"));
    expect(saveSplit()).toBeEnabled();
    await user.type(screen.getByLabelText("Original note"), "@x");
    expect(saveSplit()).toBeDisabled();
  });

  it("closes at once when nothing changed, otherwise asks Save / Discard / Go back", async () => {
    const { user, onClose, onSubmit } = setup(50);
    await user.click(screen.getByRole("button", { name: "Close" }));
    expect(onClose).toHaveBeenCalledTimes(1);

    await user.type(screen.getByLabelText("Item 1 amount"), "5");
    expect(screen.getByTestId("dirty-indicator")).toBeInTheDocument();
    await user.keyboard("{Escape}");
    const prompt = screen.getByRole("alertdialog", { name: "Unsaved changes" });
    expect(onClose).toHaveBeenCalledTimes(1);
    await user.click(within(prompt).getByRole("button", { name: "Go back" }));
    expect(screen.queryByRole("alertdialog")).toBeNull();

    await user.click(screen.getByRole("button", { name: "Close" }));
    await user.click(screen.getByRole("button", { name: "Discard" }));
    expect(onClose).toHaveBeenCalledTimes(2);
    expect(onSubmit).not.toHaveBeenCalled();
    // The mocked onClose does not unmount, so dismiss the prompt again.
    await user.click(screen.getByRole("button", { name: "Go back" }));

    await user.click(screen.getByRole("button", { name: "Close" }));
    await user.click(screen.getByRole("button", { name: "Save" }));
    expect(onSubmit).toHaveBeenCalledTimes(1);
  });
});

describe("TransactionSheet", () => {
  function setup(
    rowOver = {},
    props: Partial<Parameters<typeof TransactionSheet>[0]> = {},
  ) {
    const onChange = vi.fn();
    const onSplit = vi.fn();
    const onClose = vi.fn();
    const state = makeState([]);
    render(
      <TransactionSheet
        row={{ ...makeRow(0, { name: "Aldi", ...rowOver }), key: "row:0" }}
        isSplitParent={false}
        isSplitChild={false}
        taxMultiplier={1.0825}
        envelopeOptions={["Groceries", "Gas"]}
        envelopes={state.envelopes}
        paidWithOptions={["Credit", "Pat Venmo"]}
        onChange={onChange}
        onSplit={onSplit}
        onClose={onClose}
        {...props}
      />,
    );
    return { onChange, onSplit, onClose, user: userEvent.setup() };
  }
  const saveClose = () => screen.getByRole("button", { name: "Save & close" });

  it("assigns an envelope with one tap and closes", async () => {
    const { user, onChange, onClose } = setup();
    await user.click(screen.getByRole("button", { name: /Groceries/ }));
    expect(onChange).toHaveBeenCalledWith({ envelope: "Groceries" });
    expect(onClose).toHaveBeenCalled();
  });

  it("saves staged edits together with the envelope on a chip tap", async () => {
    const { user, onChange, onClose } = setup();
    await user.type(screen.getByLabelText("Name"), " Market");
    await user.click(screen.getByRole("button", { name: /Groceries/ }));
    expect(onChange).toHaveBeenCalledTimes(1);
    expect(onChange).toHaveBeenCalledWith({
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

  it("degrades gracefully without envelope or paid-with options", () => {
    setup(
      {},
      { envelopeOptions: null, envelopes: null, paidWithOptions: null },
    );
    expect(screen.queryByRole("button", { name: /Groceries/ })).toBeNull();
    expect(screen.getByText(/no envelope list/)).toBeInTheDocument();
    expect(screen.queryByLabelText("Paid with")).toBeNull();
  });

  it("stages cost edits (inline math and tax toggle) until Save & close", async () => {
    const { user, onChange, onClose } = setup();
    const input = screen.getByLabelText("Cost");
    await user.clear(input);
    await user.type(input, "2*3+1");
    await user.click(screen.getByLabelText(/Add tax/));
    expect(onChange).not.toHaveBeenCalled();
    expect(screen.getByTestId("dirty-indicator")).toBeInTheDocument();
    await user.click(saveClose());
    expect(onChange).toHaveBeenCalledWith({
      costExpression: "2*3+1",
      costPreTax: true,
    });
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
    const { user, onChange } = setup();
    const title = screen.getByLabelText("Name");
    expect(title).toHaveValue("Aldi");
    await user.clear(title);
    await user.type(title, "Aldi Foods");
    await user.click(saveClose());
    expect(onChange).toHaveBeenCalledWith({ name: "Aldi Foods" });
  });

  it("edits only the item note of a split row, keeping the merchant prefix", async () => {
    const { user, onChange } = setup(
      { splitGroupId: "g", name: "Aldi | milk" },
      { isSplitChild: true },
    );
    expect(screen.getByTestId("name-prefix")).toHaveTextContent("Aldi |");
    const note = screen.getByLabelText("Item note");
    expect(note).toHaveValue("milk");
    await user.clear(note);
    await user.type(note, "oat milk");
    await user.click(saveClose());
    expect(onChange).toHaveBeenCalledWith({ name: "Aldi | oat milk" });
  });

  it("keeps the split parent's note editable but not its cost", async () => {
    const { user, onChange } = setup(
      {
        splitGroupId: "g",
        name: "Aldi | ",
        costFormula: "=10-R[1]C[0]",
      },
      { isSplitParent: true },
    );
    expect(screen.queryByLabelText("Cost")).toBeNull();
    expect(screen.getByText(/split parent/)).toBeInTheDocument();
    await user.type(screen.getByLabelText("Item note"), "cart");
    await user.click(saveClose());
    expect(onChange).toHaveBeenCalledWith({ name: "Aldi | cart" });
  });

  it("opens the split editor from the Split button", async () => {
    const { user, onSplit, onClose } = setup();
    await user.click(screen.getByRole("button", { name: "Split" }));
    await user.type(screen.getByLabelText("Item 1 amount"), "3");
    await user.click(screen.getByRole("button", { name: "+ Add item" }));
    await user.type(screen.getByLabelText("Item 2 amount"), "2");
    await user.click(screen.getByRole("button", { name: "Save split" }));
    expect(onSplit).toHaveBeenCalledTimes(1);
    expect(onSplit.mock.calls[0][0]).toHaveLength(2);
    expect(onSplit.mock.calls[0][1]).toEqual({});
    expect(onClose).toHaveBeenCalled();
  });

  it("hides Split on split children", () => {
    setup({ splitGroupId: "g", name: "Aldi | milk" }, { isSplitChild: true });
    expect(screen.queryByRole("button", { name: "Split" })).toBeNull();
  });

  it("saves only changed details, with Tax deductible as Yes or blank", async () => {
    const { user, onChange } = setup();
    await user.click(screen.getByText("Edit details"));
    await user.selectOptions(screen.getByLabelText("Paid with"), "Pat Venmo");
    await user.click(screen.getByLabelText("Tax deductible"));
    await user.click(saveClose());
    expect(onChange).toHaveBeenCalledWith({
      paidWith: "Pat Venmo",
      taxDeductible: "Yes",
    });
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

  it("closes without a prompt when nothing changed", async () => {
    const { user, onClose, onChange } = setup();
    expect(screen.queryByTestId("dirty-indicator")).toBeNull();
    await user.click(screen.getByRole("button", { name: "Close" }));
    expect(onClose).toHaveBeenCalledTimes(1);
    expect(screen.queryByRole("alertdialog")).toBeNull();
    expect(onChange).not.toHaveBeenCalled();
  });

  it("asks Save / Discard / Go back when dismissed with unsaved changes", async () => {
    const { user, onClose, onChange } = setup();
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
    expect(onChange).not.toHaveBeenCalled();
    await user.click(screen.getByRole("button", { name: "Go back" }));

    await user.click(screen.getByRole("button", { name: "Close" }));
    await user.click(screen.getByRole("button", { name: "Save" }));
    expect(onChange).toHaveBeenCalledWith({ name: "Aldix" });
    expect(onClose).toHaveBeenCalledTimes(2);
  });

  it("dismisses on dragging the handle down, with the and closes straight away when clean", async () => {
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
    expect(within(original).getAllByText("(blank)").length).toBe(2);
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
