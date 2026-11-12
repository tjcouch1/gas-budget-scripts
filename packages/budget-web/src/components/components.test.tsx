import { describe, expect, it, vi } from "vitest";
import { render, screen, within, waitFor } from "@testing-library/react";
import userEvent from "@testing-library/user-event";
import { ConflictDialog } from "./ConflictDialog.tsx";
import { EnvelopesView } from "./EnvelopesView.tsx";
import { SplitEditor } from "./SplitEditor.tsx";
import { TransactionSheet } from "./TransactionSheet.tsx";
import { StatusPill, statusOf } from "./StatusPill.tsx";
import { makeRow, makeState } from "../test/fixtures.ts";
import type { EngineSnapshot, Conflict } from "../sync/types.ts";

describe("SplitEditor", () => {
  function setup(parentValue: number | null = 50) {
    const onSubmit = vi.fn();
    render(
      <SplitEditor
        parentName="Target"
        parentValue={parentValue}
        taxMultiplier={1.0825}
        envelopeOptions={["Groceries", "Gas"]}
        onSubmit={onSubmit}
        onCancel={() => {}}
      />,
    );
    return { onSubmit, user: userEvent.setup() };
  }

  it("shows live remaining-to-allocate from inline math", async () => {
    const { user } = setup();
    expect(screen.getByTestId("remaining")).toHaveTextContent("$50.00");
    await user.type(screen.getByLabelText("Item 1 amount"), "4.23*2+12.09");
    expect(screen.getByLabelText("Item 1 value")).toHaveTextContent("$20.55");
    expect(screen.getByTestId("remaining")).toHaveTextContent("$29.45");
  });

  it("applies the tax multiplier for pre-tax items", async () => {
    const { user } = setup(20);
    await user.type(screen.getByLabelText("Item 1 amount"), "10");
    await user.click(screen.getAllByLabelText(/Add tax/)[0]);
    expect(screen.getByLabelText("Item 1 value")).toHaveTextContent("$10.83");
    expect(screen.getByTestId("remaining")).toHaveTextContent("$9.18");
  });

  it("disables Split for invalid expressions and over-allocation", async () => {
    const { user } = setup(10);
    const submit = screen.getByRole("button", { name: "Split" });
    expect(submit).toBeDisabled();
    await user.type(screen.getByLabelText("Item 1 amount"), "1+");
    expect(submit).toBeDisabled();
    await user.click(screen.getByRole("button", { name: "+ Add item" }));
    await user.type(screen.getByLabelText("Item 2 amount"), "2");
    expect(submit).toBeDisabled(); // item 1 still invalid
    await user.type(screen.getByLabelText("Item 1 amount"), "1");
    expect(submit).toBeEnabled();
    await user.clear(screen.getByLabelText("Item 2 amount"));
    await user.type(screen.getByLabelText("Item 2 amount"), "99");
    expect(submit).toBeDisabled();
    expect(screen.getByTestId("remaining")).toHaveClass("text-red-600");
    expect(screen.getByRole("alert")).toHaveTextContent(
      /more than the original/,
    );
  });

  it("starts with the original entry and one empty item", () => {
    setup(50);
    const original = screen.getByTestId("original-entry");
    expect(original).toHaveTextContent("Remaining on original");
    expect(within(original).queryByRole("textbox")).toBeNull(); // locked
    expect(screen.getByTestId("remaining")).toHaveTextContent("$50.00");
    expect(screen.getByLabelText("Item 1 amount")).toBeInTheDocument();
    expect(screen.queryByLabelText("Item 2 amount")).toBeNull();
    expect(screen.queryByRole("button", { name: "Fill rest" })).toBeNull();
    expect(screen.queryByRole("button", { name: "Remove item" })).toBeNull();
  });

  it("adds and removes items, each reducing the original", async () => {
    const { user } = setup(50);
    await user.type(screen.getByLabelText("Item 1 amount"), "10");
    await user.click(screen.getByRole("button", { name: "+ Add item" }));
    await user.type(screen.getByLabelText("Item 2 amount"), "15");
    expect(screen.getByTestId("remaining")).toHaveTextContent("$25.00");
    await user.click(screen.getAllByRole("button", { name: "Remove item" })[0]);
    expect(screen.getByTestId("remaining")).toHaveTextContent("$35.00");
    expect(screen.queryByLabelText("Item 2 amount")).toBeNull();
  });

  it("submits the split items with envelopes", async () => {
    const { user, onSubmit } = setup(50);
    await user.type(screen.getByLabelText("Item 1 note"), "milk");
    await user.type(screen.getByLabelText("Item 1 amount"), "3.5");
    await user.selectOptions(
      screen.getByLabelText("Item 1 envelope"),
      "Groceries",
    );
    await user.click(screen.getByRole("button", { name: "+ Add item" }));
    await user.type(screen.getByLabelText("Item 2 amount"), "4*2");
    await user.click(screen.getAllByLabelText(/Add tax/)[1]);
    await user.click(screen.getByRole("button", { name: "Split" }));
    expect(onSubmit).toHaveBeenCalledWith([
      {
        itemNote: "milk",
        expression: "3.5",
        preTax: false,
        envelope: "Groceries",
      },
      { itemNote: "", expression: "4*2", preTax: true, envelope: "" },
    ]);
  });

  it("rejects an item note that would be a formula", async () => {
    const { user } = setup(50);
    await user.type(screen.getByLabelText("Item 1 note"), "=1+1");
    await user.type(screen.getByLabelText("Item 1 amount"), "1");
    expect(screen.getByRole("button", { name: "Split" })).toBeDisabled();
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

  it("assigns an envelope with one tap and closes", async () => {
    const { user, onChange, onClose } = setup();
    await user.click(screen.getByRole("button", { name: /Groceries/ }));
    expect(onChange).toHaveBeenCalledWith({ envelope: "Groceries" });
    expect(onClose).toHaveBeenCalled();
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

  it("edits cost with inline math and a tax toggle", async () => {
    const { user, onChange } = setup();
    const input = screen.getByLabelText("Cost");
    await user.clear(input);
    await user.type(input, "2*3+1");
    expect(screen.getByRole("button", { name: "Save cost" })).toBeEnabled();
    await user.click(screen.getByLabelText(/Add tax/));
    await user.click(screen.getByRole("button", { name: "Save cost" }));
    expect(onChange).toHaveBeenCalledWith({
      costExpression: "2*3+1",
      costPreTax: true,
    });
  });

  it("blocks an invalid cost", async () => {
    const { user } = setup();
    const input = screen.getByLabelText("Cost");
    await user.clear(input);
    await user.type(input, "2/0");
    expect(screen.getByRole("button", { name: "Save cost" })).toBeDisabled();
  });

  it("opens the split editor from the Split button", async () => {
    const { user, onSplit, onClose } = setup();
    await user.click(screen.getByRole("button", { name: "Split" }));
    await user.type(screen.getByLabelText("Item 1 amount"), "3");
    await user.click(screen.getByRole("button", { name: "+ Add item" }));
    await user.type(screen.getByLabelText("Item 2 amount"), "2");
    await user.click(screen.getByRole("button", { name: "Split" }));
    expect(onSplit).toHaveBeenCalledTimes(1);
    expect(onSplit.mock.calls[0][0]).toHaveLength(2);
    expect(onClose).toHaveBeenCalled();
  });

  it("hides cost editing for split parents and name editing for split rows", () => {
    setup({ splitGroupId: "g", name: "Aldi | " }, { isSplitParent: true });
    expect(screen.queryByLabelText("Cost")).toBeNull();
    expect(screen.getByText(/split parent/)).toBeInTheDocument();
    expect(screen.getByLabelText("Name")).toBeDisabled();
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
    await user.click(screen.getByRole("button", { name: "Save details" }));
    expect(onChange).toHaveBeenCalledWith({
      paidWith: "Pat Venmo",
      taxDeductible: "Yes",
    });
  });

  it("blocks a name that would be a formula", async () => {
    const { user } = setup();
    await user.click(screen.getByText("Edit details"));
    const name = screen.getByLabelText("Name");
    await user.clear(name);
    await user.type(name, "=cmd()");
    expect(screen.getByRole("button", { name: "Save details" })).toBeDisabled();
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
