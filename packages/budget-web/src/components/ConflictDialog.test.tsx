import { describe, expect, it, vi } from "vitest";
import { render, screen, within } from "@testing-library/react";
import userEvent from "@testing-library/user-event";
import { ConflictDialog } from "./ConflictDialog.tsx";
import { makeRow } from "../test/fixtures.ts";
import type { Conflict, GroupMutation } from "../sync/types.ts";

describe("ConflictDialog for a split group save", () => {
  /** Aldi, $30 in total, split into milk and soap. */
  const base = [
    makeRow(0, {
      name: "Aldi | ",
      cost: 18,
      costFormula: "=30-R[1]C[0]-R[2]C[0]",
      splitGroupId: "0",
    }),
    makeRow(1, { name: "Aldi | milk", cost: 10.83, splitGroupId: "0" }),
    makeRow(2, { name: "Aldi | soap", cost: 1.17, splitGroupId: "0" }),
  ];
  const mutation: GroupMutation = {
    id: "m1",
    action: "updateSplitGroup",
    sheetName: "S",
    rowIndex: 0,
    baseFingerprint: base[0].fingerprint,
    baseRow: base[0],
    createdAt: 0,
    payload: {
      rows: base.map((r, i) => ({
        rowIndex: r.rowIndex,
        baseFingerprint: r.fingerprint,
        baseRow: r,
        ...(i === 2 ? { changes: { costExpression: "5" } } : {}),
      })),
      title: "Lidl",
    },
  };
  const remote = base.map((r) =>
    r.rowIndex === 1 ? { ...r, envelope: "Gas", fingerprint: "remote1" } : r,
  );
  const conflict = (over: Partial<Conflict> = {}): Conflict => ({
    id: "c1",
    mutation,
    current: remote[1],
    currentRows: remote,
    ...over,
  });

  function setup(c: Conflict = conflict()) {
    const onResolve = vi.fn();
    render(
      <ConflictDialog
        conflict={c}
        taxMultiplier={1.0825}
        onResolve={onResolve}
      />,
    );
    return { onResolve, user: userEvent.setup() };
  }

  /** The cells of one version ("Original", "Mine", "Remote") of a row's table. */
  const version = (rowLabel: string, versionLabel: RegExp) =>
    within(
      within(screen.getByRole("region", { name: rowLabel })).getByRole("row", {
        name: versionLabel,
      }),
    );

  it("shows Original, Discarded Local Change and Current Remote for every row of the group", () => {
    setup();
    expect(
      screen.getByRole("heading", { name: "This split changed in the sheet" }),
    ).toBeInTheDocument();
    for (const label of ["Remainder", "Split item 1", "Split item 2"])
      expect(screen.getByRole("region", { name: label })).toBeInTheDocument();
    // My change: the new title on every row, soap's cost, and the parent's remainder
    expect(
      version("Remainder", /Discarded/).getByText("Lidl |"),
    ).toBeInTheDocument();
    expect(
      version("Split item 2", /Discarded/).getByText("Lidl | soap"),
    ).toBeInTheDocument();
    expect(
      version("Split item 2", /Discarded/).getByText("5"),
    ).toBeInTheDocument();
    expect(
      version("Remainder", /Discarded/).getByText(/14\.17/),
    ).toBeInTheDocument();
    // Their change: milk's envelope
    const theirs = version("Split item 1", /Current remote/).getByText("Gas");
    expect(theirs).toHaveAttribute("data-changed", "true");
    expect(
      version("Split item 1", /Original/).getByText("Aldi | milk"),
    ).not.toHaveAttribute("data-changed");
  });

  it("resolves the whole group with one choice", async () => {
    const { onResolve, user } = setup();
    await user.click(
      screen.getByRole("button", { name: "Overwrite with Mine" }),
    );
    expect(onResolve).toHaveBeenCalledWith("overwrite");
    await user.click(screen.getByRole("button", { name: "Accept Remote" }));
    expect(onResolve).toHaveBeenCalledWith("accept");
  });

  it("shows a row that is gone and then cannot overwrite", () => {
    setup(conflict({ currentRows: remote.slice(0, 2) }));
    expect(
      within(screen.getByRole("region", { name: "Split item 2" })).getByText(
        "This row is no longer in the sheet.",
      ),
    ).toBeInTheDocument();
    expect(
      screen.getByRole("button", { name: "Overwrite with Mine" }),
    ).toBeDisabled();
  });

  it("lists new split items of the discarded change", () => {
    setup(
      conflict({
        mutation: {
          ...mutation,
          payload: {
            ...mutation.payload,
            newItems: [
              { itemNote: "jam", expression: "4", preTax: false, envelope: "" },
            ],
          },
        },
      }),
    );
    expect(
      screen.getByText(/Your discarded change also added 1 new item: jam 4/),
    ).toBeInTheDocument();
  });
});
