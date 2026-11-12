import { describe, expect, it } from "vitest";
import { makeRow, makeState } from "../test/fixtures.ts";
import {
  groupRows,
  isSplitChild,
  isSplitParent,
  orderEnvelopeOptions,
  stateRows,
  type ViewRow,
} from "./rows.ts";

const rows = stateRows(
  makeState([
    makeRow(0, { name: "Aldi", envelope: "Groceries" }),
    makeRow(1, { name: "Target | milk", splitGroupId: "g1" }),
    makeRow(2, { name: "Target | toy", splitGroupId: "g1", envelope: "Gas" }),
    makeRow(3, { name: "Shell", taxDeductible: "Yes", envelope: "Gas" }),
    makeRow(4, { name: "Broken", isError: true, envelope: "Gas" }),
  ]),
);

describe("groupRows", () => {
  it("nests split groups and keeps them whole when any member matches", () => {
    const groups = groupRows(rows, "uncategorized", "");
    expect(groups.map((g) => g.rows.map((r) => r.name))).toEqual([
      ["Target | milk", "Target | toy"],
    ]);
  });
  it("filters issues", () => {
    expect(groupRows(rows, "issues", "").map((g) => g.rows[0].name)).toEqual([
      "Broken",
    ]);
  });
  it("searches name, envelope, paid with and date", () => {
    expect(groupRows(rows, "all", "aldi")).toHaveLength(1);
    expect(groupRows(rows, "all", "gas")).toHaveLength(3);
    expect(groupRows(rows, "all", "2026-03")).toHaveLength(4);
    expect(groupRows(rows, "all", "credit")).toHaveLength(4);
  });
});

describe("split helpers", () => {
  it("identifies split parents and children", () => {
    expect(isSplitParent(rows[1], rows)).toBe(true);
    expect(isSplitChild(rows[1], rows)).toBe(false);
    expect(isSplitChild(rows[2], rows)).toBe(true);
    expect(isSplitParent(rows[0], rows)).toBe(false);
  });
});

describe("orderEnvelopeOptions", () => {
  it("puts the most used envelopes first, keeping sheet order for ties", () => {
    const used = rows.map((r) => ({
      ...r,
      envelope: r.envelope === "Gas" ? "C" : "",
    }));
    expect(orderEnvelopeOptions(["A", "B", "C"], used)).toEqual([
      "C",
      "A",
      "B",
    ]);
  });
});

describe("groupRows keep predicate", () => {
  it("keeps rows that no longer match the filter while the predicate says so", () => {
    const aldi: ViewRow = { ...makeRow(0, { envelope: "Gas" }), key: "row:0" };
    const shell: ViewRow = { ...makeRow(1), key: "row:1" };
    expect(
      groupRows([aldi, shell], "uncategorized", "").map((g) => g.id),
    ).toEqual(["r:row:1"]);
    expect(
      groupRows([aldi, shell], "uncategorized", "", (r) => r.key === "row:0"),
    ).toHaveLength(2);
  });
});
