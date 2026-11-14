import { describe, expect, it } from "vitest";
import { makeRow } from "../test/fixtures.ts";
import {
  evaluateRowDraft,
  initialRowDraft,
  parseChildCost,
  type RowDraftOptions,
} from "./row-draft.ts";

describe("parseChildCost", () => {
  it("unwraps the tax multiplier formula into the expression with tax", () => {
    expect(
      parseChildCost(
        { cost: 22.25, costFormula: "=1.0825*(4.23*2+12.09)" },
        1.0825,
      ),
    ).toEqual({ expression: "4.23*2+12.09", preTax: true });
    expect(
      parseChildCost({ cost: 0, costFormula: "=1.0825*(0)" }, 1.0825),
    ).toEqual({ expression: "0", preTax: true });
  });

  it("shows other formulas without = and numbers as typed, without tax", () => {
    expect(
      parseChildCost({ cost: 7.46, costFormula: "=1+3.23*2" }, 1.0825),
    ).toEqual({ expression: "1+3.23*2", preTax: false });
    // A different multiplier is not this sheet's tax
    expect(
      parseChildCost({ cost: 2.2, costFormula: "=1.1*(2)" }, 1.0825),
    ).toEqual({ expression: "1.1*(2)", preTax: false });
    expect(parseChildCost({ cost: 2, costFormula: null }, 1.0825)).toEqual({
      expression: "2",
      preTax: false,
    });
    expect(parseChildCost({ cost: null, costFormula: null }, 1.0825)).toEqual({
      expression: "",
      preTax: false,
    });
  });
});

describe("row drafts", () => {
  const child: RowDraftOptions = {
    nameMode: "itemNote",
    costMode: "child",
    taxMultiplier: 1.1,
  };
  const row = makeRow(3, {
    name: "Aldi | milk",
    cost: 11,
    costFormula: "=1.1*(10)",
    envelope: "Groceries",
    note: "old",
  });

  it("starts from the row and reports no changes", () => {
    const initial = initialRowDraft(row, child);
    expect(initial).toEqual({
      name: "milk",
      cost: "10",
      preTax: true,
      envelope: "Groceries",
      paidWith: "Credit",
      taxDeductible: false,
      note: "old",
    });
    const result = evaluateRowDraft(row, initial, initial, child);
    expect(result.changes).toEqual({});
    expect(result.value).toBe(11);
  });

  it("collects only changed fields, keeping the merchant prefix", () => {
    const initial = initialRowDraft(row, child);
    const result = evaluateRowDraft(
      row,
      initial,
      {
        ...initial,
        name: "oat milk",
        cost: "2*3",
        preTax: false,
        paidWith: "Debit",
        taxDeductible: true,
        note: "",
      },
      child,
    );
    expect(result.changes).toEqual({
      name: "Aldi | oat milk",
      costExpression: "2*3",
      paidWith: "Debit",
      taxDeductible: "Yes",
      note: "",
    });
    expect(result.value).toBe(6);
  });

  it("re-sends the expression when only Add tax changes", () => {
    const initial = initialRowDraft(row, child);
    const result = evaluateRowDraft(
      row,
      initial,
      { ...initial, preTax: false },
      child,
    );
    expect(result.changes).toEqual({ costExpression: "10" });
    expect(result.value).toBe(10);
  });

  it("flags invalid costs and formula-like names", () => {
    const plain: RowDraftOptions = {
      nameMode: "name",
      costMode: "plain",
      taxMultiplier: 1.1,
    };
    const r = makeRow(0, { name: "Aldi", cost: 5 });
    const initial = initialRowDraft(r, plain);
    const result = evaluateRowDraft(
      r,
      initial,
      { ...initial, name: "=x", cost: "2/0" },
      plain,
    );
    expect(result.costInvalid).toBe(true);
    expect(result.nameUnsafe).toBe(true);
    expect(result.value).toBeNull();
    // A plain cost never gets the tax multiplier; clearing the box keeps the cost
    const cleared = evaluateRowDraft(
      r,
      initial,
      { ...initial, cost: " " },
      plain,
    );
    expect(cleared.changes).toEqual({});
    expect(cleared.value).toBe(5);
  });

  it("ignores the name and cost when the card does not edit them", () => {
    const none: RowDraftOptions = {
      nameMode: "none",
      costMode: "none",
      taxMultiplier: 1.1,
    };
    const parent = makeRow(0, { name: "=Odd", costFormula: "=10-R[1]C[0]" });
    const initial = initialRowDraft(parent, none);
    const result = evaluateRowDraft(
      parent,
      initial,
      { ...initial, cost: "3", envelope: "Gas" },
      none,
    );
    expect(result.changes).toEqual({ envelope: "Gas" });
    expect(result.nameUnsafe).toBe(false);
  });
});
