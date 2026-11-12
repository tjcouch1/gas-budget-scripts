import { describe, expect, it } from "vitest";
import {
  summarizeSplit,
  toSplitInputs,
  type SplitDraftItem,
} from "./split-math.ts";

const item = (over: Partial<SplitDraftItem>): SplitDraftItem => ({
  key: "k",
  itemNote: "",
  expression: "",
  preTax: false,
  envelope: "",
  ...over,
});

describe("summarizeSplit", () => {
  it("computes remaining with inline math", () => {
    const s = summarizeSplit(
      50,
      [item({ expression: "4.23*2+12.09" })],
      1.0825,
    );
    expect(s.items[0].value).toBeCloseTo(20.55);
    expect(s.remaining).toBe(29.45);
    expect(s.canSubmit).toBe(true);
  });

  it("applies the tax multiplier only to pre-tax items", () => {
    const s = summarizeSplit(
      20,
      [item({ expression: "10", preTax: true }), item({ expression: "5" })],
      1.1,
    );
    expect(s.items.map((i) => i.value)).toEqual([11, 5]);
    expect(s.remaining).toBe(4);
  });

  it("blocks invalid expressions", () => {
    expect(
      summarizeSplit(20, [item({ expression: "1+" })], 1.1).canSubmit,
    ).toBe(false);
    expect(summarizeSplit(20, [item({ expression: "" })], 1.1).canSubmit).toBe(
      false,
    );
    expect(
      summarizeSplit(20, [item({ expression: "SUM(1)" })], 1.1).canSubmit,
    ).toBe(false);
  });

  it("blocks over-allocation but allows exactly zero remaining", () => {
    expect(
      summarizeSplit(10, [item({ expression: "10.01" })], 1).canSubmit,
    ).toBe(false);
    const exact = summarizeSplit(
      10,
      [item({ expression: "4" }), item({ expression: "6" })],
      1,
    );
    expect(exact.remaining).toBe(0);
    expect(exact.canSubmit).toBe(true);
  });

  it("rejects notes that would be formulas", () => {
    const s = summarizeSplit(
      10,
      [item({ expression: "1", itemNote: "=HYPERLINK()" })],
      1,
    );
    expect(s.canSubmit).toBe(false);
  });

  it("builds the request items", () => {
    expect(
      toSplitInputs([
        item({
          itemNote: " milk ",
          expression: " 1+2 ",
          preTax: true,
          envelope: "Groceries",
        }),
      ]),
    ).toEqual([
      {
        itemNote: "milk",
        expression: "1+2",
        preTax: true,
        envelope: "Groceries",
      },
    ]);
  });
});
