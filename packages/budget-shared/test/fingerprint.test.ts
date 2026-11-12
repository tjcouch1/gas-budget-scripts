import { describe, expect, it } from "vitest";
import {
  computeRowFingerprint,
  fnv1a32,
  formatLocalDate,
} from "../src/index.ts";
import type { FingerprintRow } from "../src/index.ts";

const base: FingerprintRow = {
  date: "2024-03-05",
  name: "Store",
  cost: 12.5,
  envelope: "Food",
  paidWith: "Credit",
  taxDeductible: "",
};

describe("fnv1a32", () => {
  it("matches known FNV-1a vectors", () => {
    expect(fnv1a32("")).toBe("811c9dc5");
    expect(fnv1a32("a")).toBe("e40c292c");
    expect(fnv1a32("foobar")).toBe("bf9cf968");
  });
});

describe("computeRowFingerprint", () => {
  it("is an 8-digit lowercase hex string", () => {
    expect(computeRowFingerprint(base)).toMatch(/^[0-9a-f]{8}$/);
  });

  it("is stable across property order", () => {
    const reordered: FingerprintRow = {
      taxDeductible: "",
      paidWith: "Credit",
      envelope: "Food",
      cost: 12.5,
      name: "Store",
      date: "2024-03-05",
    };
    expect(computeRowFingerprint(reordered)).toBe(computeRowFingerprint(base));
  });

  it("treats a local Date and its yyyy-MM-dd string as equal", () => {
    const d = new Date(2024, 2, 5, 13, 45);
    expect(formatLocalDate(d)).toBe("2024-03-05");
    expect(computeRowFingerprint({ ...base, date: d })).toBe(
      computeRowFingerprint(base),
    );
  });

  it("treats null, undefined, and empty as the same blank", () => {
    const blank = computeRowFingerprint({
      date: "",
      name: "",
      cost: "",
      envelope: "",
      paidWith: "",
      taxDeductible: "",
    });
    expect(
      computeRowFingerprint({
        date: null,
        name: undefined,
        cost: null,
        envelope: null,
        paidWith: undefined,
        taxDeductible: null,
      }),
    ).toBe(blank);
    expect(computeRowFingerprint({ ...base, date: new Date(NaN) })).toBe(
      computeRowFingerprint({ ...base, date: "" }),
    );
  });

  it("trims only the right side of strings", () => {
    expect(computeRowFingerprint({ ...base, name: "Store  " })).toBe(
      computeRowFingerprint(base),
    );
    expect(computeRowFingerprint({ ...base, name: "  Store" })).not.toBe(
      computeRowFingerprint(base),
    );
  });

  it("uses the cost formula instead of the value when present", () => {
    const withFormula = computeRowFingerprint({
      ...base,
      cost: 7.46,
      costFormula: "=1+3.23*2",
    });
    expect(withFormula).not.toBe(
      computeRowFingerprint({ ...base, cost: 7.46 }),
    );
    expect(
      computeRowFingerprint({
        ...base,
        cost: 999,
        costFormula: "=1+3.23*2",
      }),
    ).toBe(withFormula);
    expect(
      computeRowFingerprint({ ...base, cost: 7.46, costFormula: null }),
    ).toBe(computeRowFingerprint({ ...base, cost: 7.46 }));
  });

  it.each<[string, Partial<FingerprintRow>]>([
    ["date", { date: "2024-03-06" }],
    ["name", { name: "Other" }],
    ["cost", { cost: 12.51 }],
    ["envelope", { envelope: "Gas" }],
    ["paidWith", { paidWith: "Cash" }],
    ["taxDeductible", { taxDeductible: "Yes" }],
  ])("changes when %s changes", (_field, change) => {
    expect(computeRowFingerprint({ ...base, ...change })).not.toBe(
      computeRowFingerprint(base),
    );
  });

  it("does not confuse values moved between fields", () => {
    expect(
      computeRowFingerprint({ ...base, name: "a", envelope: "b" }),
    ).not.toBe(computeRowFingerprint({ ...base, name: "b", envelope: "a" }));
  });

  it("distinguishes a numeric cost from the same text", () => {
    expect(computeRowFingerprint({ ...base, cost: 5 })).not.toBe(
      computeRowFingerprint({ ...base, cost: "5" }),
    );
  });

  it("is pinned for a known row (client and server must agree)", () => {
    expect(computeRowFingerprint(base)).toBe(
      fnv1a32(
        JSON.stringify(["2024-03-05", "Store", 12.5, "Food", "Credit", ""]),
      ),
    );
  });
});
