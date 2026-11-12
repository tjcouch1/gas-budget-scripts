import { describe, expect, it } from "vitest";
import { areDatesEqual, isDate, isString, Util } from "../src/util/util";
import { doRangesIntersect } from "../src/util/spreadsheet-util.util";
import { getEnv } from "./gas-mocks";

describe("util", () => {
  it("isString accepts primitives and String objects only", () => {
    expect(isString("a")).toBe(true);
    expect(isString(new String("a"))).toBe(true);
    expect(isString(1)).toBe(false);
    expect(isString(undefined)).toBe(false);
  });

  it("isDate detects Date instances", () => {
    expect(isDate(new Date())).toBe(true);
    expect(isDate("2024-01-01")).toBe(false);
  });

  it("areDatesEqual compares times for dates and strings otherwise", () => {
    expect(
      areDatesEqual(new Date("2024-01-01T00:00:00Z"), new Date("2024-01-01T00:00:00Z"))
    ).toBe(true);
    expect(
      areDatesEqual(new Date("2024-01-01T00:00:00Z"), new Date("2024-01-02T00:00:00Z"))
    ).toBe(false);
    expect(areDatesEqual("", "")).toBe(true);
    expect(areDatesEqual(undefined, undefined)).toBe(true);
    expect(areDatesEqual("", undefined)).toBe(false);
    // Date vs. its toString() is "equal" via the string fallback
    const d = new Date("2024-01-01T00:00:00Z");
    expect(areDatesEqual(d, d.toString())).toBe(true);
  });

  it("exposes the helpers on the Util object", () => {
    expect(Util).toEqual({ isString, isDate, areDatesEqual });
  });
});

describe("doRangesIntersect", () => {
  it("detects overlap on the same sheet and rejects separate/other-sheet ranges", () => {
    const ss = getEnv().spreadsheet;
    const s1 = ss.addSheet("One");
    const s2 = ss.addSheet("Two");
    const a = s1.getRange("A1:C3");
    expect(doRangesIntersect(a, s1.getRange("C3:D4"))).toBe(true);
    expect(doRangesIntersect(a, s1.getRange("B2"))).toBe(true);
    expect(doRangesIntersect(a, s1.getRange("A4:C5"))).toBe(false);
    expect(doRangesIntersect(a, s1.getRange("D1:E3"))).toBe(false);
    expect(doRangesIntersect(a, s2.getRange("A1:C3"))).toBe(false);
  });
});
