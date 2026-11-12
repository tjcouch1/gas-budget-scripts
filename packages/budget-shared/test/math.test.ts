import { describe, expect, it } from "vitest";
import {
  evaluateMathExpression,
  formatSplitCostFormula,
} from "../src/index.ts";

describe("evaluateMathExpression", () => {
  it.each([
    ["5", 5],
    ["  7  ", 7],
    [".5", 0.5],
    ["1+3.23*2", 7.46],
    ["4.23*2+12.09+3.44*6", 4.23 * 2 + 12.09 + 3.44 * 6],
    ["2+3*4", 14],
    ["(2+3)*4", 20],
    ["10/4", 2.5],
    ["10-2-3", 5],
    ["100/10/5", 2],
    ["-5", -5],
    ["--5", 5],
    ["+5", 5],
    ["2*-3", -6],
    ["-(1+2)", -3],
    ["((1))", 1],
    [" 1 + 2 * ( 3 - 1 ) ", 5],
    ["0", 0],
    ["1.50", 1.5],
  ])("evaluates %j", (expr, expected) => {
    expect(evaluateMathExpression(expr)).toBeCloseTo(expected, 10);
  });

  it.each([
    "",
    "   ",
    "1+",
    "+",
    "*2",
    "2*",
    "2/0",
    "1/(1-1)",
    "(1+2",
    "1+2)",
    "()",
    "1 2",
    "1..2",
    "1.",
    "1.5.5",
    "1e5",
    "0x10",
    "SUM(1)",
    '=IMPORTXML("x","y")',
    "A1+1",
    "1;2",
    "1,5",
    "alert(1)",
    "1+2 // c",
    "5%",
    "$5",
    "NaN",
    "Infinity",
  ])("rejects %j", (expr) => {
    expect(evaluateMathExpression(expr)).toBeNull();
  });

  it("returns null for non-string input", () => {
    expect(evaluateMathExpression(undefined as unknown as string)).toBeNull();
    expect(evaluateMathExpression(5 as unknown as string)).toBeNull();
  });

  it("returns null on a non-finite result", () => {
    expect(evaluateMathExpression("9".repeat(400))).toBeNull();
  });

  it("does not run the expression as code", () => {
    const g = globalThis as Record<string, unknown>;
    g.__pwned = false;
    expect(evaluateMathExpression("(globalThis.__pwned=true)")).toBeNull();
    expect(g.__pwned).toBe(false);
  });
});

describe("formatSplitCostFormula", () => {
  it("wraps with the tax multiplier when pre-tax is on", () => {
    expect(formatSplitCostFormula("4.23*2+12.09+3.44*6", true, 1.0825)).toBe(
      "=1.0825*(4.23*2+12.09+3.44*6)",
    );
  });

  it("matches the native split placeholder style", () => {
    expect(formatSplitCostFormula("0", true, 1.0825)).toBe("=1.0825*(0)");
  });

  it("wraps a single number when pre-tax is on", () => {
    expect(formatSplitCostFormula("5", true, 1.1)).toBe("=1.1*(5)");
  });

  it("writes a formula without outer parentheses when pre-tax is off", () => {
    expect(formatSplitCostFormula("1+3.23*2", false, 1.0825)).toBe("=1+3.23*2");
  });

  it("writes a raw number for a single numeric literal", () => {
    expect(formatSplitCostFormula("5", false, 1.0825)).toBe(5);
    expect(formatSplitCostFormula("5.25", false, 1.0825)).toBe(5.25);
    expect(formatSplitCostFormula(".5", false, 1.0825)).toBe(0.5);
    expect(formatSplitCostFormula("-5", false, 1.0825)).toBe(-5);
  });

  it("keeps user-typed parentheses", () => {
    expect(formatSplitCostFormula("(1+2)", false, 1.0825)).toBe("=(1+2)");
  });

  it("trims whitespace and strips a leading =", () => {
    expect(formatSplitCostFormula("  =1+2 ", false, 1.0825)).toBe("=1+2");
    expect(formatSplitCostFormula("= 3", false, 1.0825)).toBe(3);
    expect(formatSplitCostFormula(" 1 + 2 ", true, 1.5)).toBe("=1.5*(1 + 2)");
  });

  it.each([
    '=IMPORTXML("http://x","//a")',
    "1+",
    "2/0",
    "SUM(1)",
    "",
    "==1",
    "1+HYPERLINK(1)",
    "A1",
  ])("throws for %j", (expr) => {
    expect(() => formatSplitCostFormula(expr, false, 1.0825)).toThrow(
      /Invalid cost expression/,
    );
    expect(() => formatSplitCostFormula(expr, true, 1.0825)).toThrow(
      /Invalid cost expression/,
    );
  });
});
