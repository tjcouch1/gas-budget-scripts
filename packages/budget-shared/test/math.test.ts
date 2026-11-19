import { describe, expect, it } from "vitest";
import {
  buildSplitParentFormula,
  evaluateMathExpression,
  formatSplitCostFormula,
  isPlainAmount,
  normalizePlainAmount,
  parseAmountInput,
  splitParentFormulaParts,
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
    ["+5", 5],
    ["2*-3", -6],
    ["2/-4", -0.5],
    ["-(-5)", 5],
    ["1-(-2)", 3],
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
    "1.2.3+1",
    "1+1..2",
    "(1+2))",
    "1+(2",
    "4*",
    "-",
    ".",
    "1 .5",
    // Doubled signs (a sign directly after another sign)
    "--5",
    "+-5",
    "-+5",
    "++5",
    "1--2",
    "1+-2",
    "1++2",
    "1-+2",
    "2*--3",
    "2*-+3",
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

describe("isPlainAmount / normalizePlainAmount", () => {
  it("accepts plain numbers with an optional leading minus and one dot", () => {
    for (const ok of ["12", "-3.5", ".25", "7.", "123288.0288", " 4 ", "-.5"])
      expect(isPlainAmount(ok)).toBe(true);
  });

  it("rejects letters, operators, currency, commas, and extra dots", () => {
    for (const bad of [
      "",
      "-",
      ".",
      "1.2.3",
      "$12",
      "1,200",
      "12a",
      "1+2",
      "--1",
      "1-",
      "e5",
    ])
      expect(isPlainAmount(bad)).toBe(false);
  });

  it("keeps the number as typed, dropping only a trailing dot", () => {
    expect(normalizePlainAmount(" 123288.0288 ")).toBe("123288.0288");
    expect(normalizePlainAmount("7.")).toBe("7");
    expect(normalizePlainAmount("1,200")).toBeNull();
  });
});

describe("splitParentFormulaParts / buildSplitParentFormula", () => {
  it("reads the total of R1C1 and A1 split parent formulas", () => {
    expect(splitParentFormulaParts("=45-R[1]C[0]-R[2]C[0]")).toEqual({
      totalExpression: "45",
      totalText: "45",
      total: 45,
      refs: "-R[1]C[0]-R[2]C[0]",
    });
    // getFormulas (A1) and getFormulasR1C1 (`R[1]C`) as real Sheets returns them
    expect(splitParentFormulaParts("=45-C12-C13")).toMatchObject({
      totalText: "45",
      total: 45,
      refs: "-C12-C13",
    });
    expect(splitParentFormulaParts("=45-R[1]C-R[2]C")).toMatchObject({
      totalText: "45",
      refs: "-R[1]C-R[2]C",
    });
    expect(splitParentFormulaParts("= 45 - $C$12 - C13 ")).toMatchObject({
      totalText: "45",
      total: 45,
    });
    expect(splitParentFormulaParts("=-5-R[1]C[0]")).toMatchObject({
      totalText: "-5",
      total: -5,
    });
    expect(splitParentFormulaParts("=45")).toBeNull();
    expect(splitParentFormulaParts("=-R[1]C[0]")).toBeNull();
    expect(splitParentFormulaParts("=45-C12*2")).toBeNull();
    expect(splitParentFormulaParts(null)).toBeNull();
  });

  it("reads formulas the native split appended to", () => {
    expect(splitParentFormulaParts("=10+5-R[1]C[0]")).toMatchObject({
      totalText: "10+5",
      total: 15,
    });
    expect(splitParentFormulaParts("=10-2-C12-C13")).toMatchObject({
      totalText: "10-2",
      total: 8,
      refs: "-C12-C13",
    });
    expect(splitParentFormulaParts("=1.0825*(5)-R[1]C[0]")).toMatchObject({
      totalText: "1.0825*(5)",
      total: 5.4125,
    });
    expect(splitParentFormulaParts("=(1+2)*(3+4)-C12")).toMatchObject({
      totalText: "(1+2)*(3+4)",
      total: 21,
    });
  });

  it("reads an expression total without its outer parentheses", () => {
    expect(splitParentFormulaParts("=(40+5)-R[1]C[0]-R[2]C[0]")).toEqual({
      totalExpression: "(40+5)",
      totalText: "40+5",
      total: 45,
      refs: "-R[1]C[0]-R[2]C[0]",
    });
    expect(splitParentFormulaParts("=(40+5)-C12-C13")?.totalText).toBe("40+5");
  });

  it("gives no total for a formula it cannot read confidently", () => {
    for (const formula of [
      "=SUM(A1:A3)-R[1]C[0]",
      "=45-C12+3-C13",
      "=A1-C12",
      "=1+-2-C12",
      "=(1+2-C12",
    ]) {
      const parts = splitParentFormulaParts(formula);
      expect(parts?.total ?? null).toBeNull();
      expect(parts?.totalText ?? null).toBeNull();
    }
  });

  it("replaces only the total: plain stays plain, math gets parentheses", () => {
    expect(buildSplitParentFormula("=45-R[1]C[0]-R[2]C[0]", "50.25")).toBe(
      "=50.25-R[1]C[0]-R[2]C[0]",
    );
    expect(buildSplitParentFormula("=10+5-R[1]C[0]", "20")).toBe(
      "=20-R[1]C[0]",
    );
    expect(buildSplitParentFormula("=45-R[1]C[0]-R[2]C[0]", "40+5")).toBe(
      "=(40+5)-R[1]C[0]-R[2]C[0]",
    );
    expect(buildSplitParentFormula("=45-R[1]C[0]", "(40+5)")).toBe(
      "=(40+5)-R[1]C[0]",
    );
    expect(buildSplitParentFormula("=45-R[1]C[0]", "=40+5")).toBe(
      "=(40+5)-R[1]C[0]",
    );
    expect(buildSplitParentFormula("=45-C12-C13", "7.")).toBe("=7-C12-C13");
    expect(buildSplitParentFormula("=(40+5)-R[1]C - R[2]C", "45")).toBe(
      "=45-R[1]C-R[2]C",
    );
    for (const bad of ["", "4+", "1..2", "--1", "12a", "SUM(1)"])
      expect(() => buildSplitParentFormula("=45-R[1]C[0]", bad)).toThrow(
        /Invalid split total/,
      );
    expect(() => buildSplitParentFormula("=45", "50")).toThrow(
      /Not a split parent formula/,
    );
  });

  it.each([
    "=45-R[1]C[0]-R[2]C[0]",
    "=45-R[1]C-R[2]C",
    "=45-C12-C13",
    "=10+5-R[1]C[0]",
    "=1.0825*(5)-C12",
    "=(40+5)-R[1]C[0]",
  ])("round-trips totals through %s", (formula) => {
    const refs = splitParentFormulaParts(formula)!.refs.replace(/\s+/g, "");
    for (const [typed, shown] of [
      ["50", "50"],
      ["-3.5", "-3.5"],
      [".5", ".5"],
      ["40+5", "40+5"],
      ["(40+5)", "40+5"],
      ["(1+2)*(3+4)", "(1+2)*(3+4)"],
      ["2*-3", "2*-3"],
      ["10/4", "10/4"],
    ]) {
      const built = buildSplitParentFormula(formula, typed);
      const parts = splitParentFormulaParts(built)!;
      expect(parts.totalText).toBe(shown);
      expect(parts.total).toBeCloseTo(evaluateMathExpression(typed)!, 10);
      expect(parts.refs).toBe(refs);
      // Saving the shown total again writes the same formula
      expect(buildSplitParentFormula(built, parts.totalText!)).toBe(built);
    }
    // The existing total survives an unchanged save
    const existing = splitParentFormulaParts(formula)!;
    const resaved = splitParentFormulaParts(
      buildSplitParentFormula(formula, existing.totalText!),
    )!;
    expect(resaved.total).toBeCloseTo(existing.total!, 10);
    expect(resaved.totalText).toBe(existing.totalText);
  });
});

describe("parseAmountInput", () => {
  it("keeps a plain numeric literal exactly as typed", () => {
    expect(parseAmountInput(" 123288.0288 ")).toEqual({
      ok: true,
      plain: true,
      expression: "123288.0288",
      value: 123288.0288,
    });
    expect(parseAmountInput("-4.50")).toMatchObject({
      plain: true,
      expression: "-4.50",
      value: -4.5,
    });
    expect(parseAmountInput(".25")).toMatchObject({
      plain: true,
      expression: ".25",
    });
    // A trailing "." is dropped (the parser and the sheet do not take "7.")
    expect(parseAmountInput("7.")).toMatchObject({
      plain: true,
      expression: "7",
      value: 7,
    });
  });

  it("detects math and returns it without a leading =", () => {
    expect(parseAmountInput("4.25*2")).toEqual({
      ok: true,
      plain: false,
      expression: "4.25*2",
      value: 8.5,
    });
    expect(parseAmountInput(" =4+1 ")).toMatchObject({
      plain: false,
      expression: "4+1",
      value: 5,
    });
    expect(parseAmountInput("1*1.09")).toMatchObject({
      plain: false,
      expression: "1*1.09",
    });
    expect(parseAmountInput("(40+5)")).toMatchObject({
      plain: false,
      expression: "(40+5)",
      value: 45,
    });
  });

  it("reports blank input as empty", () => {
    for (const blank of ["", "   ", "=", " = "])
      expect(parseAmountInput(blank)).toEqual({ ok: false, reason: "empty" });
  });

  it.each([
    "4+",
    "4*",
    "1..2",
    "1.2.3",
    "1.2.3+1",
    "(1+2",
    "1+2)",
    "--1",
    "1+-2",
    "12a",
    "$5",
    "1,200",
    "A1+1",
    "SUM(1)",
    "2/0",
    "-",
    ".",
    "==1",
    "9".repeat(400),
  ])("rejects %j as invalid", (text) => {
    expect(parseAmountInput(text)).toEqual({ ok: false, reason: "invalid" });
  });

  it("gives input the formula builder accepts, with the same plain/formula choice", () => {
    for (const text of [
      "12",
      "-3.5",
      ".25",
      "7.",
      "4.25*2",
      "=4+1",
      "(1+2)*3",
    ]) {
      const parsed = parseAmountInput(text);
      if (!parsed.ok) throw new Error(text);
      const cell = formatSplitCostFormula(parsed.expression, false, 1.0825);
      expect(typeof cell === "number").toBe(parsed.plain);
      // With tax the value is always a formula
      expect(formatSplitCostFormula(parsed.expression, true, 1.0825)).toBe(
        `=1.0825*(${parsed.expression})`,
      );
    }
  });
});
