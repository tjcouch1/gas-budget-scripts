import { describe, expect, it } from "vitest";
import { addTransactionSheet, addVariablesSheet, getEnv } from "./gas-mocks";

const load = () => import("../src/util/variables.util");

describe("getVariables", () => {
  it("reads key/value pairs from Variables!A1:B30, skipping blank keys", async () => {
    const env = getEnv();
    addVariablesSheet(env);
    const { getVariables, areCached } = await load();
    expect(areCached()).toBeFalsy();
    expect(getVariables()).toEqual({
      Menu: "A1",
      SheetVariables: "H1:I3",
      PayPeriodDays: 14,
      TemplateName: "Template",
      TaxMultiplier: 1.0825,
      ForwardEmailAddress: "fwd@example.com",
      ForwardName: "Sam",
    });
    expect(areCached()).toBeTruthy();
  });

  it("caches the result for later calls", async () => {
    const env = getEnv();
    const sheet = addVariablesSheet(env);
    const { Variables } = await load();
    expect(Variables.getVariables().PayPeriodDays).toBe(14);
    sheet.getRange("B3").setValue(7);
    expect(Variables.getVariables().PayPeriodDays).toBe(14);
  });

  it("throws if a required variable is missing", async () => {
    const env = getEnv();
    addVariablesSheet(env, [["Menu", "A1"]]);
    const { getVariables } = await load();
    expect(() => getVariables()).toThrow(
      "Some variables were not found in the Variables sheet in the range Variables!A1:B30!",
    );
  });

  it("ignores variables beyond row 30", async () => {
    const env = getEnv();
    const vars = [
      ...[
        ["Menu", "A1"],
        ["SheetVariables", "H1:I3"],
        ["PayPeriodDays", 14],
        ["TemplateName", "Template"],
        ["TaxMultiplier", 1.0825],
        ["ForwardEmailAddress", "fwd@example.com"],
      ],
      ...Array.from({ length: 24 }, (_, i) => [`x${i}`, i]),
      ["ForwardName", "Sam"], // row 31: outside the range
    ];
    addVariablesSheet(env, vars);
    const { getVariables } = await load();
    expect(() => getVariables()).toThrow(/Some variables were not found/);
  });
});

describe("optional ReceiptRules variable", () => {
  it("is read when present (row 11 is inside the range)", async () => {
    const env = getEnv();
    const sheet = addVariablesSheet(env);
    sheet.seed("A11", [["ReceiptRules", '{"version":1}']]);
    const { getVariables } = await load();
    expect(getVariables().ReceiptRules).toBe('{"version":1}');
  });

  it("is absent when missing", async () => {
    addVariablesSheet(getEnv());
    const { getVariables } = await load();
    expect(getVariables()).not.toHaveProperty("ReceiptRules");
  });

  it("exports the range constants and the note for the key cell", async () => {
    const {
      VARIABLES_RANGE,
      VARIABLES_SHEET_NAME,
      OPTIONAL_VARIABLE_NAMES,
      GLOBAL_VARIABLE_NOTES,
    } = await load();
    const { RECEIPT_RULES_NOTE } = await import("@budget/shared");
    expect(VARIABLES_SHEET_NAME).toBe("Variables");
    expect(VARIABLES_RANGE).toBe("Variables!A1:B30");
    expect(OPTIONAL_VARIABLE_NAMES).toEqual(["ReceiptRules"]);
    expect(GLOBAL_VARIABLE_NOTES.ReceiptRules).toBe(RECEIPT_RULES_NOTE);
  });
});

describe("getSheetVariables", () => {
  it("reads per-sheet variables using the SheetVariables range", async () => {
    const env = getEnv();
    addVariablesSheet(env);
    const sheet = addTransactionSheet(env, "1/1/24 - 1/14/24", {
      start: "B7",
      max: 25,
      checkboxes: "J7",
    });
    const { getSheetVariables } = await load();
    expect(getSheetVariables(sheet)).toEqual({
      TransactionsStart: "B7",
      TransactionsMax: 25,
      SplitCheckboxesStart: "J7",
    });
  });

  it("throws for the Variables sheet", async () => {
    const env = getEnv();
    const vs = addVariablesSheet(env);
    const { getSheetVariables } = await load();
    expect(() => getSheetVariables(vs)).toThrow(
      "Cannot get sheet variables for Variables sheet.",
    );
  });

  it("throws when a sheet variable is missing", async () => {
    const env = getEnv();
    addVariablesSheet(env);
    const sheet = env.spreadsheet.addSheet("Bad");
    sheet.seed("H1", [["TransactionsStart", "A5"]]);
    const { getSheetVariables } = await load();
    expect(() => getSheetVariables(sheet)).toThrow(
      "Some sheet variables were not found in the Bad sheet in the range H1:I3!",
    );
  });

  it("caches per sheet name", async () => {
    const env = getEnv();
    addVariablesSheet(env);
    const sheet = addTransactionSheet(env, "S", { max: 10 });
    const { getSheetVariables } = await load();
    expect(getSheetVariables(sheet).TransactionsMax).toBe(10);
    sheet.getRange("I2").setValue(99);
    expect(getSheetVariables(sheet).TransactionsMax).toBe(10);
  });
});
