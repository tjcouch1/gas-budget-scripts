import { readFileSync } from "node:fs";
import { describe, expect, it } from "vitest";
import { DEFAULT_RECEIPT_RULES, RECEIPT_RULES_NOTE } from "@budget/shared";
import { addVariablesSheet, DEFAULT_VARIABLES, getEnv } from "./gas-mocks";

const load = () => import("../src/index");
const loadUtil = () => import("../src/util/global-variables-setup.util");

const DEFAULT_JSON = JSON.stringify(DEFAULT_RECEIPT_RULES, null, 2);

describe("setupAddOptionalGlobalVariables", () => {
  it("adds ReceiptRules with the defaults JSON and a key-cell note in the first empty row", async () => {
    const env = getEnv();
    const variables = addVariablesSheet(env);
    variables.getRange("A1").setNote("existing note");
    const { setupAddOptionalGlobalVariables } = await load();
    setupAddOptionalGlobalVariables();

    expect(variables.getRange("A8:B8").getValues()).toEqual([
      ["ReceiptRules", DEFAULT_JSON],
    ]);
    expect(JSON.parse(variables.getRange("B8").getValue())).toEqual(
      DEFAULT_RECEIPT_RULES,
    );
    const { GLOBAL_VARIABLE_NOTES } =
      await import("../src/util/variables.util");
    expect(variables.getRange("A8").getNotes()).toEqual([
      [GLOBAL_VARIABLE_NOTES.ReceiptRules],
    ]);
    expect(GLOBAL_VARIABLE_NOTES.ReceiptRules).toBe(RECEIPT_RULES_NOTE);
    expect(variables.getRange("A1").getNotes()).toEqual([["existing note"]]);
    expect(env.logs.join("\n")).toContain("Added ReceiptRules to Variables!A8");

    // getVariables sees it in the same execution (cache busted)
    const { getVariables } = await import("../src/util/variables.util");
    expect(getVariables().ReceiptRules).toBe(DEFAULT_JSON);
  });

  it("is idempotent: a second run logs Nothing to do and keeps the value and note", async () => {
    const env = getEnv();
    const variables = addVariablesSheet(env, [
      ...DEFAULT_VARIABLES,
      ["ReceiptRules", ""],
    ]);
    variables.getRange("A8").setNote("my note");
    const { setupAddOptionalGlobalVariables } = await load();
    setupAddOptionalGlobalVariables();

    expect(variables.getRange("A8:B9").getValues()).toEqual([
      ["ReceiptRules", ""],
      ["", ""],
    ]);
    expect(variables.getRange("A8").getNotes()).toEqual([["my note"]]);
    expect(env.logs.join("\n")).toContain("Nothing to do");
  });

  it("skips rows below the last key that have stray content in either column", async () => {
    const variables = addVariablesSheet(getEnv());
    variables.seed("B8", [["stray"]]);
    variables.seed("A10", [["", "x"]]);
    const { setupAddOptionalGlobalVariables } = await load();
    setupAddOptionalGlobalVariables();

    expect(variables.getRange("A8:B11").getValues()).toEqual([
      ["", "stray"],
      ["ReceiptRules", DEFAULT_JSON],
      ["", "x"],
      ["", ""],
    ]);
  });

  it("throws and changes nothing when A1:B30 has no empty row left", async () => {
    const vars = [
      ...DEFAULT_VARIABLES,
      ...Array.from({ length: 23 }, (_, i) => [`x${i}`, i]),
    ];
    const variables = addVariablesSheet(getEnv(), vars);
    const { setupAddOptionalGlobalVariables } = await load();
    expect(() => setupAddOptionalGlobalVariables()).toThrow(
      "No empty row left in Variables!A1:B30 for ReceiptRules",
    );
    expect(variables.getRange("A1:B30").getValues()).toEqual(vars);
    expect(variables.getRange("A1:A30").getNotes().flat().join("")).toBe("");
  });

  it("is exported from src/index.ts so the build emits a top-level function", () => {
    const source = readFileSync(
      new URL("../src/index.ts", import.meta.url),
      "utf8",
    );
    expect(source).toMatch(
      /export \{ setupAddOptionalGlobalVariables \} from "\.\/util\/global-variables-setup\.util";/,
    );
  });
});

describe("ensureGlobalVariableRow", () => {
  it("returns the existing value cell without writing", async () => {
    const variables = addVariablesSheet(getEnv(), [
      ...DEFAULT_VARIABLES,
      ["ReceiptRules", "{}"],
    ]);
    const { ensureGlobalVariableRow } = await loadUtil();
    const { valueRange, added } = ensureGlobalVariableRow("ReceiptRules");
    expect(added).toBe(false);
    expect(valueRange.getValue()).toBe("{}");
    expect(variables.getRange("A8").getNotes()).toEqual([[""]]);
  });

  it("adds the key with its note and an empty value when missing", async () => {
    const variables = addVariablesSheet(getEnv());
    const { ensureGlobalVariableRow } = await loadUtil();
    const { valueRange, added } = ensureGlobalVariableRow("ReceiptRules");
    expect(added).toBe(true);
    valueRange.setValues([["{ }"]]);
    expect(variables.getRange("A8:B8").getValues()).toEqual([
      ["ReceiptRules", "{ }"],
    ]);
    expect(variables.getRange("A8").getNotes()).toEqual([[RECEIPT_RULES_NOTE]]);
  });
});
