import { describe, expect, it } from "vitest";
import {
  addTransactionSheet,
  addVariablesSheet,
  DEFAULT_VARIABLES,
  getEnv,
  type AnySheet,
} from "./gas-mocks";

const load = () => import("../src/index");

const OPTIONAL_KEYS = [
  "EnvelopesStart",
  "EnvelopesMax",
  "PaidWithStart",
  "PaidWithMax",
];

function variablesWithSheetVariables(a1: string) {
  return DEFAULT_VARIABLES.map(([key, value]) =>
    key === "SheetVariables" ? [key, a1] : [key, value],
  );
}

function setup(sheetVariablesA1 = "H1:I3") {
  const env = getEnv();
  const variables = addVariablesSheet(
    env,
    variablesWithSheetVariables(sheetVariablesA1),
  );
  const template: AnySheet = addTransactionSheet(env, "Template");
  const tab: AnySheet = addTransactionSheet(env, "1/1/2026");
  return { env, variables, template, tab };
}

describe("setupAddOptionalSheetVariables", () => {
  it("writes the missing keys into empty rows of the block when there is room", async () => {
    const { env, variables, template, tab } = setup("H1:I8");
    // An empty row in the middle and an already-present key
    template.seed("H5", [["EnvelopesMax", 30]]);
    const { setupAddOptionalSheetVariables } = await load();
    setupAddOptionalSheetVariables();

    expect(template.getRange("H1:I8").getValues()).toEqual([
      ["TransactionsStart", "A5"],
      ["TransactionsMax", 10],
      ["SplitCheckboxesStart", "G5"],
      ["EnvelopesStart", ""],
      ["EnvelopesMax", 30],
      ["PaidWithStart", ""],
      ["PaidWithMax", ""],
      ["", ""],
    ]);
    expect(variables.getRange("B2").getValue()).toBe("H1:I8");
    // Other tabs are not touched
    expect(
      tab
        .getRange("H4:I8")
        .getValues()
        .flat()
        .every((v) => v === ""),
    ).toBe(true);
    const log = String(env.logs.at(-1));
    expect(log).toContain("Template!I4: EnvelopesStart = ");
    expect(log).toContain("Template!I6: PaidWithStart = ");
    expect(log).toContain("Template!I7: PaidWithMax = ");
    expect(log).not.toContain("I5:");
    expect(log).toContain("addTransactionSheet");
  });

  it("puts the TSDoc-sourced note on added key cells only", async () => {
    const { template } = setup("H1:I8");
    template.seed("H5", [["EnvelopesMax", 30]]);
    template.getRange("H1").setNote("existing note");
    template.getRange("H5").setNote("kept note");
    const { setupAddOptionalSheetVariables } = await load();
    setupAddOptionalSheetVariables();
    const { OPTIONAL_SHEET_VARIABLE_NOTES: N } =
      await import("../src/util/variables.util");
    expect(template.getRange("H1:H8").getNotes()).toEqual([
      ["existing note"],
      [""],
      [""],
      [N.EnvelopesStart],
      ["kept note"],
      [N.PaidWithStart],
      [N.PaidWithMax],
      [""],
    ]);
    // Value cells get no notes
    expect(
      template
        .getRange("I1:I8")
        .getNotes()
        .every(([n]) => n === ""),
    ).toBe(true);
  });

  it("keeps formulas in the block", async () => {
    const { template } = setup("H1:I7");
    template.getRange("I2").setValues([["=5*2"]]);
    const { setupAddOptionalSheetVariables } = await load();
    setupAddOptionalSheetVariables();
    expect(template.getRange("I2").getFormulas()).toEqual([["=5*2"]]);
  });

  it("expands the SheetVariables range when the block has no room", async () => {
    const { env, variables, template, tab } = setup();
    const { setupAddOptionalSheetVariables } = await load();
    setupAddOptionalSheetVariables();

    expect(variables.getRange("B2").getValue()).toBe("H1:I7");
    expect(template.getRange("H4:I7").getValues()).toEqual(
      OPTIONAL_KEYS.map((key) => [key, ""]),
    );
    expect(
      tab
        .getRange("H4:I7")
        .getValues()
        .flat()
        .every((v) => v === ""),
    ).toBe(true);
    expect(String(env.logs.at(-1))).toContain(
      "expanded SheetVariables in Variables to H1:I7",
    );

    // Variables memoization was reset: the new range and empty optional values work
    const { Variables } = await import("../src/util/variables.util");
    expect(Variables.getVariables().SheetVariables).toBe("H1:I7");
    expect(Variables.getSheetVariables(template)).toMatchObject({
      EnvelopesStart: "",
      PaidWithMax: "",
    });
    // The existing tab without the keys still reads its required variables
    expect(Variables.getSheetVariables(tab).TransactionsMax).toBe(10);
  });

  it("expands only by the rows still needed", async () => {
    const { variables, template } = setup("H1:I4");
    const { setupAddOptionalSheetVariables } = await load();
    setupAddOptionalSheetVariables();
    expect(variables.getRange("B2").getValue()).toBe("H1:I7");
    expect(template.getRange("H4:H7").getValues().flat()).toEqual(
      OPTIONAL_KEYS,
    );
  });

  it("refuses to expand when the new rows are not empty on another tab and changes nothing", async () => {
    const { variables, template, tab } = setup();
    tab.seed("I5", [["something"]]);
    tab.seed("H7", [["x"]]);
    const { setupAddOptionalSheetVariables } = await load();
    expect(() => setupAddOptionalSheetVariables()).toThrow(
      "these cells must be empty first: 1/1/2026!I5, 1/1/2026!H7. Nothing was changed.",
    );
    expect(variables.getRange("B2").getValue()).toBe("H1:I3");
    expect(
      template
        .getRange("H4:I7")
        .getValues()
        .flat()
        .every((v) => v === ""),
    ).toBe(true);
  });

  it("does nothing when the Template already has every optional key", async () => {
    const { env, variables, template } = setup("H1:I7");
    template.seed(
      "H4",
      OPTIONAL_KEYS.map((key) => [key, "x"]),
    );
    const before = template.getRange("H1:I7").getValues();
    const { setupAddOptionalSheetVariables } = await load();
    setupAddOptionalSheetVariables();
    expect(template.getRange("H1:I7").getValues()).toEqual(before);
    expect(variables.getRange("B2").getValue()).toBe("H1:I7");
    expect(String(env.logs.at(-1))).toContain("Nothing to do");
  });
});
