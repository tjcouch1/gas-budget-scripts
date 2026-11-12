import { beforeEach, describe, expect, it, vi } from "vitest";
import {
  addTransactionSheet,
  addVariablesSheet,
  getEnv,
  AnySheet,
} from "./gas-mocks";

const load = () => import("../src/index");

let main: AnySheet;

beforeEach(() => {
  const env = getEnv();
  // First sheet is the active sheet and hosts the in-sheet menu at A1
  main = env.spreadsheet.addSheet("Main");
  addVariablesSheet(env);
});

describe("onOpen", () => {
  it("builds the menu bar menu", async () => {
    const { onOpen } = await load();
    onOpen();
    const [menu, ...rest] = getEnv().SpreadsheetApp.ui.menus;
    expect(rest).toEqual([]);
    expect(menu.name).toBe("Budgeting");
    expect(menu.added).toBe(true);
    expect(menu.items).toEqual([
      { name: "Import 50 receipts (no mark)", functionName: "getAndRecordSomeReceiptsNoMark" },
      { name: "Import all receipts (no mark)", functionName: "getAndRecordReceiptsNoMark" },
      "separator",
      { name: "Import 25 receipts and mark", functionName: "getAndRecordSomeReceiptsAndMark" },
      { name: "Import all receipts and mark", functionName: "getAndRecordReceiptsAndMark" },
      "separator",
      { name: "Add one transaction sheet if needed", functionName: "addTransactionSheet" },
      { name: "Add transaction sheets until up-to-date", functionName: "addTransactionSheets" },
    ]);
  });

  it("fills in the in-sheet menu with a bold header, checkboxes and names", async () => {
    const { onOpen } = await load();
    onOpen();
    expect(main.getRange("A1").getValue()).toBe("Budgeting");
    expect(main.textStyles).toHaveLength(1);
    expect(main.textStyles[0].style).toEqual({ bold: true });
    expect(main.getRange("A2:A9").getValues()).toEqual(Array(8).fill([false]));
    expect(main.getRange("B2:B9").getValues().map(([v]) => v)).toEqual([
      "Import 50 receipts (no mark)",
      "Import all receipts (no mark)",
      "-------",
      "Import 25 receipts and mark",
      "Import all receipts and mark",
      "-------",
      "Add one transaction sheet if needed",
      "Add transaction sheets until up-to-date",
    ]);
  });

  it("does not overwrite an already-filled menu", async () => {
    main.seed("A1", [["Existing"]]);
    const { onOpen } = await load();
    onOpen();
    expect(main.getRange("A1").getValue()).toBe("Existing");
    expect(main.getRange("B2").getValue()).toBe("");
  });

  it("still builds the in-sheet menu when the UI is unavailable (mobile)", async () => {
    const env = getEnv();
    env.SpreadsheetApp.uiAvailable = false;
    const { onOpen } = await load();
    onOpen();
    expect(main.getRange("A1").getValue()).toBe("Budgeting");
    expect(env.logs.some((l) => String(l).includes("User may be on mobile"))).toBe(true);
  });
});

describe("onEditInstalled - menu checkboxes", () => {
  /** Build an edit event for the cell at A1 notation on a sheet */
  const edit = (sheet: AnySheet, a1: string, checked = true) => {
    const range = sheet.getRange(a1);
    range.setValue(checked);
    return { range } as unknown as GoogleAppsScript.Events.SheetsOnEdit;
  };

  it("runs the checked menu function, writes its result and unchecks the box", async () => {
    const env = getEnv();
    const { onOpen, onEditInstalled } = await load();
    onOpen();
    onEditInstalled(edit(main, "A2"));
    expect(env.GmailApp.searches).toEqual([
      expect.objectContaining({ start: 0, max: 50 }),
    ]);
    // Result ({}) is written 2 columns right as JSON; box is reset to "FALSE"
    expect(main.getRange("C2").getValue()).toBe("{}");
    expect(main.getRange("A2").getValue()).toBe("FALSE");
  });

  it.each([
    ["A2", 0, 50],
    ["A3", undefined, undefined],
    ["A5", 0, 25],
    ["A6", undefined, undefined],
  ])("row %s dispatches to the right receipt import", async (cell, start, max) => {
    const env = getEnv();
    const { onOpen, onEditInstalled } = await load();
    onOpen();
    onEditInstalled(edit(main, cell));
    expect(env.GmailApp.searches).toHaveLength(1);
    expect(env.GmailApp.searches[0].start).toBe(start);
    expect(env.GmailApp.searches[0].max).toBe(max);
  });

  it("does nothing for the separator rows", async () => {
    const env = getEnv();
    const { onOpen, onEditInstalled } = await load();
    onOpen();
    onEditInstalled(edit(main, "A4"));
    expect(env.GmailApp.searches).toEqual([]);
    expect(main.getRange("A4").getValue()).toBe(true);
  });

  it("does nothing when the box is unchecked or the edit is outside the menu", async () => {
    const env = getEnv();
    const { onOpen, onEditInstalled } = await load();
    onOpen();
    onEditInstalled(edit(main, "A2", false));
    onEditInstalled(edit(main, "B2")); // wrong column, value true but not in checkbox column
    onEditInstalled(edit(main, "A20")); // below the menu
    expect(env.GmailApp.searches).toEqual([]);
  });

  it("runs the add-transaction-sheet actions and writes their return value", async () => {
    const env = getEnv();
    env.spreadsheet.addSheet("Template");
    addTransactionSheet(env, "2/18/24 - 3/2/24");
    const { onOpen, onEditInstalled } = await load();
    onOpen();
    // Far in the future so a sheet is needed. Fake only Date.now/new Date via vi
    vi.useFakeTimers();
    vi.setSystemTime(new Date("2024-03-10T12:00:00Z"));
    try {
      onEditInstalled(edit(main, "A8"));
      expect(main.getRange("C8").getValue()).toBe(true);
      expect(env.spreadsheet.getSheetByName("3/3/24 - 3/16/24")).not.toBeNull();
      // Check the "until up-to-date" action: already caught up through 3/16
      onEditInstalled(edit(main, "A9"));
      expect(main.getRange("C9").getValue()).toBe(0);
    } finally {
      vi.useRealTimers();
    }
  });

  it("leaves the checkbox checked if the action throws", async () => {
    const { onOpen, onEditInstalled } = await load();
    onOpen();
    // No transaction sheets -> addTransactionSheet throws
    expect(() => onEditInstalled(edit(main, "A8"))).toThrow(
      "No transaction sheets found in order to add more"
    );
    expect(main.getRange("A8").getValue()).toBe(true);
  });
});

describe("onEditInstalled - split checkboxes", () => {
  const D = new Date("2024-02-20T12:00:00Z");

  it("splits the transaction in the checked row and selects the new cost cell", async () => {
    const env = getEnv();
    const sheet = addTransactionSheet(env, "2/18/24 - 3/2/24");
    sheet.seed("A5", [
      [D, "Store", 5],
      [D, "Other", 6],
    ]);
    sheet.getRange("G5").setValue(true);
    const { onEditInstalled } = await load();
    onEditInstalled({ range: sheet.getRange("G5") } as unknown as GoogleAppsScript.Events.SheetsOnEdit);

    expect(sheet.getRange("B5:C6").getValues()).toEqual([
      ["Store | ", "=5-R[1]C[0]"],
      ["Store | ", "=1.0825*(0)"],
    ]);
    const active = env.SpreadsheetApp.activeRange!;
    expect([active.getRow(), active.getColumn()]).toEqual([6, 3]);
    expect(sheet.getRange("G5").getValue()).toBe("FALSE");
  });

  it("ignores checked cells outside the split checkbox column or rows", async () => {
    const env = getEnv();
    const sheet = addTransactionSheet(env, "2/18/24 - 3/2/24", { max: 3 });
    sheet.seed("A5", [[D, "Store", 5]]);
    sheet.getRange("H5").setValue(true);
    sheet.getRange("G20").setValue(true);
    const { onEditInstalled } = await load();
    for (const a1 of ["H5", "G20"])
      onEditInstalled({ range: sheet.getRange(a1) } as unknown as GoogleAppsScript.Events.SheetsOnEdit);
    expect(sheet.getRange("C5").getValue()).toBe(5);
    expect(env.SpreadsheetApp.activeRange).toBeUndefined();
  });
});

describe("addTransactionSheet / addTransactionSheets exports", () => {
  it("log how long they took and return the results", async () => {
    const env = getEnv();
    env.spreadsheet.addSheet("Template");
    addTransactionSheet(env, "2/18/24 - 3/2/24");
    vi.useFakeTimers();
    vi.setSystemTime(new Date("2024-03-01T12:00:00Z"));
    try {
      const { addTransactionSheet: add, addTransactionSheets: addAll } = await load();
      expect(add()).toBe(false);
      expect(addAll()).toBe(0);
      expect(env.logs.some((l) => /^Adding no transaction sheet took \d+ ms$/.test(String(l)))).toBe(true);
      expect(env.logs.some((l) => /^Adding 0 transaction sheets took \d+ ms$/.test(String(l)))).toBe(true);
    } finally {
      vi.useRealTimers();
    }
  });
});
