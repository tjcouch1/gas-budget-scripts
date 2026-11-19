/* eslint-disable @typescript-eslint/no-explicit-any */
import { computeRowFingerprint } from "@budget/shared";
import { afterEach, beforeEach, describe, expect, it, vi } from "vitest";
import {
  addVariablesSheet,
  AnySheet,
  DEFAULT_VARIABLES,
  GasEnv,
  getEnv,
  makeThread,
  MockRange,
  MockTextOutput,
} from "./gas-mocks";
import { chasePurchaseEmail } from "./fixtures/receipts";

const load = () => import("../src/index");

const KEY = "test-api-key-0123456789abcdef";
const SHEET = "2/18/24 - 3/2/24";
const OLDER_SHEET = "2/4/24 - 2/17/24 (Gap)";
/** 2024-02-20 in America/Chicago */
const D = new Date("2024-02-20T12:00:00Z");
const D2 = new Date("2024-02-21T12:00:00Z");

/** Envelope table rows (11 columns) starting at L5 */
const ENVELOPE_ROWS = [
  [
    "Envelope",
    "Avail",
    "Size",
    "Spent",
    "Remain",
    "Cumul",
    "a",
    "b",
    "c",
    "d",
    "e",
  ],
  ["Food", 120, 200, 50, 150, 300, 1000, 400, 40, 5200, 2080],
  ["", "", "", "", "", "", "", "", "", "", ""],
  ["Gas", 10, 50, 20, 30, 40, 100, 80, 8, 1300, 520],
  ["Fun", "n/a", 25, 0, 25, 25, 50, 0, 0, 650, 0],
  ["Totals", 130, 275, 70, 205, 365, 1150, 480, 48, 7150, 2600],
];
/** Paid With table (1 column used) starting at X5 */
const PAID_WITH_ROWS = [
  ["Credit"],
  ["Sam Venmo"],
  ["Tax deductible"],
  [""],
  ["Cash"],
];

/** Transaction sheet laid out for the API: transactions A5:G14, checkboxes J5, envelopes L5, paid with X5 */
function addApiSheet(
  env: GasEnv,
  name: string,
  opts: { withOptions?: boolean; max?: number } = {},
): AnySheet {
  const { withOptions = true, max = 10 } = opts;
  const sheet = env.spreadsheet.addSheet(name);
  const vars: unknown[][] = [
    ["TransactionsStart", "A5"],
    ["TransactionsMax", max],
    ["SplitCheckboxesStart", "J5"],
  ];
  if (withOptions)
    vars.push(
      ["EnvelopesStart", "L5"],
      ["EnvelopesMax", 6],
      ["PaidWithStart", "X5"],
      ["PaidWithMax", 5],
    );
  sheet.seed("H1", vars);
  if (withOptions) {
    sheet.seed("L5", ENVELOPE_ROWS);
    sheet.seed("X5", PAID_WITH_ROWS);
  }
  // Current Balance (col D) is a sheet formula on every row
  sheet.seed(
    "D5",
    Array.from({ length: max }, (_, i) => [`=BALANCE(${i})`]),
  );
  return sheet;
}

let env: GasEnv;
let sheet: AnySheet;

beforeEach(() => {
  env = getEnv();
  env.spreadsheet.addSheet("Main");
  addVariablesSheet(
    env,
    DEFAULT_VARIABLES.map(([k, v]) =>
      k === "SheetVariables" ? [k, "H1:I7"] : [k, v],
    ),
  );
  env.PropertiesService.scriptProperties.setProperty("API_KEY", KEY);
  sheet = addApiSheet(env, SHEET);
  addApiSheet(env, OLDER_SHEET, { withOptions: false });
});

afterEach(() => {
  vi.useRealTimers();
  vi.restoreAllMocks();
});

/** POST a raw body to doPost and parse the JSON response */
async function postRaw(contents: string | undefined): Promise<any> {
  const { doPost } = await load();
  const output = doPost(
    (contents === undefined
      ? {}
      : {
          postData: { contents },
        }) as unknown as GoogleAppsScript.Events.DoPost,
  ) as unknown as MockTextOutput;
  expect(output.getMimeType()).toBe("JSON");
  return JSON.parse(output.getContent());
}

const call = (request: Record<string, unknown>) =>
  postRaw(JSON.stringify(request));

const api = (action: string, fields: Record<string, unknown> = {}) =>
  call({ apiVersion: 1, apiKey: KEY, action, ...fields });

async function getState(sheetName = SHEET) {
  const res = await api("getState", { sheetName });
  expect(res.ok).toBe(true);
  return res.data;
}

async function fingerprintOf(rowIndex: number, sheetName = SHEET) {
  const state = await getState(sheetName);
  return state.transactions[rowIndex].fingerprint as string;
}

/** All non-empty cells of a mock sheet, for exact sheet comparisons */
function snapshot(s: AnySheet) {
  const cells = [...(s.cells as Map<string, any>).entries()]
    .filter(([, c]) => c.value !== "" || c.formula || c.note || c.background)
    .map(([key, c]) => [key, { ...c }]);
  return new Map(cells as [string, any][]);
}

/** Records every setValues call as [row, col, numRows, numCols] */
function spySetValues() {
  const writes: number[][] = [];
  const original = MockRange.prototype.setValues;
  vi.spyOn(MockRange.prototype, "setValues").mockImplementation(function (
    this: MockRange,
    values: unknown[][],
  ) {
    writes.push([
      this.getRow(),
      this.getColumn(),
      this.getNumRows(),
      this.getNumColumns(),
    ]);
    return original.call(this, values);
  });
  return writes;
}

describe("doGet", () => {
  it("returns only the service info without auth", async () => {
    env.PropertiesService.scriptProperties.deleteProperty("API_KEY");
    const { doGet } = await load();
    const output = doGet() as unknown as MockTextOutput;
    expect(output.getMimeType()).toBe("JSON");
    expect(JSON.parse(output.getContent())).toEqual({
      ok: true,
      data: { service: "budget-api", apiVersion: 1 },
    });
  });
});

describe("auth and request validation", () => {
  it.each([
    ["missing", {}],
    ["empty", { apiKey: "" }],
    ["wrong", { apiKey: "nope" }],
    ["a prefix of the key", { apiKey: KEY.slice(0, -1) }],
    ["not a string", { apiKey: 12345 }],
  ])("rejects a %s API key with unauthorized", async (_, keyField) => {
    const res = await call({ apiVersion: 1, action: "getState", ...keyField });
    expect(res).toEqual({
      ok: false,
      error: { code: "unauthorized", message: "Invalid or missing API key" },
    });
  });

  it("always rejects when the API_KEY Script Property is unset", async () => {
    env.PropertiesService.scriptProperties.deleteProperty("API_KEY");
    for (const apiKey of [KEY, "", undefined, null]) {
      const res = await call({ apiVersion: 1, apiKey, action: "getState" });
      expect(res.error.code).toBe("unauthorized");
    }
  });

  it("checks auth before the action or version", async () => {
    const res = await call({ apiVersion: 99, apiKey: "x", action: "nope" });
    expect(res.error.code).toBe("unauthorized");
  });

  it.each([
    ["no postData", undefined, "Missing request body"],
    ["non-JSON", "{nope", "Request body must be JSON"],
    ["a JSON array", "[]", "Request body must be a JSON object"],
    ["JSON null", "null", "Request body must be a JSON object"],
  ])("rejects %s with bad_request", async (_, body, message) => {
    const res = await postRaw(body);
    expect(res).toEqual({
      ok: false,
      error: { code: "bad_request", message },
    });
  });

  it("rejects a wrong apiVersion and an unknown action", async () => {
    expect(
      (await call({ apiVersion: 2, apiKey: KEY, action: "getState" })).error,
    ).toEqual({
      code: "bad_request",
      message: "Unsupported apiVersion; expected 1",
    });
    expect((await api("deleteEverything")).error).toEqual({
      code: "bad_request",
      message: "Unknown action",
    });
  });

  it("reports unexpected errors as server_error with only the message", async () => {
    env.spreadsheet.getSheets = () => {
      throw new Error("Service Spreadsheets failed");
    };
    const res = await api("getState");
    expect(res).toEqual({
      ok: false,
      error: { code: "server_error", message: "Service Spreadsheets failed" },
    });
  });
});

describe("API key setup", () => {
  it("setupGenerateApiKey stores a random 64-hex-character key and logs it once", async () => {
    env.PropertiesService.scriptProperties.deleteProperty("API_KEY");
    const { setupGenerateApiKey } = await load();
    setupGenerateApiKey();
    const key = env.PropertiesService.scriptProperties.getProperty("API_KEY")!;
    expect(key).toMatch(/^[0-9a-f]{64}$/);
    expect(env.logs.filter((l) => String(l).includes(key))).toHaveLength(1);
    // The new key works
    const res = await call({ apiVersion: 1, apiKey: key, action: "getState" });
    expect(res.ok).toBe(true);
  });

  it("setupGenerateApiKey refuses to overwrite an existing key", async () => {
    const { setupGenerateApiKey } = await load();
    expect(() => setupGenerateApiKey()).toThrow(/API_KEY already exists/);
    expect(env.PropertiesService.scriptProperties.getProperty("API_KEY")).toBe(
      KEY,
    );
  });

  it("generateApiKey(true) rotates the key", async () => {
    const { generateApiKey } = await import("../src/api/auth");
    generateApiKey(true);
    expect(
      env.PropertiesService.scriptProperties.getProperty("API_KEY"),
    ).not.toBe(KEY);
  });

  it("constantTimeEquals compares whole strings", async () => {
    const { constantTimeEquals } = await import("../src/api/auth");
    expect(constantTimeEquals("abc", "abc")).toBe(true);
    expect(constantTimeEquals("abc", "abd")).toBe(false);
    expect(constantTimeEquals("abc", "ab")).toBe(false);
    expect(constantTimeEquals("", "a")).toBe(false);
    expect(constantTimeEquals("", "")).toBe(true);
  });
});

describe("getState", () => {
  beforeEach(() => {
    sheet.seed("A5", [
      [D, "Store | ", 5.5, "=BALANCE(0)", "Food", "Credit", "Yes"],
      [D, "Store | milk", "=1.0825*(2)", "=BALANCE(1)", "", "Credit", ""],
      ["", "", "", "=BALANCE(2)", "", "", ""],
      [
        new Date("2024-02-22T03:00:00Z"),
        "Late night",
        3,
        "=BALANCE(3)",
        "Gas",
        "Sam Venmo",
        "",
      ],
    ]);
    sheet.getRange("C5").setValues([["=5.5-R[1]C[0]"]]);
    sheet.getRange("B8").setNote("a note").setBackground("#E8A9CA");
    sheet.getRange("B6").setBackground("#ff0000");
  });

  it("returns the sheets newest first, the variables, and the requested sheet", async () => {
    const res = await api("getState", { sheetName: SHEET });
    expect(res.ok).toBe(true);
    expect(res.warnings).toBeUndefined();
    const state = res.data;
    expect(state.sheets).toEqual([
      {
        name: SHEET,
        startDate: "2024-02-18",
        endDate: "2024-03-02",
        isGap: false,
      },
      {
        name: OLDER_SHEET,
        startDate: "2024-02-04",
        endDate: "2024-02-17",
        isGap: true,
      },
    ]);
    expect(state.sheetName).toBe(SHEET);
    expect(state.taxMultiplier).toBe(1.0825);
    expect(state.forwardName).toBe("Sam");
  });

  it("serializes transactions, trimming trailing empty rows and keeping interior ones as null", async () => {
    const { transactions } = await getState();
    expect(transactions).toHaveLength(4);
    expect(transactions[2]).toBeNull();
    const [parent, child, , late] = transactions;
    expect(parent).toEqual({
      rowIndex: 0,
      date: "2024-02-20",
      name: "Store | ",
      cost: null,
      costFormula: "=5.5-R[1]C[0]",
      costDisplay: "=5.5-R[1]C[0]",
      envelope: "Food",
      paidWith: "Credit",
      taxDeductible: "Yes",
      note: "",
      isError: false,
      isNote: false,
      splitGroupId: "0",
      fingerprint: expect.stringMatching(/^[0-9a-f]{8}$/),
    });
    expect(child).toMatchObject({
      rowIndex: 1,
      name: "Store | milk",
      isError: true,
      isNote: false,
      splitGroupId: "0",
    });
    // Dates are formatted in the spreadsheet time zone (03:00Z is still the 21st in Chicago)
    expect(late).toMatchObject({
      rowIndex: 3,
      date: "2024-02-21",
      cost: 3,
      costFormula: null,
      costDisplay: "3",
      note: "a note",
      isNote: true,
      isError: false,
      splitGroupId: null,
    });
    // The fingerprint is the shared function of the DTO fields
    expect(late.fingerprint).toBe(computeRowFingerprint(late));
  });

  it("reads envelopes and paid-with options, skipping blank, header, total and tax rows", async () => {
    const state = await getState();
    expect(state.envelopeOptions).toEqual(["Food", "Gas", "Fun"]);
    expect(state.paidWithOptions).toEqual(["Credit", "Sam Venmo", "Cash"]);
    expect(state.envelopes[0]).toEqual({
      name: "Food",
      display: [
        "120",
        "200",
        "50",
        "150",
        "300",
        "1000",
        "400",
        "40",
        "5200",
        "2080",
      ],
      values: [120, 200, 50, 150, 300, 1000, 400, 40, 5200, 2080],
    });
    // Non-numeric cells keep their display value and have a null number
    expect(state.envelopes[2].display[0]).toBe("n/a");
    expect(state.envelopes[2].values[0]).toBeNull();
  });

  it("returns null envelopes/options plus warnings when the optional sheet variables are missing", async () => {
    const res = await api("getState", { sheetName: OLDER_SHEET });
    expect(res.ok).toBe(true);
    expect(res.data.envelopes).toBeNull();
    expect(res.data.envelopeOptions).toBeNull();
    expect(res.data.paidWithOptions).toBeNull();
    expect(res.data.transactions).toEqual([]);
    expect(res.warnings).toEqual([
      `Sheet "${OLDER_SHEET}" has no EnvelopesStart/EnvelopesMax sheet variables. Envelope balances and Envelope editing are unavailable for it.`,
      `Sheet "${OLDER_SHEET}" has no PaidWithStart/PaidWithMax sheet variables. Paid with editing is unavailable for it.`,
    ]);
  });

  it("warns about an invalid EnvelopesMax instead of failing", async () => {
    sheet.getRange("I5").setValue("lots");
    const res = await api("getState");
    expect(res.ok).toBe(true);
    expect(res.data.envelopes).toBeNull();
    expect(res.warnings[0]).toContain("invalid EnvelopesMax");
  });

  it("defaults to the sheet containing today, else the newest sheet", async () => {
    vi.useFakeTimers();
    vi.setSystemTime(new Date("2024-02-10T18:00:00Z"));
    expect((await api("getState")).data.sheetName).toBe(OLDER_SHEET);
    vi.setSystemTime(new Date("2025-01-01T18:00:00Z"));
    expect((await api("getState")).data.sheetName).toBe(SHEET);
  });

  it("formats pay-period end dates correctly during daylight saving time", async () => {
    addApiSheet(env, "6/2/24 - 6/15/24");
    const { sheets } = await getState("6/2/24 - 6/15/24");
    expect(sheets[0]).toEqual({
      name: "6/2/24 - 6/15/24",
      startDate: "2024-06-02",
      endDate: "2024-06-15",
      isGap: false,
    });
  });

  it("returns not_found for an unknown or non-transaction sheet", async () => {
    for (const sheetName of ["1/1/30 - 1/14/30", "Variables", "Main"]) {
      const res = await api("getState", { sheetName });
      expect(res.error.code).toBe("not_found");
    }
    expect((await api("getState", { sheetName: 5 })).error.code).toBe(
      "bad_request",
    );
  });

  it("does not take the script lock (read only)", async () => {
    env.LockService.scriptLock.contended = true;
    expect((await api("getState")).ok).toBe(true);
    expect(env.LockService.scriptLock.calls).toEqual([]);
  });
});

describe("getAllTransactions", () => {
  beforeEach(() => {
    sheet.seed("A5", [
      [D, "Store | ", 5.5, "=BALANCE(0)", "Food", "Credit", "Yes"],
      [D2, "Cafe", 3, "=BALANCE(1)", "", "Credit", ""],
    ]);
    env.spreadsheet
      .getSheetByName(OLDER_SHEET)!
      .seed("A5", [[D, "Old shop", 9, "=BALANCE(0)", "", "", ""]]);
  });

  it("returns every sheet newest first with transactions, fingerprints and option lists", async () => {
    const res = await api("getAllTransactions");
    expect(res.ok).toBe(true);
    expect(res.data.taxMultiplier).toBe(1.0825);
    expect(res.data.sheets.map((x: any) => x.name)).toEqual([
      SHEET,
      OLDER_SHEET,
    ]);
    const [current, older] = res.data.sheets;
    expect(current).toMatchObject({
      startDate: "2024-02-18",
      endDate: "2024-03-02",
      isGap: false,
      envelopeOptions: ["Food", "Gas", "Fun"],
      paidWithOptions: ["Credit", "Sam Venmo", "Cash"],
    });
    expect(current.transactions).toHaveLength(2);
    expect(older.isGap).toBe(true);
    expect(older.envelopeOptions).toBeNull();
    expect(older.paidWithOptions).toBeNull();
    expect(older.transactions[0].name).toBe("Old shop");
    // Rows are identical to the per-sheet getState (same fingerprints), so edits can use them
    expect(current.transactions).toEqual((await getState(SHEET)).transactions);
    expect(older.transactions).toEqual(
      (await getState(OLDER_SHEET)).transactions,
    );
  });

  it("reads only the named sheets when sheetNames is given", async () => {
    const res = await api("getAllTransactions", { sheetNames: [OLDER_SHEET] });
    expect(res.ok).toBe(true);
    expect(res.data.sheets.map((x: any) => x.name)).toEqual([OLDER_SHEET]);
    expect(res.data.sheets[0].transactions[0].name).toBe("Old shop");
    expect(res.data.taxMultiplier).toBe(1.0825);
  });

  it("rejects an unknown or malformed sheetNames", async () => {
    expect(
      (await api("getAllTransactions", { sheetNames: ["Nope"] })).error.code,
    ).toBe("not_found");
    for (const sheetNames of [[], "x", [1], new Array(26).fill(SHEET)])
      expect((await api("getAllTransactions", { sheetNames })).error.code).toBe(
        "bad_request",
      );
  });

  it("lets a row from another sheet be edited with its own sheetName and fingerprint", async () => {
    const { data } = await api("getAllTransactions");
    const row = data.sheets[1].transactions[0];
    const res = await api("updateTransaction", {
      sheetName: OLDER_SHEET,
      rowIndex: row.rowIndex,
      baseFingerprint: row.fingerprint,
      changes: { name: "Renamed" },
    });
    expect(res.ok).toBe(true);
    expect(res.data.name).toBe("Renamed");
  });

  it("reads each sheet with one batch read per kind of data and no per-row I/O", async () => {
    const reads: string[] = [];
    for (const method of [
      "getValues",
      "getFormulas",
      "getDisplayValues",
      "getNotes",
      "getBackgrounds",
    ] as const) {
      const original = MockRange.prototype[method];
      vi.spyOn(MockRange.prototype, method).mockImplementation(function (
        this: MockRange,
      ) {
        if (this.getNumRows() === 1 && this.getNumColumns() === 1)
          reads.push(`single-cell ${method}`);
        reads.push(method);
        return (original as () => unknown).call(this) as never;
      } as never);
    }
    expect((await api("getAllTransactions")).ok).toBe(true);
    expect(reads.filter((r) => r.startsWith("single-cell"))).toEqual([]);
    // Per sheet: 5 transaction reads, plus the option tables (values + display + paid with)
    expect(reads.filter((r) => r === "getNotes")).toHaveLength(2);
    expect(reads.filter((r) => r === "getFormulas")).toHaveLength(2);
  });

  it("is read only: no script lock, and it needs the API key", async () => {
    env.LockService.scriptLock.contended = true;
    expect((await api("getAllTransactions")).ok).toBe(true);
    expect(env.LockService.scriptLock.calls).toEqual([]);
    const bad = await call({
      apiVersion: 1,
      apiKey: "wrong",
      action: "getAllTransactions",
    });
    expect(bad.error.code).toBe("unauthorized");
  });
});

describe("updateTransaction", () => {
  beforeEach(() => {
    sheet.seed("A5", [
      [D, "Store", 5.5, "=BALANCE(0)", "", "Credit", ""],
      [D2, "Cafe", "=3+2", "=BALANCE(1)", "Food", "Credit", ""],
    ]);
  });

  async function update(
    rowIndex: number,
    changes: Record<string, unknown>,
    extra: Record<string, unknown> = {},
  ) {
    return api("updateTransaction", {
      sheetName: SHEET,
      rowIndex,
      baseFingerprint: await fingerprintOf(rowIndex),
      changes,
      ...extra,
    });
  }

  it("writes envelope, paid with, and tax deductible with one write to cols 4-6 only", async () => {
    const fp = await fingerprintOf(0);
    const writes = spySetValues();
    const res = await api("updateTransaction", {
      sheetName: SHEET,
      rowIndex: 0,
      baseFingerprint: fp,
      changes: { envelope: "Gas", paidWith: "Cash", taxDeductible: "Yes" },
    });
    expect(res.ok).toBe(true);
    expect(writes).toEqual([[5, 5, 1, 3]]);
    expect(sheet.getRange("A5:G5").getValues()).toEqual([
      [D, "Store", 5.5, "=BALANCE(0)", "Gas", "Cash", "Yes"],
    ]);
    expect(sheet.getRange("D5").getFormulas()).toEqual([["=BALANCE(0)"]]);
    expect(res.data).toMatchObject({
      rowIndex: 0,
      envelope: "Gas",
      paidWith: "Cash",
      taxDeductible: "Yes",
    });
    expect(res.data.fingerprint).not.toBe(fp);
    expect(res.data.fingerprint).toBe(await fingerprintOf(0));
  });

  it("writes unchanged metadata cells back as read", async () => {
    const res = await update(1, { taxDeductible: "Yes" });
    expect(res.ok).toBe(true);
    expect(sheet.getRange("E6:G6").getValues()).toEqual([
      ["Food", "Credit", "Yes"],
    ]);
  });

  it("writes name and cost with one write to cols 1-2 only", async () => {
    const fp = await fingerprintOf(0);
    const writes = spySetValues();
    const res = await api("updateTransaction", {
      sheetName: SHEET,
      rowIndex: 0,
      baseFingerprint: fp,
      changes: { name: "Grocery Store", costExpression: "1+3.23*2" },
    });
    expect(res.ok).toBe(true);
    expect(writes).toEqual([[5, 2, 1, 2]]);
    expect(sheet.getRange("A5:G5").getValues()).toEqual([
      [D, "Grocery Store", "=1+3.23*2", "=BALANCE(0)", "", "Credit", ""],
    ]);
    expect(res.data.costFormula).toBe("=1+3.23*2");
  });

  it.each([
    ["12.5", undefined, 12.5],
    ["-4", false, -4],
    ["=2*3", false, "=2*3"],
    ["(1+2)", false, "=(1+2)"],
    ["4.23*2+12.09", true, "=1.0825*(4.23*2+12.09)"],
    ["4", true, "=1.0825*(4)"],
  ])(
    "writes cost expression %s (pre-tax %s) as %s",
    async (costExpression, costPreTax, expected) => {
      const res = await update(0, { costExpression, costPreTax });
      expect(res.ok).toBe(true);
      expect(sheet.getRange("C5").getValue()).toBe(expected);
    },
  );

  it("keeps an unchanged formula cost when editing the name", async () => {
    const res = await update(1, { name: "Coffee shop" });
    expect(res.ok).toBe(true);
    expect(sheet.getRange("B6:C6").getValues()).toEqual([
      ["Coffee shop", "=3+2"],
    ]);
    expect(sheet.getRange("C6").getFormulas()).toEqual([["=3+2"]]);
  });

  it("allows clearing the envelope", async () => {
    const res = await update(1, { envelope: "" });
    expect(res.ok).toBe(true);
    expect(sheet.getRange("E6").getValue()).toBe("");
  });

  it.each([
    [
      { name: '=IMPORTXML("http://x","//a")' },
      "changes.name must not start with =, +, -, or @",
    ],
    [{ name: "+1" }, "changes.name must not start with =, +, -, or @"],
    [{ name: " -1" }, "changes.name must not start with =, +, -, or @"],
    [{ name: "@me" }, "changes.name must not start with =, +, -, or @"],
    [{ name: 5 }, "changes.name must be a string"],
    [
      { envelope: '=HYPERLINK("x")' },
      "changes.envelope must not start with =, +, -, or @",
    ],
    [
      { envelope: "Rent" },
      'changes.envelope "Rent" is not one of the sheet\'s options',
    ],
    [
      { envelope: "Totals" },
      'changes.envelope "Totals" is not one of the sheet\'s options',
    ],
    [
      { paidWith: "Bitcoin" },
      'changes.paidWith "Bitcoin" is not one of the sheet\'s options',
    ],
    [
      { paidWith: "Tax deductible" },
      'changes.paidWith "Tax deductible" is not one of the sheet\'s options',
    ],
    [
      { costExpression: "1+" },
      "changes.costExpression is not a valid expression",
    ],
    [
      { costExpression: "SUM(1)" },
      "changes.costExpression is not a valid expression",
    ],
    [
      { costExpression: "=IMPORTXML(1)" },
      "changes.costExpression is not a valid expression",
    ],
    [
      { costExpression: "2/0" },
      "changes.costExpression is not a valid expression",
    ],
    [
      { costPreTax: true },
      "changes.costPreTax requires changes.costExpression",
    ],
    [
      { costExpression: "1", costPreTax: "yes" },
      "changes.costPreTax must be a boolean",
    ],
    [{ taxDeductible: "No" }, 'changes.taxDeductible must be "Yes" or ""'],
    [{ date: "2024-01-01" }, "Unknown changes: date"],
    [{ currentBalance: 1 }, "Unknown changes: currentBalance"],
    [{}, "No changes given"],
  ])("rejects %j with bad_request", async (changes, message) => {
    const before = snapshot(sheet);
    const res = await update(0, changes);
    expect(res).toEqual({ ok: false, error: { code: "bad_request", message } });
    expect(snapshot(sheet)).toEqual(before);
    // Rejected before taking the lock
    expect(env.LockService.scriptLock.calls).toEqual([]);
  });

  it("rejects bad targets", async () => {
    const fp = await fingerprintOf(0);
    const base = {
      sheetName: SHEET,
      rowIndex: 0,
      baseFingerprint: fp,
      changes: { envelope: "Gas" },
    };
    for (const [fields, code, message] of [
      [
        { rowIndex: 10 },
        "bad_request",
        "rowIndex must be an integer in [0, 10)",
      ],
      [
        { rowIndex: -1 },
        "bad_request",
        "rowIndex must be an integer in [0, 10)",
      ],
      [
        { rowIndex: 0.5 },
        "bad_request",
        "rowIndex must be an integer in [0, 10)",
      ],
      [
        { rowIndex: "0" },
        "bad_request",
        "rowIndex must be an integer in [0, 10)",
      ],
      [
        { sheetName: "Variables" },
        "not_found",
        'No transaction sheet named "Variables"',
      ],
      [
        { baseFingerprint: "" },
        "bad_request",
        "baseFingerprint must not be empty",
      ],
      [
        { baseFingerprint: undefined },
        "bad_request",
        "baseFingerprint must be a string",
      ],
      [{ changes: "nope" }, "bad_request", "changes must be an object"],
      [{ force: "yes" }, "bad_request", "force must be a boolean"],
    ] as const) {
      const res = await api("updateTransaction", { ...base, ...fields });
      expect(res.error).toEqual({ code, message });
    }
  });

  it("rejects an edit of an empty row", async () => {
    // Empty rows are not in getState, so the client could only force it
    const res = await api("updateTransaction", {
      sheetName: SHEET,
      rowIndex: 4,
      baseFingerprint: "none",
      changes: { envelope: "Gas" },
      force: true,
    });
    expect(res.error).toEqual({
      code: "bad_request",
      message: "Row 4 is empty",
    });
    expect(sheet.getRange("E9").getValue()).toBe("");
  });

  it("lets a split-group row change only the item note after the |", async () => {
    sheet.seed("A5", [
      [D, "Store | ", "=5.5-R[1]C[0]", "=BALANCE(0)", "", "Credit", ""],
      [D, "Store | milk", 2, "=BALANCE(1)", "", "Credit", ""],
    ]);
    expect((await update(1, { name: "Store | whole milk" })).ok).toBe(true);
    expect((await update(0, { name: "Store | groceries" })).ok).toBe(true);
    expect(sheet.getRange("B5:C6").getValues()).toEqual([
      ["Store | groceries", "=5.5-R[1]C[0]"],
      ["Store | whole milk", 2],
    ]);
    for (const name of ["Other | milk", "Store milk", "Stor | milk"])
      expect((await update(1, { name })).error?.code).toBe("bad_request");
  });

  it("rejects cost edits on a split parent and name edits on split-group rows", async () => {
    sheet.seed("A5", [
      [D, "Store | ", "=5.5-R[1]C[0]", "=BALANCE(0)", "", "Credit", ""],
      [D, "Store | milk", 2, "=BALANCE(1)", "", "Credit", ""],
    ]);
    expect((await update(0, { costExpression: "6" })).error).toEqual({
      code: "bad_request",
      message:
        "Cannot edit the cost of a split parent row; edit the child rows",
    });
    expect((await update(1, { name: "Milk" })).error).toEqual({
      code: "bad_request",
      message:
        "Cannot rename a row in a split group (its name contains |) because that could break the group",
    });
    expect(sheet.getRange("B5:C6").getValues()).toEqual([
      ["Store | ", "=5.5-R[1]C[0]"],
      ["Store | milk", 2],
    ]);
    // Children's costs and the group's other fields stay editable
    expect(
      (await update(1, { costExpression: "3", envelope: "Food" })).ok,
    ).toBe(true);
    expect((await update(0, { envelope: "Food" })).ok).toBe(true);
  });

  it("rejects renaming a split group (that writes unchecked child rows; use updateSplitGroup) but allows the parent's own item note", async () => {
    sheet.seed("A5", [
      [
        D,
        "Store | ",
        "=5.5-R[1]C[0]-R[2]C[0]",
        "=BALANCE(0)",
        "",
        "Credit",
        "",
      ],
      [D, "Store | milk", 2, "=BALANCE(1)", "", "Credit", ""],
      [D, "Store | eggs", 1, "=BALANCE(2)", "", "Credit", ""],
      [D2, "Cafe", 3, "=BALANCE(3)", "", "Credit", ""],
    ]);
    const before = snapshot(sheet);
    for (const name of ["Market", "Shop | food"])
      expect((await update(0, { name })).error).toEqual({
        code: "bad_request",
        message:
          "Cannot rename a split group with updateTransaction; use updateSplitGroup so every row is checked",
      });
    expect(snapshot(sheet)).toEqual(before);
    // The parent's own item note (same text before "|") is still its own edit
    expect((await update(0, { name: "Store | food" })).ok).toBe(true);
    expect(sheet.getRange("B5:B7").getValues()).toEqual([
      ["Store | food"],
      ["Store | milk"],
      ["Store | eggs"],
    ]);
    // A child still cannot rename the group
    expect((await update(1, { name: "Other | milk" })).error?.code).toBe(
      "bad_request",
    );
  });

  it("returns the whole state when a split child's cost changes", async () => {
    sheet.seed("A5", [
      [D, "Store | ", "=5.5-R[1]C[0]", "=BALANCE(0)", "", "Credit", ""],
      [D, "Store | milk", 2, "=BALANCE(1)", "", "Credit", ""],
    ]);
    const res = await update(1, { costExpression: "3" });
    expect(res.ok).toBe(true);
    expect(res.data.sheetName).toBe(SHEET);
    expect(res.data.transactions[1]).toMatchObject({ rowIndex: 1, cost: 3 });
    // Other edits of a group row still return just the row
    const env2 = await update(1, { envelope: "Food" });
    expect(env2.data).toMatchObject({ rowIndex: 1, envelope: "Food" });
  });

  it("replaces a split parent's total with a plain number or validated math", async () => {
    sheet.seed("A5", [
      [D, "Store | ", "=45-R[1]C[0]-R[2]C[0]", "=BALANCE(0)", "", "Credit", ""],
      [D, "Store | milk", 2, "=BALANCE(1)", "", "Credit", ""],
      [D, "Store | eggs", 1, "=BALANCE(2)", "", "Credit", ""],
    ]);
    const res = await update(0, { splitTotal: "50.125" });
    expect(res.ok).toBe(true);
    expect(sheet.getRange("C5:C7").getValues()).toEqual([
      ["=50.125-R[1]C[0]-R[2]C[0]"],
      [2],
      [1],
    ]);
    // Math is written in parentheses before the child references
    expect((await update(0, { splitTotal: "40+5" })).ok).toBe(true);
    expect(sheet.getRange("C5").getValue()).toBe("=(40+5)-R[1]C[0]-R[2]C[0]");
    expect((await update(0, { splitTotal: "50.125" })).ok).toBe(true);
    for (const splitTotal of [
      "1+",
      "4*",
      "1..2",
      "(1+2",
      "--1",
      "$5",
      "1,000",
      "",
      "abc",
      "=SUM(A1)",
    ])
      expect((await update(0, { splitTotal })).error).toEqual({
        code: "bad_request",
        message: "changes.splitTotal must be a number or valid math",
      });
    expect((await update(1, { splitTotal: "5" })).error).toEqual({
      code: "bad_request",
      message: "changes.splitTotal is only allowed on a split parent row",
    });
    expect(
      (await update(0, { splitTotal: "5", costExpression: "5" })).error?.code,
    ).toBe("bad_request");
    expect(sheet.getRange("C5").getValue()).toBe("=50.125-R[1]C[0]-R[2]C[0]");
  });

  it("returns conflict with the current row when the row changed since the client loaded it", async () => {
    const staleFp = await fingerprintOf(0);
    // Someone else edits the row on the sheet
    sheet.getRange("E5").setValue("Food");
    const before = snapshot(sheet);
    const res = await api("updateTransaction", {
      sheetName: SHEET,
      rowIndex: 0,
      baseFingerprint: staleFp,
      changes: { envelope: "Gas" },
    });
    expect(res.ok).toBe(false);
    expect(res.error.code).toBe("conflict");
    expect(res.conflict.current).toMatchObject({
      rowIndex: 0,
      envelope: "Food",
    });
    expect(res.conflict.current.fingerprint).not.toBe(staleFp);
    expect(snapshot(sheet)).toEqual(before);
    expect(env.LockService.scriptLock.held).toBe(false);

    // "Overwrite with Mine": resend with the current fingerprint
    const retry = await api("updateTransaction", {
      sheetName: SHEET,
      rowIndex: 0,
      baseFingerprint: res.conflict.current.fingerprint,
      changes: { envelope: "Gas" },
    });
    expect(retry.ok).toBe(true);
    expect(sheet.getRange("E5").getValue()).toBe("Gas");
  });

  it("edits the note alone with one setNote on the Name cell, leaving the background", async () => {
    sheet.getRange("B5").setBackground("#E8A9CA");
    const writes = spySetValues();
    const res = await update(0, { note: "new note" }, { baseNote: "" });
    expect(res.ok).toBe(true);
    expect(writes).toEqual([]);
    expect(sheet.getRange("B5").getNotes()).toEqual([["new note"]]);
    expect(sheet.getRange("B5").getBackgrounds()).toEqual([["#E8A9CA"]]);
    expect(res.data.note).toBe("new note");
    const cleared = await update(0, { note: "" }, { baseNote: "new note" });
    expect(cleared.ok).toBe(true);
    expect(sheet.getRange("B5").getNotes()).toEqual([[""]]);
  });

  it("requires baseNote with changes.note and limits the note length", async () => {
    const missing = await update(0, { note: "x" });
    expect(missing.ok).toBe(false);
    expect(missing.error.code).toBe("bad_request");
    const long = await update(0, { note: "x".repeat(5001) }, { baseNote: "" });
    expect(long.ok).toBe(false);
    expect(long.error.code).toBe("bad_request");
    expect(sheet.getRange("B5").getNotes()).toEqual([[""]]);
  });

  it("returns conflict when the note changed on the sheet, unless forced", async () => {
    sheet.getRange("B5").setNote("theirs");
    const res = await update(0, { note: "mine" }, { baseNote: "" });
    expect(res.ok).toBe(false);
    expect(res.error.code).toBe("conflict");
    expect(res.conflict.current).toMatchObject({ note: "theirs" });
    expect(sheet.getRange("B5").getNotes()).toEqual([["theirs"]]);
    const forced = await update(
      0,
      { note: "mine" },
      { baseNote: "", force: true },
    );
    expect(forced.ok).toBe(true);
    expect(sheet.getRange("B5").getNotes()).toEqual([["mine"]]);
  });

  it("force skips the fingerprint check", async () => {
    const res = await api("updateTransaction", {
      sheetName: SHEET,
      rowIndex: 0,
      baseFingerprint: "stale",
      changes: { envelope: "Gas" },
      force: true,
    });
    expect(res.ok).toBe(true);
    expect(sheet.getRange("E5").getValue()).toBe("Gas");
  });

  it("returns busy without writing when the script lock is held elsewhere", async () => {
    const fp = await fingerprintOf(0);
    env.LockService.scriptLock.contended = true;
    const res = await api("updateTransaction", {
      sheetName: SHEET,
      rowIndex: 0,
      baseFingerprint: fp,
      changes: { envelope: "Gas" },
    });
    expect(res.error.code).toBe("busy");
    expect(sheet.getRange("E5").getValue()).toBe("");
    expect(env.LockService.scriptLock.calls).toEqual(["tryLock:20000"]);
  });

  it("takes the script lock (tryLock 20 s) and releases it", async () => {
    await update(0, { envelope: "Gas" });
    expect(env.LockService.scriptLock.calls).toEqual([
      "tryLock:20000",
      "release",
    ]);
    expect(env.LockService.scriptLock.held).toBe(false);
  });

  it("rejects non-blank envelope/paid with on a sheet without option lists but allows blank", async () => {
    const older = env.spreadsheet.getSheetByName(OLDER_SHEET);
    older.seed("A5", [
      [new Date("2024-02-10T12:00:00Z"), "Old", 1, "=B", "X", "Credit", ""],
    ]);
    const fp = await fingerprintOf(0, OLDER_SHEET);
    const base = { sheetName: OLDER_SHEET, rowIndex: 0, baseFingerprint: fp };
    expect(
      (
        await api("updateTransaction", {
          ...base,
          changes: { envelope: "Food" },
        })
      ).error,
    ).toEqual({
      code: "bad_request",
      message:
        "changes.envelope cannot be set on this sheet because it has no EnvelopesStart/EnvelopesMax sheet variables",
    });
    expect(
      (
        await api("updateTransaction", {
          ...base,
          changes: { paidWith: "Credit" },
        })
      ).error.code,
    ).toBe("bad_request");
    const res = await api("updateTransaction", {
      ...base,
      changes: { envelope: "", costExpression: "2" },
    });
    expect(res.ok).toBe(true);
    expect(older.getRange("B5:F5").getValues()).toEqual([
      ["Old", 2, "=B", "", "Credit"],
    ]);
  });
});

describe("splitTransaction", () => {
  const split = async (
    rowIndex: number,
    splits: unknown,
    extra: Record<string, unknown> = {},
  ) =>
    api("splitTransaction", {
      sheetName: SHEET,
      rowIndex,
      baseFingerprint: await fingerprintOf(rowIndex),
      splits,
      ...extra,
    });

  const item = (fields: Record<string, unknown> = {}) => ({
    itemNote: "",
    expression: "0",
    preTax: true,
    envelope: "",
    ...fields,
  });

  /** Values (A:G) of rows 5.. */
  const rows = (n: number) => sheet.getRange(5, 1, n, 7).getValues();

  it("a one-row web split produces exactly the same sheet as the checkbox (native) split", async () => {
    const seedRows = [
      [D, "Store", 5.5, "=BALANCE(0)", "Food", "Credit", "Yes"],
      [D2, "Next", 7, "=BALANCE(1)", "Gas", "Sam Venmo", ""],
      [D2, "After", "=1+2", "=BALANCE(2)", "", "Credit", "Yes"],
    ];
    const nativeSheet = addApiSheet(env, "3/3/24 - 3/16/24");
    for (const s of [sheet, nativeSheet]) {
      s.seed("A5", seedRows);
      s.getRange("B6").setNote("note").setBackground("#E8A9CA");
      s.getRange("C7").setBackground("#FF0000");
    }

    const { Budgeting } = await import("../src/util/budgeting.util");
    Budgeting.splitTransaction(nativeSheet, 0);

    const res = await split(0, [item()]);
    expect(res.ok).toBe(true);
    expect(snapshot(sheet)).toEqual(snapshot(nativeSheet));
    expect(rows(4)).toEqual([
      [D, "Store | ", "=5.5-R[1]C[0]", "=BALANCE(0)", "Food", "Credit", "Yes"],
      [D, "Store | ", "=1.0825*(0)", "=BALANCE(1)", "", "Credit", ""],
      [D2, "Next", 7, "=BALANCE(2)", "Gas", "Sam Venmo", ""],
      [D2, "After", "=1+2", "=BALANCE(3)", "", "Credit", "Yes"],
    ]);
  });

  it("parity also holds when splitting an already-split group again", async () => {
    const seedRows = [
      [D, "Store | ", "=5.5-R[1]C[0]", "=BALANCE(0)", "Food", "Credit", ""],
      [D, "Store | bread", "=1.0825*(2)", "=BALANCE(1)", "Food", "Credit", ""],
      [D2, "Next", 7, "=BALANCE(2)", "Gas", "Credit", ""],
    ];
    const nativeSheet = addApiSheet(env, "3/3/24 - 3/16/24");
    sheet.seed("A5", seedRows);
    nativeSheet.seed("A5", seedRows);
    const { Budgeting } = await import("../src/util/budgeting.util");
    Budgeting.splitTransaction(nativeSheet, 0);
    expect((await split(0, [item()])).ok).toBe(true);
    expect(snapshot(sheet)).toEqual(snapshot(nativeSheet));
  });

  it("splits into 3 rows with mixed pre-tax, envelopes, and paid with", async () => {
    sheet.seed("A5", [[D, "Costco", 40, "=BALANCE(0)", "", "Credit", "Yes"]]);
    const res = await split(0, [
      item({
        itemNote: "milk",
        expression: "4.23*2+1",
        preTax: true,
        envelope: "Food",
      }),
      item({
        itemNote: "gas",
        expression: "1+3.23*2",
        preTax: false,
        envelope: "Gas",
        paidWith: "Cash",
      }),
      item({
        itemNote: "toy",
        expression: "5",
        preTax: false,
        envelope: "Fun",
        paidWith: "",
      }),
    ]);
    expect(res.ok).toBe(true);
    expect(rows(4)).toEqual([
      [
        D,
        "Costco | ",
        "=40-R[1]C[0]-R[2]C[0]-R[3]C[0]",
        "=BALANCE(0)",
        "",
        "Credit",
        "Yes",
      ],
      [
        D,
        "Costco | milk",
        "=1.0825*(4.23*2+1)",
        "=BALANCE(1)",
        "Food",
        "Credit",
        "",
      ],
      [D, "Costco | gas", "=1+3.23*2", "=BALANCE(2)", "Gas", "Cash", ""],
      [D, "Costco | toy", 5, "=BALANCE(3)", "Fun", "Credit", ""],
    ]);
    // Returns the refreshed state, where the 4 rows are one split group
    expect(res.data.sheetName).toBe(SHEET);
    expect(res.data.transactions.map((t: any) => t.splitGroupId)).toEqual([
      "0",
      "0",
      "0",
      "0",
    ]);
  });

  it("appends to an already-split group", async () => {
    sheet.seed("A5", [
      [D, "Store | ", "=5.5-R[1]C[0]", "=BALANCE(0)", "", "Credit", ""],
      [D, "Store | bread", "=1.0825*(2)", "=BALANCE(1)", "Food", "Credit", ""],
    ]);
    const res = await split(0, [
      item({ itemNote: "eggs", expression: "1" }),
      item({ itemNote: "jam", expression: "2", preTax: false }),
    ]);
    expect(res.ok).toBe(true);
    expect(rows(4).map((r) => r.slice(1, 3))).toEqual([
      ["Store | ", "=5.5-R[1]C[0]-R[2]C[0]-R[3]C[0]"],
      ["Store | bread", "=1.0825*(2)"],
      ["Store | eggs", "=1.0825*(1)"],
      ["Store | jam", 2],
    ]);
  });

  it("writes new children's Tax Deductible and Name notes in one batch each; rows below are untouched", async () => {
    sheet.seed("A5", [
      [D, "Store | ", "=5.5-R[1]C[0]", "=BALANCE(0)", "", "Credit", "Yes"],
      [D, "Store | bread", "=1.0825*(2)", "=BALANCE(1)", "Food", "Credit", ""],
      [D2, "Next", 7, "=BALANCE(2)", "Gas", "Cash", "Yes"],
      [D2, "After", 3, "=BALANCE(3)", "", "Credit", ""],
    ]);
    sheet.getRange("B7").setNote("next note");
    sheet.getRange("B6").setNote("bread note");
    const noteWrites: number[][] = [];
    const originalSetNotes = MockRange.prototype.setNotes;
    vi.spyOn(MockRange.prototype, "setNotes").mockImplementation(function (
      this: MockRange,
      notes: (string | null)[][],
    ) {
      noteWrites.push([
        this.getRow(),
        this.getColumn(),
        this.getNumRows(),
        this.getNumColumns(),
      ]);
      return originalSetNotes.call(this, notes);
    });
    const res = await split(0, [
      item({ itemNote: "eggs", expression: "1", taxDeductible: "Yes" }),
      item({ itemNote: "jam", expression: "2" }),
      item({ itemNote: "tea", expression: "3", note: "for Pat" }),
    ]);
    expect(res.ok).toBe(true);
    expect(rows(7).map((r) => [r[1], r[6]])).toEqual([
      ["Store | ", "Yes"],
      ["Store | bread", ""],
      ["Store | eggs", "Yes"],
      ["Store | jam", ""],
      ["Store | tea", ""],
      ["Next", "Yes"],
      ["After", ""],
    ]);
    expect(
      sheet
        .getRange(5, 2, 7, 1)
        .getNotes()
        .map((r: string[]) => r[0]),
    ).toEqual(["", "bread note", "", "", "for Pat", "next note", ""]);
    // One setNotes over exactly the new children's Name cells (rows 7-9, col B)
    expect(noteWrites.filter((w) => w[1] === 2 && w[2] === 3)).toEqual([
      [7, 2, 3, 1],
    ]);
    const t = res.data.transactions;
    expect(t[2]).toMatchObject({ taxDeductible: "Yes", note: "" });
    expect(t[4]).toMatchObject({ taxDeductible: "", note: "for Pat" });
  });

  it("writes Tax Deductible with one setValues over the new children's col G", async () => {
    sheet.seed("A5", [[D, "Costco", 40, "=BALANCE(0)", "", "Credit", ""]]);
    const writes = spySetValues();
    const res = await split(0, [
      item({ expression: "1", taxDeductible: "" }),
      item({ expression: "2", taxDeductible: "Yes" }),
    ]);
    expect(res.ok).toBe(true);
    expect(writes.filter((w) => w[1] === 7)).toEqual([[6, 7, 2, 1]]);
    expect(rows(3).map((r) => r[6])).toEqual(["", "", "Yes"]);
  });

  it("does not touch col G or notes when no item supplies them", async () => {
    sheet.seed("A5", [[D, "Costco", 40, "=BALANCE(0)", "", "Credit", ""]]);
    const writes = spySetValues();
    const notesSpy = vi.spyOn(MockRange.prototype, "setNotes");
    expect((await split(0, [item({ expression: "1" })])).ok).toBe(true);
    expect(writes.filter((w) => w[1] === 7)).toEqual([]);
    expect(notesSpy).not.toHaveBeenCalled();
  });

  it.each([
    ["an invalid taxDeductible", { taxDeductible: "No" }, /taxDeductible/],
    ["a non-string note", { note: 5 }, /note/],
    ["a too-long note", { note: "x".repeat(5001) }, /at most 5000/],
  ])("rejects %s and leaves the sheet unchanged", async (_, fields, msg) => {
    sheet.seed("A5", [[D, "Costco", 40, "=BALANCE(0)", "", "Credit", ""]]);
    const before = snapshot(sheet);
    const res = await split(0, [item({ expression: "1", ...fields })]);
    expect(res.ok).toBe(false);
    expect(res.error.code).toBe("bad_request");
    expect(res.error.message).toMatch(msg);
    expect(snapshot(sheet)).toEqual(before);
  });

  it("writes the parent's item note and envelope atomically with the split", async () => {
    sheet.seed("A5", [
      [D, "Costco", 40, "=BALANCE(0)", "", "Credit", "Yes"],
      [D2, "Next", 7, "=BALANCE(1)", "Gas", "Cash", ""],
    ]);
    const res = await split(
      0,
      [
        item({ itemNote: "milk", expression: "4", envelope: "Food" }),
        item({
          itemNote: "toy",
          expression: "5",
          preTax: false,
          envelope: "Fun",
        }),
      ],
      { parentItemNote: "household", parentEnvelope: "Gas" },
    );
    expect(res.ok).toBe(true);
    expect(rows(4)).toEqual([
      [
        D,
        "Costco | household",
        "=40-R[1]C[0]-R[2]C[0]",
        "=BALANCE(0)",
        "Gas",
        "Credit",
        "Yes",
      ],
      [D, "Costco | milk", "=1.0825*(4)", "=BALANCE(1)", "Food", "Credit", ""],
      [D, "Costco | toy", 5, "=BALANCE(2)", "Fun", "Credit", ""],
      [D2, "Next", 7, "=BALANCE(3)", "Gas", "Cash", ""],
    ]);
    expect(res.data.transactions.map((t: any) => t.splitGroupId)).toEqual([
      "0",
      "0",
      "0",
      null,
    ]);
  });

  it("renames an already-split parent's note and keeps its own prefix for new children", async () => {
    sheet.seed("A5", [
      [D, "Store | old", "=5.5-R[1]C[0]", "=BALANCE(0)", "Food", "Credit", ""],
      [D, "Store | bread", "=1.0825*(2)", "=BALANCE(1)", "Food", "Credit", ""],
    ]);
    const res = await split(0, [item({ itemNote: "eggs", expression: "1" })], {
      parentItemNote: "groceries",
      parentEnvelope: "",
    });
    expect(res.ok).toBe(true);
    expect(rows(3).map((r) => [r[1], r[2], r[4]])).toEqual([
      ["Store | groceries", "=5.5-R[1]C[0]-R[2]C[0]", ""],
      ["Store | bread", "=1.0825*(2)", "Food"],
      ["Store | eggs", "=1.0825*(1)", ""],
    ]);
  });

  it("without the parent fields the parent keeps the native name and envelope", async () => {
    sheet.seed("A5", [[D, "Costco", 40, "=BALANCE(0)", "Food", "Credit", ""]]);
    expect((await split(0, [item()])).ok).toBe(true);
    expect(rows(2).map((r) => [r[1], r[4]])).toEqual([
      ["Costco | ", "Food"],
      ["Costco | ", ""],
    ]);
  });

  it.each([
    [
      { parentItemNote: "=cmd" },
      "parentItemNote must not start with =, +, -, or @",
    ],
    [{ parentItemNote: 5 }, "parentItemNote must be a string"],
    [
      { parentEnvelope: "@x" },
      "parentEnvelope must not start with =, +, -, or @",
    ],
    [
      { parentEnvelope: "Rent" },
      'parentEnvelope "Rent" is not one of the sheet\'s options',
    ],
  ])(
    "rejects bad parent fields %j and writes nothing",
    async (extra, message) => {
      sheet.seed("A5", [[D, "Costco", 40, "=BALANCE(0)", "", "Credit", ""]]);
      const before = snapshot(sheet);
      const res = await split(0, [item()], extra);
      expect(res.error).toEqual({ code: "bad_request", message });
      expect(snapshot(sheet)).toEqual(before);
    },
  );

  it("shifts following rows (cols 0-2 and 4-6 with notes and formats) down by N", async () => {
    sheet.seed("A5", [
      [D, "First", 10, "=BALANCE(0)", "c1", "Credit", "Yes"],
      [D2, "Second", 20, "=BALANCE(1)", "c2", "Cash", "T2"],
      [D2, "Third", 30, "=BALANCE(2)", "c3", "Credit", ""],
    ]);
    sheet.getRange("B6").setNote("second note").setBackground("#E8A9CA");
    const res = await split(0, [item(), item(), item()]);
    expect(res.ok).toBe(true);
    expect(rows(7)).toEqual([
      [
        D,
        "First | ",
        "=10-R[1]C[0]-R[2]C[0]-R[3]C[0]",
        "=BALANCE(0)",
        "c1",
        "Credit",
        "Yes",
      ],
      [D, "First | ", "=1.0825*(0)", "=BALANCE(1)", "", "Credit", ""],
      [D, "First | ", "=1.0825*(0)", "=BALANCE(2)", "", "Credit", ""],
      [D, "First | ", "=1.0825*(0)", "=BALANCE(3)", "", "Credit", ""],
      [D2, "Second", 20, "=BALANCE(4)", "c2", "Cash", "T2"],
      [D2, "Third", 30, "=BALANCE(5)", "c3", "Credit", ""],
      ["", "", "", "=BALANCE(6)", "", "", ""],
    ]);
    expect(sheet.getRange("B9").getNotes()).toEqual([["second note"]]);
    expect(sheet.getRange("B9").getBackgrounds()).toEqual([["#E8A9CA"]]);
    expect(sheet.getRange("B6").getNotes()).toEqual([[""]]);
    expect(sheet.getRange("B6").getBackgrounds()).toEqual([[null]]);
  });

  it("gap: a split on a row above a gap fills the blank rows without shifting; groups never span the gap", async () => {
    sheet.seed("A5", [
      [D, "Store", 10, "=BALANCE(0)", "Food", "Credit", "Yes"],
      ["", "", "", "=BALANCE(1)", "", "", ""],
      [D, "Store", 7, "=BALANCE(2)", "Food", "Credit", "No"],
    ]);
    // Same name and date on both sides of the gap: they are not one group
    const state = await getState();
    expect(state.transactions[1]).toBeNull();
    expect(state.transactions[0].splitGroupId).toBeNull();
    expect(state.transactions[2].splitGroupId).toBeNull();

    const res = await split(0, [item({ itemNote: "milk", expression: "4" })]);
    expect(res.ok).toBe(true);
    expect(rows(3)).toEqual([
      [D, "Store | ", "=10-R[1]C[0]", "=BALANCE(0)", "Food", "Credit", "Yes"],
      [D, "Store | milk", "=1.0825*(4)", "=BALANCE(1)", "", "Credit", ""],
      [D, "Store", 7, "=BALANCE(2)", "Food", "Credit", "No"],
    ]);
    // Once the child fills the gap, the later "Store" row is adjacent and has the same name prefix,
    // so serialize now groups it with the split (current behavior)
    expect(res.data.transactions.map((r: any) => r && r.splitGroupId)).toEqual([
      "0",
      "0",
      "0",
    ]);
  });

  it("shifts when only a later one of the N target rows has content", async () => {
    sheet.seed("A5", [[D, "First", 10, "=BALANCE(0)", "", "Credit", ""]]);
    sheet.seed("A7", [[D2, "Later", 3, "=BALANCE(2)", "", "Cash", ""]]);
    const res = await split(0, [item(), item()]);
    expect(res.ok).toBe(true);
    expect(rows(5).map((r) => r.slice(0, 3))).toEqual([
      [D, "First | ", "=10-R[1]C[0]-R[2]C[0]"],
      [D, "First | ", "=1.0825*(0)"],
      [D, "First | ", "=1.0825*(0)"],
      ["", "", ""],
      [D2, "Later", 3],
    ]);
  });

  it("rejects a split with no room", async () => {
    sheet.seed(
      "A5",
      Array.from({ length: 8 }, (_, i) => [
        D2,
        `${i} Shop`,
        1,
        "=B",
        "",
        "Credit",
        "",
      ]),
    );
    // 8 rows used, 2 free: 2 new rows fit at the end, 3 do not
    const res = await split(7, [item(), item(), item()]);
    expect(res.error).toEqual({
      code: "bad_request",
      message: `No room on sheet ${SHEET} to split group of 1 starting at row 12!`,
    });
    // Shifting needs N free rows at the end
    const res2 = await split(0, [item(), item(), item()]);
    expect(res2.error.code).toBe("bad_request");
    expect(res2.error.message).toMatch(/not enough empty transaction rows/);
    expect(sheet.getRange("B5").getValue()).toBe("0 Shop");
    expect((await split(7, [item(), item()])).ok).toBe(true);
  });

  it("rejects splitting a row in the middle of a group and an empty row", async () => {
    sheet.seed("A5", [
      [D, "Store | ", "=5-R[1]C[0]", "=B", "", "Credit", ""],
      [D, "Store | ", "=1.0825*(0)", "=B", "", "Credit", ""],
    ]);
    expect((await split(1, [item()])).error).toEqual({
      code: "bad_request",
      message: `Transaction in sheet ${SHEET} at row 6 seems to be in a transaction group already!`,
    });
    expect(
      (
        await api("splitTransaction", {
          sheetName: SHEET,
          rowIndex: 5,
          baseFingerprint: "none",
          splits: [item()],
          force: true,
        })
      ).error,
    ).toEqual({
      code: "bad_request",
      message: "Row 5 is empty",
    });
  });

  it.each([
    [[], "splits must be an array of 1 to 50 items"],
    ["x", "splits must be an array of 1 to 50 items"],
    [
      Array.from({ length: 51 }, () => item()),
      "splits must be an array of 1 to 50 items",
    ],
    [[5], "splits[0] must be an object"],
    [
      [item({ itemNote: "=cmd|' /C calc'!A0" })],
      "splits[0].itemNote must not start with =, +, -, or @",
    ],
    [
      [item(), item({ itemNote: "@SUM(1)" })],
      "splits[1].itemNote must not start with =, +, -, or @",
    ],
    [
      [item({ envelope: '=HYPERLINK("x")' })],
      "splits[0].envelope must not start with =, +, -, or @",
    ],
    [
      [item({ envelope: "Rent" })],
      'splits[0].envelope "Rent" is not one of the sheet\'s options',
    ],
    [[item({ envelope: undefined })], "splits[0].envelope must be a string"],
    [
      [item({ paidWith: "+1" })],
      "splits[0].paidWith must not start with =, +, -, or @",
    ],
    [
      [item({ paidWith: "Gold" })],
      'splits[0].paidWith "Gold" is not one of the sheet\'s options',
    ],
    [
      [item({ expression: "1+" })],
      "splits[0].expression is not a valid expression",
    ],
    [
      [item({ expression: "=IMPORTXML(1)" })],
      "splits[0].expression is not a valid expression",
    ],
    [[item({ expression: 5 })], "splits[0].expression must be a string"],
    [[item({ preTax: undefined })], "splits[0].preTax must be a boolean"],
  ])("rejects splits %j with bad_request", async (splits, message) => {
    sheet.seed("A5", [[D, "Store", 5.5, "=B", "", "Credit", ""]]);
    const before = snapshot(sheet);
    const res = await split(0, splits);
    expect(res).toEqual({ ok: false, error: { code: "bad_request", message } });
    expect(snapshot(sheet)).toEqual(before);
  });

  it("returns conflict without splitting when the row changed", async () => {
    sheet.seed("A5", [[D, "Store", 5.5, "=B", "", "Credit", ""]]);
    const staleFp = await fingerprintOf(0);
    sheet.getRange("C5").setValue(6);
    const res = await api("splitTransaction", {
      sheetName: SHEET,
      rowIndex: 0,
      baseFingerprint: staleFp,
      splits: [item()],
    });
    expect(res.error.code).toBe("conflict");
    expect(res.conflict.current).toMatchObject({ rowIndex: 0, cost: 6 });
    expect(sheet.getRange("B5:C6").getValues()).toEqual([
      ["Store", 6],
      ["", ""],
    ]);
  });

  it("returns busy when the script lock is held elsewhere", async () => {
    sheet.seed("A5", [[D, "Store", 5.5, "=B", "", "Credit", ""]]);
    const fp = await fingerprintOf(0);
    env.LockService.scriptLock.contended = true;
    const res = await api("splitTransaction", {
      sheetName: SHEET,
      rowIndex: 0,
      baseFingerprint: fp,
      splits: [item()],
    });
    expect(res.error.code).toBe("busy");
    expect(sheet.getRange("B5").getValue()).toBe("Store");
  });
});

describe("deleteTransaction", () => {
  /** Rows A:G, with col D the sheet's =BALANCE(i) formula seeded by addApiSheet */
  const seed = () => {
    sheet.seed("A5", [
      [D, "Coffee", 4.5, "=BALANCE(0)", "Food", "Credit", ""],
      [
        D,
        "Store | ",
        "=87.5-R[1]C[0]-R[2]C[0]",
        "=BALANCE(1)",
        "Food",
        "Credit",
        "Yes",
      ],
      [D, "Store | milk", "=1.0825*(4)", "=BALANCE(2)", "Food", "Credit", ""],
      [D, "Store | toy", "=10", "=BALANCE(3)", "Fun", "Cash", "Yes"],
      [D2, "Gas | ", "=30-R[1]C[0]", "=BALANCE(4)", "Gas", "Credit", ""],
      [D2, "Gas | snack", "=1.0825*(2)", "=BALANCE(5)", "Food", "Credit", ""],
      [D2, "Last", 9, "=BALANCE(6)", "", "Sam Venmo", "Yes"],
    ]);
    sheet.getRange("B9").setNote("gas note").setBackground("#E8A9CA");
    sheet.getRange("C10").setBackground("#FF0000");
    sheet.getRange("B10").setNote("snack note");
    // Split checkboxes (col J) and anything outside cols A-G must never move
    sheet.getRange("J5:J11").insertCheckboxes();
    sheet.getRange("J6").setValue(true);
  };

  const del = async (
    rowIndex: number,
    count: number,
    extra: Record<string, unknown> = {},
  ) => {
    const state = await getState();
    return api("deleteTransaction", {
      sheetName: SHEET,
      rowIndex,
      rows: Array.from({ length: count }, (_, i) => ({
        rowIndex: rowIndex + i,
        baseFingerprint: state.transactions[rowIndex + i]?.fingerprint ?? "x",
      })),
      ...extra,
    });
  };

  const rows = (n: number) => sheet.getRange(5, 1, n, 7).getValues();
  const balanceColumn = () =>
    sheet
      .getRange("D5:D14")
      .getFormulas()
      .map(([f]: string[]) => f);
  const originalBalance = Array.from(
    { length: 10 },
    (_, i) => `=BALANCE(${i})`,
  );

  /** The seeded rows after clearing `blank` row indexes in cols A-C and E-G */
  const seeded = [
    [D, "Coffee", 4.5, "=BALANCE(0)", "Food", "Credit", ""],
    [
      D,
      "Store | ",
      "=87.5-R[1]C[0]-R[2]C[0]",
      "=BALANCE(1)",
      "Food",
      "Credit",
      "Yes",
    ],
    [D, "Store | milk", "=1.0825*(4)", "=BALANCE(2)", "Food", "Credit", ""],
    [D, "Store | toy", "=10", "=BALANCE(3)", "Fun", "Cash", "Yes"],
    [D2, "Gas | ", "=30-R[1]C[0]", "=BALANCE(4)", "Gas", "Credit", ""],
    [D2, "Gas | snack", "=1.0825*(2)", "=BALANCE(5)", "Food", "Credit", ""],
    [D2, "Last", 9, "=BALANCE(6)", "", "Sam Venmo", "Yes"],
  ];
  const blanked = (blank: number[]) =>
    seeded.map((row, i) =>
      blank.includes(i) ? ["", "", "", row[3], "", "", ""] : row,
    );

  it("deletes a single row in place: the row is blank, nothing else moves or is rewritten", async () => {
    seed();
    const writes = spySetValues();
    const res = await del(0, 1);
    expect(res.ok).toBe(true);
    expect(rows(7)).toEqual(blanked([0]));
    // Rows below keep their formulas, notes, and backgrounds in place
    expect(sheet.getRange("B9").getNotes()).toEqual([["gas note"]]);
    expect(sheet.getRange("B9").getBackgrounds()).toEqual([["#E8A9CA"]]);
    expect(sheet.getRange("B10").getNotes()).toEqual([["snack note"]]);
    expect(sheet.getRange("C10").getBackgrounds()).toEqual([["#FF0000"]]);
    // Col D and the checkbox column are untouched
    expect(balanceColumn()).toEqual(originalBalance);
    expect(sheet.getRange("J5:J7").getValues()).toEqual([
      [false],
      [true],
      [false],
    ]);
    // One batch write per cleared column block (A:C and E:G), for exactly the deleted row
    expect(writes).toEqual([
      [5, 1, 1, 3],
      [5, 5, 1, 3],
    ]);
    // The refreshed state is returned: the gap is null and the other rows keep their indexes
    expect(res.data.sheetName).toBe(SHEET);
    expect(res.data.transactions[0]).toBeNull();
    expect(res.data.transactions[1]).toMatchObject({
      rowIndex: 1,
      name: "Store | ",
      splitGroupId: "1",
    });
  });

  it("clears the notes and backgrounds of the deleted cells only", async () => {
    seed();
    expect((await del(4, 2)).ok).toBe(true);
    expect(sheet.getRange("B9:B10").getNotes()).toEqual([[""], [""]]);
    expect(sheet.getRange("B9").getBackgrounds()).toEqual([[null]]);
    expect(sheet.getRange("C10").getBackgrounds()).toEqual([[null]]);
  });

  it("deletes a whole split group in place and leaves another split group's formulas untouched", async () => {
    seed();
    const res = await del(1, 3);
    expect(res.ok).toBe(true);
    expect(rows(7)).toEqual(blanked([1, 2, 3]));
    expect(sheet.getRange("B9").getNotes()).toEqual([["gas note"]]);
    expect(sheet.getRange("C10").getBackgrounds()).toEqual([["#FF0000"]]);
    expect(balanceColumn()).toEqual(originalBalance);
    expect(
      res.data.transactions.map((r: any) => (r ? r.splitGroupId : "gap")),
    ).toEqual([null, "gap", "gap", "gap", "4", "4", null]);
  });

  it("keeps the other split group evaluating correctly after a delete", async () => {
    env.spreadsheet.evaluateFormulas = true;
    seed();
    expect((await del(0, 1)).ok).toBe(true);
    const state = await getState();
    expect(state.transactions[1].cost).toBeCloseTo(87.5 - 1.0825 * 4 - 10);
    expect(state.transactions[4].cost).toBeCloseTo(30 - 1.0825 * 2);
  });

  it("deletes the last row of a full table", async () => {
    sheet.seed(
      "A5",
      // Alternating dates so no rows form a split group
      Array.from({ length: 10 }, (_, i) => [
        i % 2 ? D2 : D,
        `Row ${i}`,
        i,
        `=BALANCE(${i})`,
        "",
        "",
        "",
      ]),
    );
    expect((await del(9, 1)).ok).toBe(true);
    expect(sheet.getRange("A14:C14").getValues()).toEqual([["", "", ""]]);
    expect(sheet.getRange("B13").getValue()).toBe("Row 8");
    expect(sheet.getRange("B5").getValue()).toBe("Row 0");
    expect(balanceColumn()).toEqual(originalBalance);
  });

  it("returns a conflict with the current row when any row changed, and deletes nothing", async () => {
    seed();
    const state = await getState();
    sheet.getRange("C7").setValue("=1.0825*(5)");
    const before = snapshot(sheet);
    const request = {
      sheetName: SHEET,
      rowIndex: 1,
      rows: [1, 2, 3].map((r) => ({
        rowIndex: r,
        baseFingerprint: state.transactions[r].fingerprint,
      })),
    };
    const res = await api("deleteTransaction", request);
    expect(res.error.code).toBe("conflict");
    expect(res.conflict.current).toMatchObject({
      rowIndex: 2,
      costFormula: "=1.0825*(5)",
    });
    expect(snapshot(sheet)).toEqual(before);

    // force deletes anyway
    const forced = await api("deleteTransaction", { ...request, force: true });
    expect(forced.ok).toBe(true);
    expect(sheet.getRange("B6:B8").getValues()).toEqual([[""], [""], [""]]);
    expect(sheet.getRange("B9").getValue()).toBe("Gas | ");
  });

  it.each([
    ["only part of a split group", 1, 2],
    ["a child row of a split group", 2, 1],
    ["the tail of a split group", 2, 2],
    ["a split group plus the next row", 4, 3],
    ["a single row plus the next group", 0, 2],
  ])("rejects %s", async (_label, rowIndex, count) => {
    seed();
    const before = snapshot(sheet);
    const res = await del(rowIndex, count);
    expect(res.error.code).toBe("bad_request");
    expect(snapshot(sheet)).toEqual(before);
  });

  it("rejects an empty row", async () => {
    seed();
    // force skips the fingerprint check (an empty row is not in the pulled list)
    const res = await del(8, 1, { force: true });
    expect(res.error.code).toBe("bad_request");
    expect(res.error.message).toBe("Row 8 is empty");
  });

  it.each([
    ["missing rows", { rows: undefined }],
    ["empty rows", { rows: [] }],
    [
      "non-contiguous rows",
      {
        rows: [
          { rowIndex: 1, baseFingerprint: "a" },
          { rowIndex: 3, baseFingerprint: "b" },
        ],
      },
    ],
    [
      "rows not starting at rowIndex",
      { rows: [{ rowIndex: 2, baseFingerprint: "a" }] },
    ],
    ["empty fingerprint", { rows: [{ rowIndex: 1, baseFingerprint: "" }] }],
    [
      "rows past the table end",
      {
        rowIndex: 9,
        rows: [
          { rowIndex: 9, baseFingerprint: "a" },
          { rowIndex: 10, baseFingerprint: "b" },
        ],
      },
    ],
    [
      "a bad rowIndex",
      { rowIndex: 10, rows: [{ rowIndex: 10, baseFingerprint: "a" }] },
    ],
  ])("validates the request: %s", async (_label, fields) => {
    seed();
    const before = snapshot(sheet);
    const res = await api("deleteTransaction", {
      sheetName: SHEET,
      rowIndex: 1,
      ...fields,
    });
    expect(res.error.code).toBe("bad_request");
    expect(snapshot(sheet)).toEqual(before);
  });

  it("Budgeting.deleteTransactionRows rejects out-of-range rows", async () => {
    const { Budgeting } = await import("../src/util/budgeting.util");
    for (const [rowIndex, count] of [
      [-1, 1],
      [0, 0],
      [9, 2],
      [0.5, 1],
    ])
      expect(() =>
        Budgeting.deleteTransactionRows(sheet, rowIndex, count),
      ).toThrow("out of range");
  });

  it("returns not_found for an unknown sheet and busy when the lock is held", async () => {
    seed();
    const missing = await api("deleteTransaction", {
      sheetName: "Nope",
      rowIndex: 0,
      rows: [{ rowIndex: 0, baseFingerprint: "a" }],
    });
    expect(missing.error.code).toBe("not_found");
    const fp = await fingerprintOf(0);
    env.LockService.scriptLock.contended = true;
    const res = await api("deleteTransaction", {
      sheetName: SHEET,
      rowIndex: 0,
      rows: [{ rowIndex: 0, baseFingerprint: fp }],
    });
    expect(res.error.code).toBe("busy");
    expect(sheet.getRange("B5").getValue()).toBe("Coffee");
  });
});

describe("syncReceipts", () => {
  const chase = chasePurchaseEmail;

  it("imports up to 25 threads, labels them, and returns errors without using the UI", async () => {
    env.SpreadsheetApp.uiAvailable = false;
    vi.useFakeTimers();
    vi.setSystemTime(new Date("2024-02-25T18:00:00Z"));
    const good = makeThread(chase("10.50", "Shop", "2024-02-20T12:00:00Z"));
    const bad = makeThread({
      from: "Weird <weird at example.org>",
      subject: "huh",
      body: "",
      date: new Date("2024-02-21T12:00:00Z"),
    });
    const ignored = makeThread({
      from: "Chase <no.reply.alerts@chase.com>",
      subject: "Statement ready",
      body: "hi",
    });
    env.GmailApp.threads = [good, bad, ignored];

    const res = await api("syncReceipts");
    expect(res.ok).toBe(true);
    expect(env.GmailApp.searches).toEqual([
      expect.objectContaining({ start: 0, max: 25 }),
    ]);
    expect(res.data.receiptInfosAdded).toBe(2);
    expect(res.data.receiptInfosAddedBySheet).toEqual({ [SHEET]: 2 });
    expect(res.data.numErrors).toBe(1);
    expect(res.data.errors[0]).toContain(
      "Could not get email address from Weird <weird at example.org>",
    );
    expect(env.SpreadsheetApp.ui.alerts).toEqual([]);
    expect(good.archived).toBe(true);
    expect(bad.labels.map((l: any) => l.getName())).toEqual([
      "Receipts/Scripted",
      "Receipts/Error",
    ]);
    expect(ignored.labels.map((l: any) => l.getName())).toEqual([
      "Receipts/Scripted",
      "Receipts/Ignored",
    ]);
    // Includes the refreshed state of the current sheet
    expect(res.data.state.sheetName).toBe(SHEET);
    expect(res.data.state.transactions[0]).toMatchObject({
      name: "Shop",
      cost: 10.5,
      paidWith: "Credit",
    });
    expect(res.data.state.transactions[1]).toMatchObject({ isError: true });
    expect(env.LockService.scriptLock.calls).toEqual([
      "tryLock:20000",
      "release",
    ]);
  });

  it("returns bad_request when a receipt fits no sheet, without creating sheets or labeling", async () => {
    const t = makeThread(chase("1.00", "Shop", "2030-01-01T12:00:00Z"));
    env.GmailApp.threads = [t];
    const sheetCount = env.spreadsheet.getSheets().length;
    const res = await api("syncReceipts");
    expect(res.error.code).toBe("bad_request");
    expect(res.error.message).toMatch(
      /^Could not find transaction sheet for receipt/,
    );
    expect(env.spreadsheet.getSheets()).toHaveLength(sheetCount);
    expect(t.labels).toEqual([]);
    expect(env.LockService.scriptLock.held).toBe(false);
  });

  it("mark:false imports rows but does not label or archive; omitted mark still marks", async () => {
    vi.useFakeTimers();
    vi.setSystemTime(new Date("2024-02-25T18:00:00Z"));
    const a = makeThread(chase("10.50", "Shop", "2024-02-20T12:00:00Z"));
    env.GmailApp.threads = [a];
    const res = await api("syncReceipts", { mark: false });
    expect(res.ok).toBe(true);
    expect(res.data.receiptInfosAdded).toBe(1);
    expect(a.labels).toEqual([]);
    expect(a.archived).toBe(false);
    const res2 = await api("syncReceipts");
    expect(res2.data.receiptInfosAdded).toBe(1);
    expect(a.archived).toBe(true);
    expect(a.labels.map((l: any) => l.getName())).toContain(
      "Receipts/Scripted",
    );
  });

  it("rejects a non-boolean mark", async () => {
    const res = await api("syncReceipts", { mark: "no" });
    expect(res.error.code).toBe("bad_request");
    expect(env.GmailApp.searches).toEqual([]);
  });

  it("returns busy without searching Gmail when the lock is held elsewhere", async () => {
    env.LockService.scriptLock.contended = true;
    const res = await api("syncReceipts");
    expect(res.error.code).toBe("busy");
    expect(env.GmailApp.searches).toEqual([]);
  });

  it("reports Gmail failures as server_error and releases the lock", async () => {
    env.GmailApp.search = () => {
      throw new Error("Gmail down");
    };
    const res = await api("syncReceipts");
    expect(res).toEqual({
      ok: false,
      error: { code: "server_error", message: "Gmail down" },
    });
    expect(env.LockService.scriptLock.calls).toEqual([
      "tryLock:20000",
      "release",
    ]);
  });
});

describe("addTransaction", () => {
  const REQUEST_ID = "req-0001-abcdefgh";
  const FIELDS = {
    date: "2024-02-22",
    name: "Farmers market",
    costExpression: "4.25*2",
    envelope: "Food",
    paidWith: "Cash",
    taxDeductible: "Yes",
    note: "cash only",
  };
  const add = (
    fields: Record<string, unknown> = FIELDS,
    extra: Record<string, unknown> = {},
  ) =>
    api("addTransaction", {
      sheetName: SHEET,
      requestId: REQUEST_ID,
      fields,
      ...extra,
    });

  /** Rows 0-1 and 3 have content; row 2 is an interior gap */
  const seed = () => {
    sheet.seed("A5", [
      [D, "Coffee", 4.5, "=BALANCE(0)", "Food", "Credit", ""],
      [D, "Gas", 30, "=BALANCE(1)", "Gas", "Credit", ""],
      ["", "", "", "=BALANCE(2)", "", "", ""],
      [D2, "Last", 9, "=BALANCE(3)", "", "Sam Venmo", ""],
    ]);
    sheet.getRange("J5:J14").insertCheckboxes();
    sheet.getRange("J9").setValue(true);
  };
  const balanceColumn = () =>
    sheet
      .getRange("D5:D14")
      .getFormulas()
      .map(([f]: string[]) => f);

  it("writes the row after the last non-empty row (never into a gap) and returns the state", async () => {
    seed();
    const checkboxes = sheet.getRange("J5:J14").getValues();
    const writes = spySetValues();
    const res = await add();
    expect(res.ok).toBe(true);
    expect(res.data.sheetName).toBe(SHEET);
    expect(res.data.addedRowIndex).toBe(4);
    // The interior gap stays empty
    expect(res.data.transactions[2]).toBeNull();
    expect(res.data.transactions[4]).toMatchObject({
      rowIndex: 4,
      date: "2024-02-22",
      name: "Farmers market",
      // The mock does not evaluate formulas, so only the formula is checked
      costFormula: "=4.25*2",
      envelope: "Food",
      paidWith: "Cash",
      taxDeductible: "Yes",
      note: "cash only",
    });
    const date = sheet.getRange("A9").getValue() as Date;
    expect(date).toBeInstanceOf(Date);
    // Midnight in the spreadsheet time zone (America/Chicago, UTC-6 in February)
    expect(date.toISOString()).toBe("2024-02-22T06:00:00.000Z");
    expect(sheet.getRange("B7:C7").getValues()).toEqual([["", ""]]);
    // Two batch writes (cols A-C and E-G of the one row); col D and the checkboxes untouched
    expect(writes).toEqual([
      [9, 1, 1, 3],
      [9, 5, 1, 3],
    ]);
    expect(balanceColumn()).toEqual(
      Array.from({ length: 10 }, (_, i) => `=BALANCE(${i})`),
    );
    expect(sheet.getRange("J5:J14").getValues()).toEqual(checkboxes);
    expect(env.LockService.scriptLock.calls).toEqual([
      "tryLock:20000",
      "release",
    ]);
  });

  it("writes the first row of an empty table, a plain number cost, and blank optional fields", async () => {
    const res = await add({
      date: "2024-02-20",
      name: "Cash",
      costExpression: "12",
    });
    expect(res.ok).toBe(true);
    expect(res.data.addedRowIndex).toBe(0);
    expect(sheet.getRange("A5:G5").getValues()[0].slice(1)).toEqual([
      "Cash",
      12,
      "=BALANCE(0)",
      "",
      "",
      "",
    ]);
    expect(sheet.getRange("B5").getNotes()).toEqual([[""]]);
  });

  it("does not insert twice for a repeated requestId and remembers it for 6 hours", async () => {
    seed();
    const first = await add();
    expect(first.ok).toBe(true);
    const before = snapshot(sheet);
    const again = await add();
    expect(again.ok).toBe(true);
    expect(again.data.addedRowIndex).toBe(4);
    expect(again.data.transactions).toHaveLength(5);
    expect(snapshot(sheet)).toEqual(before);
    expect(
      env.CacheService.scriptCache.expirations.get(
        `addTransaction:${REQUEST_ID}`,
      ),
    ).toBe(21600);
    // A different requestId inserts a new row
    const other = await add(FIELDS, { requestId: "req-0002-abcdefgh" });
    expect(other.data.addedRowIndex).toBe(5);
  });

  it("returns bad_request when the last row is used, even with gaps above", async () => {
    seed();
    sheet.seed("A14", [[D2, "Bottom", 1]]);
    const before = snapshot(sheet);
    const res = await add();
    expect(res).toMatchObject({ ok: false, error: { code: "bad_request" } });
    expect(res.error.message).toMatch(/no empty transaction row/i);
    expect(snapshot(sheet)).toEqual(before);
    expect(
      env.CacheService.scriptCache.get(`addTransaction:${REQUEST_ID}`),
    ).toBeNull();
  });

  it.each([
    ["missing fields", { fields: undefined }, /fields must be an object/],
    ["unknown field", { fields: { ...FIELDS, extra: 1 } }, /Unknown fields/],
    ["bad requestId", { requestId: "a b" }, /requestId/],
    ["missing requestId", { requestId: undefined }, /requestId/],
    ["empty name", { fields: { ...FIELDS, name: "  " } }, /fields.name/],
    [
      "formula name",
      { fields: { ...FIELDS, name: "=HYPERLINK(1)" } },
      /fields.name must not start/,
    ],
    [
      "bad cost",
      { fields: { ...FIELDS, costExpression: "SUM(A1)" } },
      /fields.costExpression/,
    ],
    [
      "empty cost",
      { fields: { ...FIELDS, costExpression: "" } },
      /fields.costExpression/,
    ],
    [
      "bad date format",
      { fields: { ...FIELDS, date: "2/22/2024" } },
      /fields.date/,
    ],
    [
      "impossible date",
      { fields: { ...FIELDS, date: "2024-02-30" } },
      /fields.date/,
    ],
    [
      "unknown envelope",
      { fields: { ...FIELDS, envelope: "Rent" } },
      /fields.envelope "Rent"/,
    ],
    [
      "unknown paid with",
      { fields: { ...FIELDS, paidWith: "Gold" } },
      /fields.paidWith "Gold"/,
    ],
    [
      "bad taxDeductible",
      { fields: { ...FIELDS, taxDeductible: "No" } },
      /fields.taxDeductible/,
    ],
    ["note not a string", { fields: { ...FIELDS, note: 3 } }, /fields.note/],
  ])("rejects %s without writing", async (_label, extra, message) => {
    seed();
    const before = snapshot(sheet);
    const res = await add(FIELDS, extra);
    expect(res).toMatchObject({ ok: false, error: { code: "bad_request" } });
    expect(res.error.message).toMatch(message);
    expect(snapshot(sheet)).toEqual(before);
  });

  it("returns not_found for an unknown sheet and busy when the lock is taken", async () => {
    expect(await add(FIELDS, { sheetName: "Nope" })).toMatchObject({
      ok: false,
      error: { code: "not_found" },
    });
    env.LockService.scriptLock.contended = true;
    expect(await add()).toMatchObject({ ok: false, error: { code: "busy" } });
    expect(sheet.getRange("B5").getValue()).toBe("");
  });

  it("rejects an envelope on a sheet without option lists", async () => {
    const res = await add(FIELDS, { sheetName: OLDER_SHEET });
    expect(res).toMatchObject({ ok: false, error: { code: "bad_request" } });
    expect(res.error.message).toMatch(/no EnvelopesStart/);
  });
});

describe("updateSplitGroup", () => {
  /** A split group at rows 1-3 between two plain rows */
  function seedGroup() {
    sheet.seed("A5", [
      [D, "Cafe", 3, "=BALANCE(0)", "", "Credit", ""],
      [
        D,
        "Store | ",
        "=45-R[1]C[0]-R[2]C[0]",
        "=BALANCE(1)",
        "Food",
        "Credit",
        "",
      ],
      [D, "Store | milk", 2, "=BALANCE(2)", "Food", "Credit", ""],
      [D, "Store | eggs", "=1.0825*(1)", "=BALANCE(3)", "", "Credit", "Yes"],
      [D2, "Next", 7, "=BALANCE(4)", "Gas", "Cash", ""],
    ]);
  }

  /** Every listed row (default: the group, rows 1-3) with its current fingerprint and note */
  async function groupRows(
    changes: Record<number, Record<string, unknown>> = {},
    indexes = [1, 2, 3],
  ) {
    const state = await getState();
    return indexes.map((r) => ({
      rowIndex: r,
      baseFingerprint: state.transactions[r].fingerprint,
      baseNote: state.transactions[r].note,
      ...(changes[r] ? { changes: changes[r] } : {}),
    }));
  }

  const updateGroup = (fields: Record<string, unknown>) =>
    api("updateSplitGroup", { sheetName: SHEET, ...fields });

  it("writes every row's edits, the title, and the total atomically and returns only the group's rows", async () => {
    seedGroup();
    const rows = await groupRows({
      1: { name: "Store | groceries", note: "parent note" },
      2: { costExpression: "3", envelope: "Gas", name: "Store | whole milk" },
      3: { costExpression: "2", costPreTax: true, taxDeductible: "" },
    });
    const res = await updateGroup({ rows, title: "Market", splitTotal: "50" });
    expect(res.ok).toBe(true);
    expect(
      sheet
        .getRange("A5:G9")
        .getValues()
        .map((r) => r.slice(1)),
    ).toEqual([
      ["Cafe", 3, "=BALANCE(0)", "", "Credit", ""],
      [
        "Market | groceries",
        "=50-R[1]C[0]-R[2]C[0]",
        "=BALANCE(1)",
        "Food",
        "Credit",
        "",
      ],
      ["Market | whole milk", 3, "=BALANCE(2)", "Gas", "Credit", ""],
      ["Market | eggs", "=1.0825*(2)", "=BALANCE(3)", "", "Credit", ""],
      ["Next", 7, "=BALANCE(4)", "Gas", "Cash", ""],
    ]);
    expect(sheet.getRange("B6:B8").getNotes()).toEqual([
      ["parent note"],
      [""],
      [""],
    ]);
    // Only the group's rows come back (not the whole pay-period state)
    expect(res.data.state).toBeUndefined();
    expect(res.data.rows.map((r: any) => [r.rowIndex, r.name])).toEqual([
      [1, "Market | groceries"],
      [2, "Market | whole milk"],
      [3, "Market | eggs"],
    ]);
    const state = await getState();
    expect(res.data.rows).toEqual(state.transactions.slice(1, 4));
  });

  it("renames children that keep their own item note, and writes each column with one batch call", async () => {
    seedGroup();
    const rows = await groupRows({ 2: { envelope: "Gas" } });
    const writes = spySetValues();
    const notes = vi.spyOn(MockRange.prototype, "setNotes");
    const res = await updateGroup({ rows, title: "Market" });
    expect(res.ok).toBe(true);
    expect(sheet.getRange("B6:B8").getValues()).toEqual([
      ["Market | "],
      ["Market | milk"],
      ["Market | eggs"],
    ]);
    // One write of the group's names (B6:B8) and one of its metadata (E6:G8); no cost or note write
    expect(writes).toEqual([
      [6, 2, 3, 1],
      [6, 5, 3, 3],
    ]);
    expect(notes).not.toHaveBeenCalled();
    expect(sheet.getRange("E6:G8").getValues()).toEqual([
      ["Food", "Credit", ""],
      ["Gas", "Credit", ""],
      ["", "Credit", "Yes"],
    ]);
  });

  it("returns conflict with the group's current rows and writes nothing when any row changed", async () => {
    seedGroup();
    const rows = await groupRows({ 1: { envelope: "Gas" } });
    // Someone edits the last child's cost on the sheet
    sheet.getRange("C8").setValue("=1.0825*(4)");
    const before = snapshot(sheet);
    const res = await updateGroup({ rows, title: "Market" });
    expect(res.ok).toBe(false);
    expect(res.error.code).toBe("conflict");
    expect(res.conflict.current).toMatchObject({
      rowIndex: 3,
      costFormula: "=1.0825*(4)",
    });
    expect(res.conflict.rows.map((r: any) => r.rowIndex)).toEqual([1, 2, 3]);
    expect(res.conflict.rows[2].costFormula).toBe("=1.0825*(4)");
    expect(snapshot(sheet)).toEqual(before);
    expect(env.LockService.scriptLock.held).toBe(false);

    // "Overwrite with Mine": resend with every row's current fingerprint
    const fresh = await groupRows({ 1: { envelope: "Gas" } });
    const retry = await updateGroup({ rows: fresh, title: "Market" });
    expect(retry.ok).toBe(true);
    expect(sheet.getRange("B6:C8").getValues()).toEqual([
      ["Market | ", "=45-R[1]C[0]-R[2]C[0]"],
      ["Market | milk", 2],
      ["Market | eggs", "=1.0825*(4)"],
    ]);
    expect(sheet.getRange("E6").getValue()).toBe("Gas");
  });

  it("returns conflict when a row's note changed on the sheet", async () => {
    seedGroup();
    const rows = await groupRows({ 2: { note: "mine" } });
    sheet.getRange("B8").setNote("theirs");
    const res = await updateGroup({ rows });
    expect(res.error.code).toBe("conflict");
    expect(res.conflict.current).toMatchObject({ rowIndex: 3, note: "theirs" });
    expect(sheet.getRange("B7").getNotes()).toEqual([[""]]);
  });

  it("returns conflict when the group gained or lost rows since the client loaded it", async () => {
    seedGroup();
    const rows = await groupRows({ 2: { envelope: "Gas" } });
    // The row below now joins the group (same date and prefix); no listed row changed
    sheet.seed("A9", [[D, "Store | jam", 1, "=BALANCE(4)", "", "Cash", ""]]);
    const before = snapshot(sheet);
    const res = await updateGroup({ rows });
    expect(res.error.code).toBe("conflict");
    expect(res.conflict.rows.map((r: any) => r.rowIndex)).toEqual([1, 2, 3, 4]);
    expect(snapshot(sheet)).toEqual(before);
    // Listing only part of the group is a conflict too
    seedGroup();
    const part = await groupRows({ 2: { envelope: "Gas" } }, [1, 2]);
    expect((await updateGroup({ rows: part })).error.code).toBe("conflict");
    expect(sheet.getRange("E7").getValue()).toBe("Food");
  });

  it("adds new items with the native split in the same request and returns the state", async () => {
    seedGroup();
    const rows = await groupRows({ 2: { costExpression: "3" } });
    const res = await updateGroup({
      rows,
      title: "Market",
      newItems: [
        {
          itemNote: "jam",
          expression: "4",
          preTax: false,
          envelope: "Fun",
          note: "jam note",
        },
      ],
    });
    expect(res.ok).toBe(true);
    expect(sheet.getRange("B6:C10").getValues()).toEqual([
      ["Market | ", "=45-R[1]C[0]-R[2]C[0]-R[3]C[0]"],
      ["Market | milk", 3],
      ["Market | eggs", "=1.0825*(1)"],
      ["Market | jam", 4],
      ["Next", 7],
    ]);
    expect(sheet.getRange("E9").getValue()).toBe("Fun");
    expect(sheet.getRange("B9").getNotes()).toEqual([["jam note"]]);
    // Rows below moved, so the whole state comes back with the group's rows
    expect(res.data.rows.map((r: any) => r.rowIndex)).toEqual([1, 2, 3, 4]);
    expect(res.data.state.transactions[5]).toMatchObject({ name: "Next" });
  });

  it("validates the request before taking the lock", async () => {
    seedGroup();
    const rows = await groupRows();
    const before = snapshot(sheet);
    const withChange = (i: number, changes: Record<string, unknown>) =>
      rows.map((r, j) => (j === i ? { ...r, changes } : r));
    const cases: [Record<string, unknown>, string][] = [
      [{ rows: rows.slice(0, 1), title: "X" }, "rows must be an array of 2 to"],
      [
        { rows: [rows[1], rows[0], rows[2]], title: "X" },
        "rows must list contiguous",
      ],
      [{ rows, title: "=1+1" }, "title must not start with"],
      [{ rows, title: "A | B" }, "title must not contain |"],
      [{ rows, title: " " }, "title must not be empty"],
      [{ rows, splitTotal: "1+" }, "splitTotal must be a number or valid math"],
      [
        { rows: withChange(0, { costExpression: "5" }) },
        "Cannot edit the cost of a split parent row",
      ],
      [
        { rows: withChange(1, { costExpression: "1+" }) },
        "rows[1].changes.costExpression is not a valid expression",
      ],
      [
        { rows: withChange(1, { envelope: "Nope" }) },
        'rows[1].changes.envelope "Nope" is not one of',
      ],
      [
        { rows: withChange(1, { splitTotal: "5" }) },
        "Unknown rows[1].changes: splitTotal",
      ],
      [
        { rows: rows.map((r) => ({ ...r, baseNote: 5 })), title: "X" },
        "baseNote",
      ],
      [{ rows }, "No changes given"],
    ];
    for (const [fields, message] of cases) {
      const res = await updateGroup(fields);
      expect(res.error?.code, message).toBe("bad_request");
      expect(res.error.message).toContain(message);
    }
    expect(snapshot(sheet)).toEqual(before);
    expect(env.LockService.scriptLock.calls).toEqual([]);
  });

  it("rejects a row name that changes the text before | (that would break the group)", async () => {
    seedGroup();
    const before = snapshot(sheet);
    for (const [r, name] of [
      [2, "Other | milk"],
      [2, "Store milk"],
      [1, "Market | "],
    ] as const) {
      const rows = await groupRows({ [r]: { name } });
      const res = await updateGroup({ rows });
      expect(res.error.message).toBe(
        `rows[${r - 1}].changes.name must keep the text before | (rename the group with title)`,
      );
    }
    expect(snapshot(sheet)).toEqual(before);
  });

  it("writes a math split total in parentheses", async () => {
    seedGroup();
    const res = await updateGroup({
      rows: await groupRows(),
      splitTotal: "(40+10)",
    });
    expect(res.ok).toBe(true);
    expect(sheet.getRange("C6").getValue()).toBe("=(40+10)-R[1]C[0]-R[2]C[0]");
  });

  it("rejects splitTotal when the parent's cost is not a split formula", async () => {
    seedGroup();
    sheet.getRange("C6").setValue(40);
    const rows = await groupRows();
    const res = await updateGroup({ rows, splitTotal: "50" });
    expect(res.error).toEqual({
      code: "bad_request",
      message:
        "splitTotal is only allowed when the first row is a split parent",
    });
    expect(sheet.getRange("C6").getValue()).toBe(40);
  });

  it("returns conflict for rows that are not one split group", async () => {
    seedGroup();
    const state = await getState();
    const notAGroup = [0, 1].map((r) => ({
      rowIndex: r,
      baseFingerprint: state.transactions[r].fingerprint,
      baseNote: "",
    }));
    const res = await updateGroup({ rows: notAGroup, title: "X" });
    expect(res.error.code).toBe("conflict");
    expect(sheet.getRange("B5:B6").getValues()).toEqual([
      ["Cafe"],
      ["Store | "],
    ]);
  });

  it("returns busy without writing when the script lock is held elsewhere", async () => {
    seedGroup();
    const rows = await groupRows({ 2: { envelope: "Gas" } });
    env.LockService.scriptLock.contended = true;
    const res = await updateGroup({ rows });
    expect(res.error.code).toBe("busy");
    expect(sheet.getRange("E7").getValue()).toBe("Food");
  });
});
