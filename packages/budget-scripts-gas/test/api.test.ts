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
      [D2, "Next", 7, "=BALANCE(2)", "Gas", "Sam Venmo", "Yes"],
      [D2, "After", "=1+2", "=BALANCE(3)", "", "Credit", ""],
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

  it("shifts following rows (cols 0-2 and 4-5 with notes and formats) down by N", async () => {
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
      [D, "First | ", "=1.0825*(0)", "=BALANCE(1)", "", "Credit", "T2"],
      [D, "First | ", "=1.0825*(0)", "=BALANCE(2)", "", "Credit", ""],
      [D, "First | ", "=1.0825*(0)", "=BALANCE(3)", "", "Credit", ""],
      [D2, "Second", 20, "=BALANCE(4)", "c2", "Cash", ""],
      [D2, "Third", 30, "=BALANCE(5)", "c3", "Credit", ""],
      ["", "", "", "=BALANCE(6)", "", "", ""],
    ]);
    expect(sheet.getRange("B9").getNotes()).toEqual([["second note"]]);
    expect(sheet.getRange("B9").getBackgrounds()).toEqual([["#E8A9CA"]]);
    expect(sheet.getRange("B6").getNotes()).toEqual([[""]]);
    expect(sheet.getRange("B6").getBackgrounds()).toEqual([[null]]);
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

describe("syncReceipts", () => {
  const chase = (cost: string, merchant: string, date: string) => ({
    from: "Chase <no.reply.alerts@chase.com>",
    subject: `You made a $${cost} transaction with ${merchant}`,
    body: "",
    date: new Date(date),
  });

  it("imports up to 25 threads, labels them, and returns errors without using the UI", async () => {
    env.SpreadsheetApp.uiAvailable = false;
    vi.useFakeTimers();
    vi.setSystemTime(new Date("2024-02-25T18:00:00Z"));
    const good = makeThread(chase("10.50", "Shop", "2024-02-20T12:00:00Z"));
    const bad = makeThread({
      from: "Weird <weird@example.org>",
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
      "No function to handle receipt message from weird@example.org",
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
