import { MockGmailApp } from "./gmail";
import { MockSpreadsheet, MockSpreadsheetApp } from "./spreadsheet";

export * from "./gmail";
export * from "./spreadsheet";

/** Sheet handle typed loosely: the mock implements only a subset of the real Sheet interface */
// eslint-disable-next-line @typescript-eslint/no-explicit-any
export type AnySheet = any;

export interface GasEnv {
  spreadsheet: MockSpreadsheet;
  SpreadsheetApp: MockSpreadsheetApp;
  GmailApp: MockGmailApp;
  logs: unknown[];
}

export const DEFAULT_VARIABLES: unknown[][] = [
  ["Menu", "A1"],
  ["SheetVariables", "H1:I3"],
  ["PayPeriodDays", 14],
  ["TemplateName", "Template"],
  ["TaxMultiplier", 1.0825],
  ["ForwardEmailAddress", "fwd@example.com"],
  ["ForwardName", "Sam"],
];

let currentEnv: GasEnv | undefined;

/** The environment installed for the current test */
export function getEnv(): GasEnv {
  if (!currentEnv) throw new Error('GAS globals are not installed');
  return currentEnv;
}

/** Create a fresh in-memory environment and install the mocks on globalThis */
export function installGasGlobals(): GasEnv {
  const spreadsheet = new MockSpreadsheet();
  const logs: unknown[] = [];
  const env: GasEnv = {
    spreadsheet,
    SpreadsheetApp: new MockSpreadsheetApp(spreadsheet),
    GmailApp: new MockGmailApp(),
    logs,
  };
  currentEnv = env;
  Object.assign(globalThis, {
    SpreadsheetApp: env.SpreadsheetApp,
    GmailApp: env.GmailApp,
    Logger: {
      log: (...args: unknown[]) =>
        logs.push(args.length === 1 ? args[0] : args),
    },
  });
  return env;
}

/** Add a Variables sheet whose A1:B? holds [name, value] rows */
export function addVariablesSheet(env: GasEnv, vars: unknown[][] = DEFAULT_VARIABLES): AnySheet {
  const sheet = env.spreadsheet.addSheet("Variables");
  sheet.seed("A1", vars);
  return sheet;
}

/** Add a transaction sheet with per-sheet variables at H1:I3 */
export function addTransactionSheet(
  env: GasEnv,
  name: string,
  opts: { start?: string; max?: number; checkboxes?: string } = {}
) {
  const sheet = env.spreadsheet.addSheet(name);
  sheet.seed("H1", [
    ["TransactionsStart", opts.start ?? "A5"],
    ["TransactionsMax", opts.max ?? 10],
    ["SplitCheckboxesStart", opts.checkboxes ?? "G5"],
  ]);
  return sheet;
}
