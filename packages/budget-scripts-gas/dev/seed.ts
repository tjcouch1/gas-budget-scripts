/**
 * Sample spreadsheet and Gmail data for the local demo backend. Everything here is obviously fake
 * (no real merchants, people, emails, or URLs). Sheets follow web-plan.md section 3.
 */
import {
  addVariablesSheet,
  makeThread,
  type AnySheet,
  type GasEnv,
} from "../test/gas-mocks/index.ts";

import { demoInbox } from "../test/fixtures/receipts.ts";
import {
  CHECKBOXES_START,
  CURRENT_PERIOD,
  ENVELOPE_CARRIED,
  ENVELOPE_NAMES,
  ENVELOPE_SIZES,
  ENVELOPES_MAX,
  ENVELOPES_START,
  OLDEST_PERIOD,
  PAID_WITH_MAX,
  PAID_WITH_OPTIONS,
  PAID_WITH_START,
  PREVIOUS_PERIOD,
  SHEET_VARIABLES_RANGE,
  TAX_MULTIPLIER,
  TRANSACTIONS_MAX,
  TRANSACTIONS_START,
  type TransactionSpec,
} from "../test/fixtures/sheets.ts";

export { TRANSACTIONS_MAX, TRANSACTIONS_START };

/** API key of the demo backend. Printed at startup; fake by design */
export const DEMO_API_KEY = "dev-key";

// #region dates

const DAY_MS = 24 * 60 * 60 * 1000;

/** Today's calendar date in the spreadsheet time zone, as a UTC-midnight timestamp */
function chicagoToday(): number {
  const [y, m, d] = new Intl.DateTimeFormat("en-CA", {
    timeZone: "America/Chicago",
    year: "numeric",
    month: "2-digit",
    day: "2-digit",
  })
    .format(new Date())
    .split("-")
    .map(Number);
  return Date.UTC(y, m - 1, d);
}

/** `M/D/YY` like the sheet names (not zero-padded) */
function formatSheetDate(utcMidnight: number): string {
  const date = new Date(utcMidnight);
  return `${date.getUTCMonth() + 1}/${date.getUTCDate()}/${String(date.getUTCFullYear()).slice(2)}`;
}

/** A transaction date: noon in America/Chicago (UTC-6 or UTC-5), so it is the same calendar day there */
function transactionDate(utcMidnight: number): Date {
  return new Date(utcMidnight + 18 * 60 * 60 * 1000);
}

/** The tab name and start timestamp of the pay period `periodsAgo` before the current one */
export function payPeriod(periodsAgo: number) {
  // The current period started 5 days ago, so "now" is always inside it
  const start = chicagoToday() - 5 * DAY_MS - periodsAgo * 14 * DAY_MS;
  const end = start + 13 * DAY_MS;
  return {
    start,
    name: `${formatSheetDate(start)} - ${formatSheetDate(end)}`,
  };
}

// #endregion

function writeSheetVariables(sheet: AnySheet, includeOptional: boolean) {
  const rows: unknown[][] = [
    ["TransactionsStart", TRANSACTIONS_START],
    ["TransactionsMax", TRANSACTIONS_MAX],
    ["SplitCheckboxesStart", CHECKBOXES_START],
  ];
  if (includeOptional)
    rows.push(
      ["EnvelopesStart", ENVELOPES_START],
      ["EnvelopesMax", ENVELOPES_MAX],
      ["PaidWithStart", PAID_WITH_START],
      ["PaidWithMax", PAID_WITH_MAX],
    );
  sheet.seed("J1", rows);
}

function writeTransactionLayout(sheet: AnySheet, title: string) {
  sheet.seed("A1", [[title]]);
  sheet.seed("C3", [["Starting balance", 2400]]);
  sheet.seed("A4", [
    [
      "Date",
      "Name",
      "Cost",
      "Current Balance",
      "Envelope",
      "Paid with",
      "Tax Deductible",
    ],
  ]);
  // Current Balance (col 3) is a fill-down formula on every row, like the real sheet
  sheet.seed(
    "D5",
    Array.from({ length: TRANSACTIONS_MAX }, (_, r) => [
      r === 0 ? "=2400-RC[-1]" : "=R[-1]C[0]-RC[-1]",
    ]),
  );
  sheet
    .getRange(CHECKBOXES_START)
    .offset(0, 0, TRANSACTIONS_MAX, 1)
    .insertCheckboxes();
}

function writeTransactions(
  sheet: AnySheet,
  periodStart: number,
  specs: TransactionSpec[],
) {
  specs.forEach((spec, r) => {
    const row = 5 + r;
    sheet.seed(`A${row}`, [
      [transactionDate(periodStart + spec.day * DAY_MS), spec.name, spec.cost],
    ]);
    sheet.seed(`E${row}`, [
      [
        spec.envelope ?? "",
        spec.paidWith ?? "",
        spec.taxDeductible ? "Yes" : "",
      ],
    ]);
    const nameCell = sheet.getRange(`B${row}`);
    if (spec.note) nameCell.setNote(spec.note);
    if (spec.flag === "error") nameCell.setBackground("#FF0000");
    if (spec.flag === "note") nameCell.setBackground("#E8A9CA");
  });
}

function writeOptionTables(sheet: AnySheet, specs: TransactionSpec[]) {
  // Spent comes from the evaluated Cost cells of this tab
  const spent = new Map<string, number>();
  sheet
    .getRange(TRANSACTIONS_START)
    .offset(0, 0, specs.length, 7)
    .getValues()
    .forEach((row: unknown[]) => {
      const envelope = String(row[4]);
      if (typeof row[2] === "number" && envelope)
        spent.set(envelope, (spent.get(envelope) ?? 0) + row[2]);
    });
  const round = (n: number) => Math.round(n * 100) / 100;

  sheet.seed("J9", [
    [
      "Envelope",
      "Envelope Available Cumulative",
      "Envelope Size",
      "Spent",
      "Remain",
      "Remain Cumul",
      "Envelope Size YTD",
      "Spent YTD",
      "Spent Average YTD",
      "Envelope Size Year Projected",
      "Spent Year Projected",
    ],
  ]);
  // Balances are plain numbers: they do not update when transactions change
  const rows = ENVELOPE_NAMES.map((name) => {
    const size = ENVELOPE_SIZES[name];
    const spentNow = round(spent.get(name) ?? 0);
    const remain = round(size - spentNow);
    const carried = ENVELOPE_CARRIED[name];
    return [
      name,
      round(size + carried),
      size,
      spentNow,
      remain,
      round(remain + carried),
      round(size * 10),
      round(spentNow * 9 + 12.5),
      round((spentNow * 9 + 12.5) / 10),
      round(size * 26),
      round(((spentNow * 9 + 12.5) / 10) * 26),
    ];
  });
  // A totals row inside EnvelopesMax: the API skips it
  rows.push(["Total", ...Array<number>(10).fill(0)]);
  sheet.seed(ENVELOPES_START, rows);

  sheet.seed("V9", [["Paid with"]]);
  sheet.seed(PAID_WITH_START, [
    ...PAID_WITH_OPTIONS.map((name) => [name]),
    ["Tax deductible"],
  ]);
}

/** Fake Gmail receipts (from the shared receipt fixtures), dated yesterday: inside the current pay period */
export function createDemoThreads() {
  const date = new Date(chicagoToday() - DAY_MS + 18 * 60 * 60 * 1000);
  return Object.values(demoInbox).map((email) =>
    makeThread({ ...email, date }),
  );
}

/**
 * Replaces everything in the environment with fresh seed data. Mutates in place (the bundled `src/`
 * caches hold references to the spreadsheet), so the seed must stay deterministic for a given day.
 */
export function seedDemo(env: GasEnv) {
  const { spreadsheet } = env;
  spreadsheet.sheets.length = 0;
  spreadsheet.activeSheet = undefined;
  env.logs.length = 0;
  env.GmailApp.threads = createDemoThreads();
  env.PropertiesService.getScriptProperties().setProperty(
    "API_KEY",
    DEMO_API_KEY,
  );

  addVariablesSheet(env, [
    ["Menu", "A1"],
    ["SheetVariables", SHEET_VARIABLES_RANGE],
    ["PayPeriodDays", 14],
    ["TemplateName", "Template"],
    ["TaxMultiplier", TAX_MULTIPLIER],
    ["ForwardEmailAddress", "fwd@example.com"],
    ["ForwardName", "Sam"],
  ]);

  const template = spreadsheet.addSheet("Template");
  writeSheetVariables(template, true);
  writeTransactionLayout(template, "Template");

  const periods: [number, TransactionSpec[], boolean][] = [
    [2, OLDEST_PERIOD, false],
    [1, PREVIOUS_PERIOD, true],
    [0, CURRENT_PERIOD, true],
  ];
  for (const [periodsAgo, specs, optional] of periods) {
    const { start, name } = payPeriod(periodsAgo);
    const sheet = spreadsheet.addSheet(name);
    writeSheetVariables(sheet, optional);
    writeTransactionLayout(sheet, name);
    writeTransactions(sheet, start, specs);
    if (optional) writeOptionTables(sheet, specs);
  }
}
