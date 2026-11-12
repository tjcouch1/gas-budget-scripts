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

/** API key of the demo backend. Printed at startup; fake by design */
export const DEMO_API_KEY = "dev-key";

export const TAX_MULTIPLIER = 1.0825;

// Per-sheet layout (web-plan.md sections 3.3 to 3.6)
export const TRANSACTIONS_START = "A5";
export const TRANSACTIONS_MAX = 40;
const CHECKBOXES_START = "H5";
const SHEET_VARIABLES_RANGE = "J1:K7";
const ENVELOPES_START = "J10";
const ENVELOPES_MAX = 14;
const PAID_WITH_START = "V10";
const PAID_WITH_MAX = 9;

export const ENVELOPE_NAMES = [
  "Groceries",
  "Dining Out",
  "Gas",
  "Household",
  "Utilities",
  "Entertainment",
  "Medical",
  "Gifts",
  "Clothing",
  "Pets",
  "Charity",
];
export const PAID_WITH_OPTIONS = [
  "Credit",
  "Debit",
  "Cash",
  "TJ Venmo",
  "TJ Paypal",
  "Sam Venmo",
  "Sam Paypal",
];

/** Envelope size per pay period */
const ENVELOPE_SIZES: Record<string, number> = {
  Groceries: 320,
  "Dining Out": 120,
  Gas: 90,
  Household: 80,
  Utilities: 150,
  Entertainment: 60,
  Medical: 50,
  Gifts: 40,
  Clothing: 70,
  Pets: 45,
  Charity: 30,
};
/** Unspent balance carried into the period, per envelope */
const ENVELOPE_CARRIED: Record<string, number> = {
  Groceries: 25.5,
  "Dining Out": -14.2,
  Gas: 33,
  Household: 12.75,
  Utilities: 0,
  Entertainment: 48.1,
  Medical: 120,
  Gifts: 66,
  Clothing: -22.4,
  Pets: 10,
  Charity: 0,
};

type TransactionSpec = {
  /** Days after the start of the pay period */
  day: number;
  name: string;
  /** Number, or a formula in R1C1 (which is how the backend writes them) */
  cost: number | string;
  envelope?: string;
  paidWith?: string;
  taxDeductible?: boolean;
  note?: string;
  /** Marks the name cell like the receipt import does */
  flag?: "error" | "note";
};

/** Rows 0-2 form an existing split group (parent subtracts both children) */
const CURRENT_PERIOD: TransactionSpec[] = [
  {
    day: 0,
    name: "Fake Mart | ",
    cost: "=87.5-R[1]C[0]-R[2]C[0]",
    envelope: "Groceries",
    paidWith: "Credit",
  },
  {
    day: 0,
    name: "Fake Mart | Paper towels",
    cost: "=1.0825*(9.5)",
    envelope: "Household",
    paidWith: "Credit",
  },
  {
    day: 0,
    name: "Fake Mart | Birthday card",
    cost: 3.99,
    envelope: "Gifts",
    paidWith: "Credit",
  },
  { day: 1, name: "Sample Cafe", cost: 12.35, paidWith: "Credit" },
  {
    day: 1,
    name: "Example Fuel Stop (Gas)",
    cost: 41.2,
    envelope: "Gas",
    paidWith: "Credit",
  },
  {
    day: 2,
    name: "Demo Power Co",
    cost: 96.4,
    envelope: "Utilities",
    paidWith: "Debit",
  },
  { day: 2, name: "Test Hardware", cost: 27.83, paidWith: "Credit" },
  {
    day: 3,
    name: "Placeholder Pizza",
    cost: 31.5,
    envelope: "Dining Out",
    paidWith: "Credit",
  },
  { day: 3, name: "Imaginary Books", cost: 18.99, paidWith: "Credit" },
  {
    day: 3,
    name: "Mock Pharmacy",
    cost: 14.25,
    envelope: "Medical",
    paidWith: "Credit",
    taxDeductible: true,
  },
  {
    day: 4,
    name: "Jane Example - Lunch money",
    cost: 15,
    paidWith: "TJ Venmo",
  },
  {
    day: 4,
    name: "Fictional Grocer",
    cost: 64.12,
    envelope: "Groceries",
    paidWith: "Credit",
  },
  {
    day: 4,
    name: "Pretend Pet Supply",
    cost: 22.8,
    envelope: "Pets",
    paidWith: "Credit",
  },
  {
    day: 5,
    name: "Demo Streaming",
    cost: 15.49,
    envelope: "Entertainment",
    paidWith: "Credit",
  },
  {
    day: 5,
    name: "Unknown Charge 1234",
    cost: 8.75,
    paidWith: "Credit",
    note: "Could not tell what this was. Check the statement.",
    flag: "note",
  },
  { day: 5, name: "Faux Bakery", cost: 9.4, paidWith: "Cash" },
  {
    day: 6,
    name: "Sample Cafe",
    cost: 7.8,
    envelope: "Dining Out",
    paidWith: "Credit",
  },
  {
    day: 6,
    name: "Test Gift Shop",
    cost: 35,
    envelope: "Gifts",
    paidWith: "Credit",
  },
  { day: 7, name: "Dummy Department Store", cost: 58.6, paidWith: "Credit" },
  {
    day: 7,
    name: "Sam Example - Movie night",
    cost: 24,
    paidWith: "Sam Venmo",
  },
  {
    day: 8,
    name: "Fake Charity Drive",
    cost: 20,
    envelope: "Charity",
    paidWith: "Debit",
    taxDeductible: true,
  },
  { day: 8, name: "Cafe Refund", cost: -6.5, paidWith: "Credit" },
  {
    day: 9,
    name: "Example Fuel Stop (Gas)",
    cost: 38.9,
    envelope: "Gas",
    paidWith: "Credit",
  },
  {
    day: 9,
    name: "Unreadable receipt",
    cost: 0,
    paidWith: "Credit",
    note: "Error: could not parse the email.",
    flag: "error",
  },
  { day: 10, name: "Placeholder Tacos", cost: 17.25, paidWith: "Credit" },
  {
    day: 10,
    name: "Demo Clothing Outlet",
    cost: 44.99,
    envelope: "Clothing",
    paidWith: "Credit",
  },
  {
    day: 11,
    name: "Fictional Grocer",
    cost: 71.3,
    envelope: "Groceries",
    paidWith: "Credit",
  },
];

const PREVIOUS_PERIOD: TransactionSpec[] = [
  {
    day: 0,
    name: "Fictional Grocer",
    cost: 82.44,
    envelope: "Groceries",
    paidWith: "Credit",
  },
  {
    day: 1,
    name: "Demo Water Utility",
    cost: 54.1,
    envelope: "Utilities",
    paidWith: "Debit",
  },
  {
    day: 2,
    name: "Sample Cafe",
    cost: 11.2,
    envelope: "Dining Out",
    paidWith: "Credit",
  },
  {
    day: 3,
    name: "Example Fuel Stop (Gas)",
    cost: 44.05,
    envelope: "Gas",
    paidWith: "Credit",
  },
  {
    day: 4,
    name: "Test Hardware",
    cost: 33.6,
    envelope: "Household",
    paidWith: "Credit",
  },
  {
    day: 5,
    name: "Mock Pharmacy",
    cost: 19.99,
    envelope: "Medical",
    paidWith: "Credit",
    taxDeductible: true,
  },
  {
    day: 7,
    name: "Fictional Grocer",
    cost: 59.7,
    envelope: "Groceries",
    paidWith: "Credit",
  },
  {
    day: 8,
    name: "Demo Streaming",
    cost: 15.49,
    envelope: "Entertainment",
    paidWith: "Credit",
  },
  { day: 10, name: "Faux Bakery", cost: 6.25, paidWith: "Cash" },
  {
    day: 12,
    name: "Pretend Pet Supply",
    cost: 31.4,
    envelope: "Pets",
    paidWith: "Credit",
  },
];

/** The oldest tab predates the optional sheet variables (no envelope or Paid with tables) */
const OLDEST_PERIOD: TransactionSpec[] = [
  { day: 0, name: "Fictional Grocer", cost: 77.1, envelope: "Groceries" },
  { day: 2, name: "Sample Cafe", cost: 9.9, envelope: "Dining Out" },
  { day: 4, name: "Example Fuel Stop (Gas)", cost: 40.5, envelope: "Gas" },
  { day: 6, name: "Demo Power Co", cost: 101.2, envelope: "Utilities" },
  { day: 9, name: "Test Hardware", cost: 12.3, envelope: "" },
];

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

/** Fake Gmail receipts in formats the parsers accept, plus a non-receipt email */
export function createDemoThreads() {
  // Yesterday, inside the current pay period
  const date = new Date(chicagoToday() - DAY_MS + 18 * 60 * 60 * 1000);
  const CHASE = "Chase <no.reply.alerts@chase.com>";
  const PAYPAL = "PayPal <service@paypal.com>";
  const VENMO = "Venmo <venmo@venmo.com>";
  return [
    makeThread({
      from: CHASE,
      subject: "You made a $18.42 transaction with Demo Cafe",
      body: "Your card was used.",
      date,
    }),
    makeThread({
      from: CHASE,
      subject: "You used your card at a gas station",
      body: "Hello\nMerchant   Fake Fuel 42\n\nMore text\n",
      date,
    }),
    makeThread({
      from: VENMO,
      subject: "You paid Jane Example $15.00",
      body: "You paid\nJane Example\n$15.00\nPizza night\nTransaction details\n",
      date,
    }),
    makeThread({
      from: PAYPAL,
      subject: "You sent a $25.50 USD payment",
      body: "Hello,\n\nYou sent $25.50 to Sample Seller\n\nTransaction ID\nABC123\n\nLamp\nThanks\n",
      date,
    }),
    makeThread({
      from: CHASE,
      subject: "Your statement is ready",
      body: "Statement text.",
      date,
    }),
  ];
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
