/**
 * Sample spreadsheet data shared by the local demo backend and any test or e2e spec that needs a
 * populated sheet. Pure data (no GAS mocks). Layout follows web-plan.md sections 3.3 to 3.6; every
 * name and merchant is obviously fake.
 */

export const TAX_MULTIPLIER = 1.0825;

// Per-sheet layout (web-plan.md sections 3.3 to 3.6)
export const TRANSACTIONS_START = "A5";
export const TRANSACTIONS_MAX = 40;
export const CHECKBOXES_START = "H5";
export const SHEET_VARIABLES_RANGE = "J1:K7";
export const ENVELOPES_START = "J10";
export const ENVELOPES_MAX = 14;
export const PAID_WITH_START = "V10";
export const PAID_WITH_MAX = 9;

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
export const ENVELOPE_SIZES: Record<string, number> = {
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
export const ENVELOPE_CARRIED: Record<string, number> = {
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

export type TransactionSpec = {
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
export const CURRENT_PERIOD: TransactionSpec[] = [
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

export const PREVIOUS_PERIOD: TransactionSpec[] = [
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
export const OLDEST_PERIOD: TransactionSpec[] = [
  { day: 0, name: "Fictional Grocer", cost: 77.1, envelope: "Groceries" },
  { day: 2, name: "Sample Cafe", cost: 9.9, envelope: "Dining Out" },
  { day: 4, name: "Example Fuel Stop (Gas)", cost: 40.5, envelope: "Gas" },
  { day: 6, name: "Demo Power Co", cost: 101.2, envelope: "Utilities" },
  { day: 9, name: "Test Hardware", cost: 12.3, envelope: "" },
];
