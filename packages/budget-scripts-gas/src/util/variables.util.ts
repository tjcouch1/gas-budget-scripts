/** Names of all global variables that must be provided */
const VARIABLE_NAMES = [
  "Menu",
  "SheetVariables",
  "PayPeriodDays",
  "TemplateName",
  "TaxMultiplier",
  "ForwardEmailAddress",
  "ForwardName",
] as const;

/** Global variables from the Variables sheet */
export type Variables = {
  /** Location of the top-left corner of the menu - A1 notation */
  Menu: string;
  /** Location of the per-sheet variables - A1 notation */
  SheetVariables: string;
  /** Length of one pay period in number of days */
  PayPeriodDays: number;
  /** Name of template sheet to duplicate to make new transaction sheets */
  TemplateName: string;
  /** Multiplier to use when computing tax for splitting transactions */
  TaxMultiplier: number;
  /** Email address of person who is forwarding emails to the main inbox */
  ForwardEmailAddress: string;
  /** Name of person who is forwarding emails to the main inbox */
  ForwardName: string;
};

/** Names of all sheet variables that must be provided */
const SHEET_VARIABLE_NAMES = [
  "TransactionsStart",
  "TransactionsMax",
  "SplitCheckboxesStart",
] as const;

/** Per-sheet variables */
export type SheetVariables = {
  /** Location of the first available space for transactions - A1 notation */
  TransactionsStart: string;
  /** Max number of transactions this sheet can support */
  TransactionsMax: number;
  /** Location of the first checkbox to click to split a transaction into two - A1 notation */
  SplitCheckboxesStart: string;
  // Optional keys (web-plan.md section 3.3). Deliberately NOT in SHEET_VARIABLE_NAMES so older tabs
  // without them keep working for receipt import and native splitting.
  /** Optional: first envelope-name data cell of the Envelope table - A1 notation */
  EnvelopesStart?: string;
  /** Optional: max number of envelope rows */
  EnvelopesMax?: number;
  /** Optional: first data cell of the Paid With table (left-most column = option names) - A1 notation */
  PaidWithStart?: string;
  /** Optional: max number of rows in the Paid With table */
  PaidWithMax?: number;
};

const VARIABLES_SHEET_NAME = "Variables";
const VARIABLES_RANGE = `${VARIABLES_SHEET_NAME}!A1:B10`;

let variablesRangeCache: GoogleAppsScript.Spreadsheet.Range | undefined;
function getVariablesRange() {
  if (variablesRangeCache) return variablesRangeCache;

  variablesRangeCache =
    SpreadsheetApp.getActiveSpreadsheet().getRange(VARIABLES_RANGE);

  return variablesRangeCache!;
}

let variablesCache: Variables | undefined;
/** Get global variables from the Variables sheet */
export function getVariables(): Variables {
  if (variablesCache) return variablesCache;

  const variables = Object.fromEntries(
    getVariablesRange()
      .getValues()
      .filter(([key]) => key),
  ) as Variables;
  const variablesPresent = Object.keys(variables);

  if (
    VARIABLE_NAMES.some(
      (variableName) => !variablesPresent.includes(variableName),
    )
  )
    throw new Error(
      `Some variables were not found in the ${VARIABLES_SHEET_NAME} sheet in the range ${VARIABLES_RANGE}!`,
    );

  variablesCache = variables;

  return variables;
}

export function areCached() {
  return variablesRangeCache || variablesCache;
}

const sheetVariablesRangeCache = new Map<
  string,
  GoogleAppsScript.Spreadsheet.Range | undefined
>();

/** Throws an error if the sheet is the Variables sheet */
function getSheetVariablesRange(sheet: GoogleAppsScript.Spreadsheet.Sheet) {
  const sheetName = sheet.getName();
  if (sheetName === VARIABLES_SHEET_NAME)
    throw new Error(
      `Cannot get sheet variables range for ${VARIABLES_SHEET_NAME} sheet.`,
    );

  const cachedSheetVariablesRange = sheetVariablesRangeCache.get(sheetName);
  if (cachedSheetVariablesRange) return cachedSheetVariablesRange;

  sheetVariablesRangeCache.set(
    sheetName,
    sheet.getRange(getVariables().SheetVariables),
  );

  return sheetVariablesRangeCache.get(sheetName)!;
}

const sheetVariablesCache = new Map<string, SheetVariables>();

/** Get per-sheet variables for the provided sheet. Throws an error if the sheet is the Variables sheet */
export function getSheetVariables(sheet: GoogleAppsScript.Spreadsheet.Sheet) {
  const sheetName = sheet.getName();
  if (sheetName === VARIABLES_SHEET_NAME)
    throw new Error(
      `Cannot get sheet variables for ${VARIABLES_SHEET_NAME} sheet.`,
    );

  const cachedSheetVariables = sheetVariablesCache.get(sheetName);
  if (cachedSheetVariables) return cachedSheetVariables;

  const sheetVariables = Object.fromEntries(
    getSheetVariablesRange(sheet)
      .getValues()
      .filter(([key]) => key),
  ) as SheetVariables;
  const sheetVariablesPresent = Object.keys(sheetVariables);

  if (
    SHEET_VARIABLE_NAMES.some(
      (sheetVariableName) => !sheetVariablesPresent.includes(sheetVariableName),
    )
  )
    throw new Error(
      `Some sheet variables were not found in the ${sheet.getName()} sheet in the range ${
        getVariables().SheetVariables
      }!`,
    );

  sheetVariablesCache.set(sheetName, sheetVariables);

  return sheetVariables;
}

function bustSheetVariablesCache(
  sheet: GoogleAppsScript.Spreadsheet.Sheet | undefined,
) {
  if (sheet) {
    const sheetName = sheet.getName();
    sheetVariablesRangeCache.delete(sheetName);
    sheetVariablesCache.delete(sheetName);
  } else {
    sheetVariablesRangeCache.clear();
    sheetVariablesCache.clear();
  }
}

// eslint-disable-next-line @typescript-eslint/no-unused-vars -- kept for when cache busting is needed again
function bustVariablesCache() {
  variablesRangeCache = undefined;
  variablesCache = undefined;
  bustSheetVariablesCache(undefined);
}

export const Variables = {
  getVariables,
  areCached,
  getSheetVariables,
};
