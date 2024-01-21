namespace Variables {
  /** Names of all global variables that must be provided */
  const VARIABLE_NAMES = [
    "Menu",
    "SheetVariables",
    "PayPeriodDays",
    "TemplateName",
    "TaxMultiplier",
  ] as const;

  /** Global variables from the Variables sheet */
  type Variables = {
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
  };

  /** Names of all sheet variables that must be provided */
  const SHEET_VARIABLE_NAMES = [
    "TransactionsStart",
    "TransactionsMax",
    "SplitCheckboxesStart",
  ] as const;
  /** Per-sheet variables */
  type SheetVariables = {
    /** Location of the first available space for transactions - A1 notation */
    TransactionsStart: string;
    /** Max number of transactions this sheet can support */
    TransactionsMax: number;
    /** Location of the first checkbox to click to split a transaction into two - A1 notation */
    SplitCheckboxesStart: string;
  };

  const VARIABLES_SHEET_NAME = "Variables";
  const VARIABLES_RANGE = `${VARIABLES_SHEET_NAME}!A1:B10`;

  let variablesRangeCache: GoogleAppsScript.Spreadsheet.Range | undefined;
  function getVariablesRange() {
    if (variablesRangeCache) return variablesRangeCache;

    variablesRangeCache =
      SpreadsheetApp.getActiveSheet().getRange(VARIABLES_RANGE);

    return variablesRangeCache!;
  }

  let variablesCache: Variables | undefined;
  /** Get global variables from the Variables sheet */
  export function getVariables(): Variables {
    if (variablesCache) return variablesCache;

    const variables: Variables = Object.fromEntries(
      getVariablesRange()
        .getValues()
        .filter(([key]) => key)
    );
    const variablesPresent = Object.keys(variables);

    if (
      VARIABLE_NAMES.some(
        (variableName) => !variablesPresent.includes(variableName)
      )
    )
      throw new Error(
        `Some variables were not found in the ${VARIABLES_SHEET_NAME} sheet in the range ${VARIABLES_RANGE}!`
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
        `Cannot get sheet variables range for ${VARIABLES_SHEET_NAME} sheet.`
      );

    const cachedSheetVariablesRange = sheetVariablesRangeCache.get(sheetName);
    if (cachedSheetVariablesRange) return cachedSheetVariablesRange;

    sheetVariablesRangeCache.set(
      sheetName,
      sheet.getRange(getVariables().SheetVariables)
    );

    return sheetVariablesRangeCache.get(sheetName)!;
  }

  const sheetVariablesCache = new Map<string, SheetVariables>();
  /** Get per-sheet variables for the provided sheet. Throws an error if the sheet is the Variables sheet */
  export function getSheetVariables(sheet: GoogleAppsScript.Spreadsheet.Sheet) {
    const sheetName = sheet.getName();
    if (sheetName === VARIABLES_SHEET_NAME)
      throw new Error(
        `Cannot get sheet variables for ${VARIABLES_SHEET_NAME} sheet.`
      );

    const cachedSheetVariables = sheetVariablesCache.get(sheetName);
    if (cachedSheetVariables) return cachedSheetVariables;

    const sheetVariables: SheetVariables = Object.fromEntries(
      getSheetVariablesRange(sheet)
        .getValues()
        .filter(([key]) => key)
    );
    const sheetVariablesPresent = Object.keys(sheetVariables);

    if (
      SHEET_VARIABLE_NAMES.some(
        (sheetVariableName) =>
          !sheetVariablesPresent.includes(sheetVariableName)
      )
    )
      throw new Error(
        `Some sheet variables were not found in the ${sheet.getName()} sheet in the range ${
          getVariables().SheetVariables
        }!`
      );

    sheetVariablesCache.set(sheetName, sheetVariables);

    return sheetVariables;
  }

  function bustSheetVariablesCache(
    sheet: GoogleAppsScript.Spreadsheet.Sheet | undefined
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

  function bustVariablesCache() {
    variablesRangeCache = undefined;
    variablesCache = undefined;
    bustSheetVariablesCache(undefined);
  }

  /**
   * Unfortunately I realized too late that scripts load and run every edit instead of staying running,
   * so caching is useless. If we ever find a reason to bust the cache on edit, here it is.
   */
  /* export function onEdit(e: GoogleAppsScript.Events.SheetsOnEdit) {
    const sheet = e.range.getSheet();
    if (sheet.getName() === VARIABLES_SHEET_NAME) {
      if (SpreadsheetUtil.doRangesIntersect(e.range, getVariablesRange())) {
        bustVariablesCache();
      }
    } else if (
      SpreadsheetUtil.doRangesIntersect(e.range, getSheetVariablesRange(sheet))
    ) {
      bustSheetVariablesCache(sheet);
    }
  } */
}
