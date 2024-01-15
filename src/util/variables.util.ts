namespace Variables {
  /** Global variables from the Variables sheet */
  type Variables = {
    /** Location of the menu - A1 notation */
    Menu: string;
    /** Location of the per-sheet variables - A1 notation */
    SheetVariables: string;
  };

  /** Per-sheet variables */
  type SheetVariables =
    | {
        /** Location of the first available space for transactions - A1 notation */
        TransactionsStart: string;
        /** Max number of transactions this sheet can support */
        TransactionsMax: number;
      }
    | undefined;

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

    variablesCache = Object.fromEntries(
      getVariablesRange()
        .getValues()
        .filter(([key]) => key)
    );

    return variablesCache!;
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

    sheetVariablesCache.set(
      sheetName,
      Object.fromEntries(
        sheet
          .getRange(getVariables().SheetVariables)
          .getValues()
          .filter(([key]) => key)
      )
    );

    return sheetVariablesCache.get(sheetName)!;
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
