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
      }
    | undefined;

  const VARIABLES_SHEET_NAME = "Variables";
  const VARIABLES_RANGE = `${VARIABLES_SHEET_NAME}!A1:B20`;

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

    let start = Date.now();
    const range = getVariablesRange();
    let end = Date.now();
    Logger.log(`Getting variables range took ${end - start} ms`);
    start = Date.now();
    const values = range.getValues();
    end = Date.now();
    Logger.log(`Getting variables values took ${end - start} ms`);
    start = Date.now();
    variablesCache = Object.fromEntries(values.filter(([key]) => key));
    end = Date.now();
    Logger.log(`Building variables cache took ${end - start} ms`);

    return variablesCache!;
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

    const start = Date.now();
    sheetVariablesCache.set(
      sheetName,
      Object.fromEntries(
        sheet
          .getRange(getVariables().SheetVariables)
          .getValues()
          .filter(([key]) => key)
      )
    );
    const end = Date.now();
    Logger.log(
      `Get uncached sheet variables for ${sheet} took ${end - start} ms`
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

  export function onEdit(e: GoogleAppsScript.Events.SheetsOnEdit) {
    const start = Date.now();
    const sheet = e.range.getSheet();
    if (sheet.getName() === VARIABLES_SHEET_NAME) {
      if (SpreadsheetUtil.doRangesIntersect(e.range, getVariablesRange())) {
        bustVariablesCache();
        const end = Date.now();
        SpreadsheetApp.getUi().alert(
          `Busted ${VARIABLES_SHEET_NAME} cache in ${end - start} ms!`
        );
      }
    } else if (
      SpreadsheetUtil.doRangesIntersect(e.range, getSheetVariablesRange(sheet))
    ) {
      bustSheetVariablesCache(sheet);
      const end = Date.now();
      SpreadsheetApp.getUi().alert(
        `Busted ${sheet.getName()} sheet variables cache in ${end - start} ms!`
      );
    }
  }
}
