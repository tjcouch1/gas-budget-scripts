/// <reference path="./models/receipt-info.model.ts" />
/// <reference path="./models/thread-info.model.ts" />
/// <reference path="./models/thread-list.model.ts" />
/// <reference path="./util/budgeting.util.ts" />
/// <reference path="./util/spreadsheet-util.util.ts" />
/// <reference path="./util/util.ts" />
/// <reference path="./util/variables.util.ts" />

/** Menu to add to the spreadsheet and to show in the menu bar on desktop */
const menu: {
  [menuName: string]: ({ name: string; functionName: string } | undefined)[];
} = {
  Budgeting: [
    {
      name: "Import 50 receipts (no mark)",
      functionName: "getAndRecordSomeReceiptsNoMark",
    },
    {
      name: "Import all receipts (no mark)",
      functionName: "getAndRecordReceiptsNoMark",
    },
    undefined,
    {
      name: "Import 25 receipts and mark",
      functionName: "getAndRecordSomeReceiptsAndMark",
    },
    {
      name: "Import all receipts and mark",
      functionName: "getAndRecordReceiptsAndMark",
    },
    undefined,
    {
      name: "Add one transaction sheet if needed",
      functionName: "addTransactionSheet",
    },
    {
      name: "Add transaction sheets until up-to-date",
      functionName: "addTransactionSheets",
    },
  ],
};

// WARNING: This doesn't seem to run on mobile
function onOpen() {
  // Set up menu bar items (desktop only)
  try {
    const ui = SpreadsheetApp.getUi();
    Object.entries(menu).forEach(([menuName, menuEntries]) => {
      const uiMenu = ui.createMenu(menuName);
      menuEntries.forEach((menuEntry) => {
        if (menuEntry) uiMenu.addItem(menuEntry.name, menuEntry.functionName);
        else uiMenu.addSeparator();
      });
      uiMenu.addToUi();
    });
  } catch (e) {
    Logger.log(
      `Error on setting up menu bar items! User may be on mobile. ${e}`
    );
  }

  // Set up in-sheet menu (desktop and mobile)
  const menuRange = SpreadsheetApp.getActiveSheet().getRange(
    Variables.getVariables().Menu
  );
  if (!menuRange.getValue()) {
    // Assume menu is not filled in and fill it in
    const menuSheet = menuRange.getSheet();
    let nextRow = menuRange.getRow();
    const menuColumnBase = menuRange.getColumn();
    const menuNameStyle = SpreadsheetApp.newTextStyle().setBold(true).build();
    Object.entries(menu).forEach(([menuName, menuEntries]) => {
      // Set up the header for this menu section
      const menuNameRange = menuSheet.getRange(nextRow, menuColumnBase);
      menuNameRange.setValue(menuName);
      menuNameRange.setTextStyle(menuNameStyle);
      nextRow += 1;
      // Set up checkboxes for this menu section
      const menuCheckboxesRange = menuSheet.getRange(
        nextRow,
        menuColumnBase,
        menuEntries.length,
        1
      );
      menuCheckboxesRange.insertCheckboxes();
      // Set up menu entry names
      const menuNamesRange = menuSheet.getRange(
        nextRow,
        menuColumnBase + 1,
        menuEntries.length,
        1
      );
      menuNamesRange.setValues(
        menuEntries.map((menuEntry) => [menuEntry ? menuEntry.name : "-------"])
      );
      nextRow += menuEntries.length;
    });
  }
}

/** onEdit trigger to be installed to run with user's permissions */
function onEditInstalled(e: GoogleAppsScript.Events.SheetsOnEdit) {
  // The whole script is loaded and run every edit, so cache busting is useless
  // Variables.onEdit(e);

  if (e.range.isChecked()) {
    // #region Handle menu checkboxes
    const menuTopLeftRange = SpreadsheetApp.getActiveSheet().getRange(
      Variables.getVariables().Menu
    );
    const editSheet = e.range.getSheet();
    if (editSheet.getName() === menuTopLeftRange.getSheet().getName()) {
      const menuRowBase = menuTopLeftRange.getRow();
      const menuCheckboxesColumn = menuTopLeftRange.getColumn();
      if (e.range.getColumn() === menuCheckboxesColumn) {
        const menuObjectEntries = Object.entries(menu);
        const numMenuRowsTotal = menuObjectEntries.reduce(
          (currentRowsCount, [, menuEntries]) =>
            currentRowsCount + menuEntries.length,
          menuObjectEntries.length
        );
        const editRow = e.range.getRow();
        if (
          editRow >= menuRowBase &&
          editRow < menuRowBase + numMenuRowsTotal
        ) {
          // Checked a checkbox in the menu, so run the function
          /** The row of the checked menu item relative to the start of the menu */
          const checkedMenuEntryIndex = editRow - menuRowBase;
          /** The current row to check against the checked row relative to the start of the menu */
          let currentMenuEntryIndex = 0;
          /** The name of the function to run */
          let checkedFunctionName: string | undefined;
          menuObjectEntries.some(([, menuEntries]) => {
            currentMenuEntryIndex += 1;
            menuEntries.some((menuEntry) => {
              if (currentMenuEntryIndex === checkedMenuEntryIndex) {
                checkedFunctionName = menuEntry ? menuEntry.functionName : "";
                return true;
              }
              currentMenuEntryIndex += 1;
            });
            // We found the function, so stop looking
            if (checkedFunctionName !== undefined) return true;
          });

          if (checkedFunctionName) {
            // Run the function
            const result = this[checkedFunctionName]();

            // Write the return from the function next to the menu item
            e.range
              .offset(0, 2, 1, 1)
              .setValue(
                typeof result === "object" ? JSON.stringify(result) : result
              );

            // Uncheck the checkbox. Note: this will not happen if an error occurs
            // to keep a record of the error
            e.range.setValue("FALSE");
          }
        }
      }
    }
    // #endregion

    // #region Not on menu sheet - handle transaction split checkboxes
    else {
      const splitCheckboxesRange =
        Budgeting.getTransactionSplitCheckboxesRange(editSheet);
      if (e.range.getColumn() === splitCheckboxesRange.getColumn()) {
        const editRow = e.range.getRow();
        const checkboxesRowBase = splitCheckboxesRange.getRow();
        if (
          editRow >= checkboxesRowBase &&
          editRow <= splitCheckboxesRange.getLastRow()
        ) {
          // Checked a checkbox in the split checkboxes column, so split the transaction
          /** The row of the checked transaction relative to the start of the transactions */
          const checkedTransactionIndex = editRow - checkboxesRowBase;

          const result = Budgeting.splitTransaction(
            editSheet,
            checkedTransactionIndex
          );

          // Set selected cell to the cost of the new bottom transaction
          SpreadsheetApp.setActiveRange(
            result.offset(
              result.getNumRows() - 1,
              result.getNumColumns() - 1,
              1,
              1
            )
          );

          // Uncheck the checkbox. Note: this will not happen if an error occurs
          // to keep a record of the error
          e.range.setValue("FALSE");
        }
      }
    }
  }
}

function getAndRecordSomeReceiptsNoMark() {
  return Budgeting.getAndRecordReceipts(0, 50, false);
}

function getAndRecordReceiptsNoMark() {
  return Budgeting.getAndRecordReceipts(null, null, false);
}

function getAndRecordSomeReceiptsAndMark() {
  return Budgeting.getAndRecordReceipts(0, 25, true);
}

function getAndRecordReceiptsAndMark() {
  return Budgeting.getAndRecordReceipts(null, null, true);
}

function addTransactionSheet() {
  const start = Date.now();
  const didAdd = Budgeting.addTransactionSheet();
  const end = Date.now();
  Logger.log(
    `Adding ${didAdd ? "one" : "no"} transaction sheet took ${end - start} ms`
  );

  return didAdd;
}

function addTransactionSheets() {
  const start = Date.now();
  const numSheetsAdded = Budgeting.addTransactionSheets();
  const end = Date.now();
  Logger.log(
    `Adding ${numSheetsAdded} transaction sheets took ${end - start} ms`
  );

  return numSheetsAdded;
}

// Test scripts
function logTheadById(id: string) {
  Logger.log(GmailApp.getThreadById(id));
}
function logReceipts(
  start: number | null = 0,
  max: number | null = 10,
  shouldMarkProcessed = false
) {
  const threadList = Budgeting.getChaseReceipts(start, max);
  Logger.log(`ThreadList: ${JSON.stringify(threadList)}`);
  //Logger.log(`ThreadInfos: ${JSON.stringify(threadList.threadInfos)}`);
  /* Logger.log(
    `AllReceiptInfos: ${JSON.stringify(threadList.getAllReceiptInfos())}`
  ); */
}
function logVariables() {
  const start = Date.now();
  Logger.log(Variables.getVariables());
  Logger.log(
    Variables.getSheetVariables(
      SpreadsheetApp.getActiveSpreadsheet().getSheets()[0]
    )
  );
  const end = Date.now();
  Logger.log(end - start);
}
function testLogTransactionSheetInfos() {
  Logger.log(JSON.stringify(Budgeting.getTransactionSheetInfos()));
}
function testRangeCachePerformance() {
  // 150-200ms
  const range = SpreadsheetApp.getActiveSheet().getRange(1, 1, 100, 100);
  // 75-100ms
  range.getValues();
  // 6-8ms
  range.getValues();
  // 4ms
  const rangeDup = SpreadsheetApp.getActiveSheet().getRange(1, 1, 100, 100);
  const start = Date.now();
  // 7ms
  rangeDup.getValues();
  const end = Date.now();
  // 2ms
  const range2 = SpreadsheetApp.getActiveSheet().getRange(101, 101, 100, 101);
  // 45ms
  range2.getValues();
  Logger.log(end - start);
}
function testSplit() {
  Logger.log(
    JSON.stringify(
      Budgeting.splitTransaction(SpreadsheetApp.getActiveSheet(), 10)
    )
  );
}

// #region Budget sheet utility functions
const MONTHS = [
  "January",
  "February",
  "March",
  "April",
  "May",
  "June",
  "July",
  "August",
  "September",
  "October",
  "November",
  "December",
];

/**
 * Get the name of the sheet at an offset from the current one
 * @param off number of sheets offset
 * @param _seed functionless parameter that exists to bust Google Sheets' cache
 * for this function's return
 * @returns name of sheet offset from this sheet by `off`
 */
function sheetNameOffset(off: number, _seed: number) {
  const ss = SpreadsheetApp.getActiveSpreadsheet();
  const sheets = ss.getSheets();
  const offInd = ss.getActiveSheet().getIndex() + off;
  for (let i = 0; i < sheets.length; i++) {
    if (sheets[i].getIndex() === offInd) {
      return sheets[i].getName();
    }
  }
  return "No sheet at offset " + off;
}

/**
 * Get the name of the month offset from the input month
 * @param month current month name
 * @param off number of months offset
 * @param _seed functionless parameter that exists to bust Google Sheets' cache
 * for this function's return
 * @returns name of month offset from `month` by `off`
 */
function getMonthOffset(month: string, off: number, _seed: number) {
  const monthInd = MONTHS.indexOf(month);
  if (monthInd >= 0) {
    return MONTHS[(monthInd + off) % MONTHS.length];
  }
  return "No month named " + month;
}

// #endregion
