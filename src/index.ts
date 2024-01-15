/// <reference path="./models/receipt-info.model.ts" />
/// <reference path="./models/thread-info.model.ts" />
/// <reference path="./models/thread-list.model.ts" />

/** Menu to add to the spreadsheet and to show in the menu bar on desktop */
const menu: {
  [menuName: string]: ({ name: string; functionName: string } | undefined)[];
} = {
  Budgeting: [
    {
      name: "Import 50 Receipts (No mark)",
      functionName: "getAndRecordSomeReceiptsNoMark",
    },
    {
      name: "Import All Receipts (No mark)",
      functionName: "getAndRecordReceiptsNoMark",
    },
    undefined,
    {
      name: "Import 25 Receipts and mark",
      functionName: "getAndRecordSomeReceiptsAndMark",
    },
    {
      name: "Import All Receipts and mark",
      functionName: "getAndRecordReceiptsAndMark",
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

  // Handle menu checkboxes
  const menuTopLeftRange = SpreadsheetApp.getActiveSheet().getRange(
    Variables.getVariables().Menu
  );
  if (e.range.getSheet().getName() === menuTopLeftRange.getSheet().getName()) {
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
        editRow < menuRowBase + numMenuRowsTotal &&
        e.range.isChecked()
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
          e.range.setValue("FALSE");
          this[checkedFunctionName]();
        }
      }
    }
  }
}

function getAndRecordSomeReceiptsNoMark() {
  Budgeting.getAndRecordReceipts(0, 50, false);
}

function getAndRecordReceiptsNoMark() {
  Budgeting.getAndRecordReceipts(null, null, false);
}

function getAndRecordSomeReceiptsAndMark() {
  Budgeting.getAndRecordReceipts(0, 25, true);
}

function getAndRecordReceiptsAndMark() {
  Budgeting.getAndRecordReceipts(null, null, true);
}

// Test scripts
/* function logTheadById(id: string) {
  Logger.log(GmailApp.getThreadById(id));
}
function logReceipts(
  start: number | null = 0,
  max: number | null = 10,
  shouldMarkProcessed = false
) {
  const threadList = Budgeting.getChaseReceipts(start, max);
  Logger.log(`ThreadList: ${JSON.stringify(threadList)}`);
  Logger.log(`ReceiptInfos: ${JSON.stringify(threadList.receiptInfos)}`);
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
} */
