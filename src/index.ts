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

function onOpen() {
  // Set up menu bar items (desktop only)
  // Menu only works on desktop :(
  // Links for potential workarounds:
  // https://stackoverflow.com/questions/77385083/custom-menu-for-mobile-version-of-google-sheet
  // https://stackoverflow.com/questions/57840757/button-click-is-only-working-on-windows-not-working-on-android-mobile-sheet
  // https://webapps.stackexchange.com/questions/87346/add-a-script-trigger-to-google-sheet-that-will-work-in-android-mobile-app
  const ui = SpreadsheetApp.getUi();
  Object.entries(menu).forEach(([menuName, menuEntries]) => {
    const uiMenu = ui.createMenu(menuName);
    menuEntries.forEach((menuEntry) => {
      if (menuEntry) uiMenu.addItem(menuEntry.name, menuEntry.functionName);
      else uiMenu.addSeparator();
    });
    uiMenu.addToUi();
  });
}

function onEdit(e: GoogleAppsScript.Events.SheetsOnEdit) {
  Variables.onEdit(e);
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
}
