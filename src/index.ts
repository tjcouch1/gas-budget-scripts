/// <reference path="./models/receipt-info.model.ts" />
/// <reference path="./models/thread-info.model.ts" />
/// <reference path="./models/thread-list.model.ts" />

function onOpen() {
  // Menu only works on desktop :(
  // Links for potential workarounds:
  // https://stackoverflow.com/questions/77385083/custom-menu-for-mobile-version-of-google-sheet
  // https://stackoverflow.com/questions/57840757/button-click-is-only-working-on-windows-not-working-on-android-mobile-sheet
  // https://webapps.stackexchange.com/questions/87346/add-a-script-trigger-to-google-sheet-that-will-work-in-android-mobile-app
  const ui = SpreadsheetApp.getUi();
  ui.createMenu("Budgeting")
    .addItem("Import 50 Receipts (No mark)", "getAndRecordSomeReceiptsNoMark")
    .addItem("Import All Receipts (No mark)", "getAndRecordReceiptsNoMark")
    .addSeparator()
    .addItem("Import 25 Receipts and mark", "getAndRecordSomeReceiptsAndMark")
    .addItem("Import All Receipts and mark", "getAndRecordReceiptsAndMark")
    .addToUi();
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
