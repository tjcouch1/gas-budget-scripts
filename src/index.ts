/// <reference path="./models/receipt-info.model.ts" />
/// <reference path="./models/thread-info.model.ts" />
/// <reference path="./models/thread-list.model.ts" />

function onOpen() {
  const ss = SpreadsheetApp.getActiveSpreadsheet();
  const menuEntries: Parameters<
    GoogleAppsScript.Spreadsheet.Spreadsheet["addMenu"]
  >[1] = [];
  menuEntries.push({
    name: "Import 50 Receipts (No mark)",
    functionName: "getAndRecordSomeReceiptsNoMark",
  });
  menuEntries.push({
    name: "Import All Receipts (No mark)",
    functionName: "getAndRecordReceiptsNoMark",
  });
  menuEntries.push(null); // divider
  menuEntries.push({
    name: "Import 25 Receipts and mark",
    functionName: "getAndRecordSomeReceiptsAndMark",
  });
  menuEntries.push({
    name: "Import All Receipts and mark",
    functionName: "getAndRecordReceiptsAndMark",
  });
  ss.addMenu("Receipts", menuEntries);
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
