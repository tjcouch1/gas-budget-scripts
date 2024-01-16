namespace Budgeting {
  /** Sheet that has transactions on it and info about it */
  type TransactionSheetInfo = {
    sheet: GoogleAppsScript.Spreadsheet.Sheet;
    startDate: Date;
    endDate: Date;
    /**
     * Transaction receipt infos that should be added to this sheet.
     *
     * Does NOT represent the present receipts
     */
    receiptInfosToAdd: ReceiptInfoBase[];
  };

  /** Chase Gmail search query for receipt email threads */
  const chaseGmailSearchQuery =
    "from:(no.reply.alerts@chase.com) in:inbox NOT label:receipts NOT label:receipts-cru-reimburse NOT label:receipts-tax-deductible NOT label:receipts-scripted";

  /** RegExp pattern matching to transaction sheet name
   *
   * Named groups: `start`, `end` (dates)
   */
  const transactionSheetNameRegex = /(?<start>\S+)\s*-\s*(?<end>\S+)/;

  /** Cache for GmailLabels */
  const labelCache = {};

  /**
   * Get or create a label
   *
   * @param name name of the label to get. Defaults to "Receipts/Scripted"
   *
   * @returns Gmail label object
   *
   * Note: this function is memoized. Feel free to call it as many times as desired
   */
  function getGmailLabel(
    name = "Receipts/Scripted"
  ): GoogleAppsScript.Gmail.GmailLabel {
    if (labelCache[name]) return labelCache[name];

    let label = GmailApp.getUserLabelByName(name);
    if (!label) label = GmailApp.createLabel(name);

    labelCache[name] = label;

    return label;
  }

  /**
   * Marks a thread as processed by these receipt-recording scripts
   *
   * @param thread thread to mark as processed
   */
  function markThreadProcessed(thread: GoogleAppsScript.Gmail.GmailThread) {
    // Get the script label
    const scriptLabel = getGmailLabel();
    const receiptLabel = getGmailLabel("Receipts");

    Logger.log(`Marking thread ${thread.getId()} processed!`);

    // Add the labels and archive the thread
    thread.addLabel(scriptLabel);
    thread.addLabel(receiptLabel);
    thread.moveToArchive();
  }

  /**
   * Get receipt info for all unprocessed Chase emails in the inbox
   *
   * @param start index of starting thread in the query to get email threads. Set this and max to null to get all. Defaults to 0
   * @param max max number of threads from which to get receipts. Set this and start to null to get all. Defaults to 10
   *
   * @returns receipt stuff - threads and receipt info for all unprocessed Chase emails
   */
  export function getChaseReceipts(
    start: number | null = 0,
    max: number | null = 10
  ): ThreadList {
    // Query Gmail for receipt emails
    const threads =
      start !== null && max !== null
        ? GmailApp.search(chaseGmailSearchQuery, start, max)
        : GmailApp.search(chaseGmailSearchQuery);

    return new Budgeting.ThreadList(threads);
  }

  /**
   * Get full range of all transaction rows in a sheet. Contains only date, name, and cost columns
   * @param sheet
   */
  function getTransactionsRange(sheet: GoogleAppsScript.Spreadsheet.Sheet) {
    // Get top left of range
    const transactionsTopLeftRange = sheet.getRange(
      Variables.getSheetVariables(sheet).TransactionsStart
    );
    // Expand to full range of all transactions
    const transactionsRange = transactionsTopLeftRange.offset(
      0,
      0,
      Variables.getSheetVariables(sheet).TransactionsMax,
      3
    );
    return transactionsRange;
  }

  /**
   * Returns an array of all sheets that have transactions on them according to the sheet name
   * along with the start and end dates for each sheet
   *
   * NOTE: each transaction sheet comes with an empty array `receiptInfosToAdd` for adding stuff into later
   */
  export function getTransactionSheetInfos() {
    const transactionSheetInfos: TransactionSheetInfo[] = [];
    SpreadsheetApp.getActiveSpreadsheet()
      .getSheets()
      .forEach((sheet) => {
        const matches = transactionSheetNameRegex.exec(sheet.getName());
        if (matches && matches.length === 3 && matches.groups) {
          const startDate = new Date(
            Date.parse(`${matches.groups.start} 00:00:00.000 GMT-06:00`)
          );
          const endDate = new Date(
            Date.parse(`${matches.groups.end} 23:59:59.999 GMT-06:00`)
          );
          transactionSheetInfos.push({
            sheet,
            startDate,
            endDate,
            receiptInfosToAdd: [],
          });
        }
      });
    return transactionSheetInfos;
  }

  /**
   * Compares transactionSheetInfos by startDate in descending order
   *
   * Used in [`Array.prototype.sort`](https://developer.mozilla.org/en-US/docs/Web/JavaScript/Reference/Global_Objects/Array/sort)
   */
  function compareTransactionSheetInfosByStartDateDescending(
    a: TransactionSheetInfo,
    b: TransactionSheetInfo
  ) {
    if (a.startDate > b.startDate) return -1;
    if (a.startDate === b.startDate) return 0;
    return 1;
  }

  /**
   * Adds a new transaction sheet if needed to catch up to current date
   *
   * @returns true if added a new sheet, false otherwise
   */
  export function addTransactionSheet() {
    /**
     * All transaction sheets for this spreadsheet in date-descending order
     * to make sure we don't duplicate date ranges
     */
    const transactionSheetInfos = getTransactionSheetInfos().sort(
      compareTransactionSheetInfosByStartDateDescending
    );
    if (transactionSheetInfos.length < 1)
      throw new Error("No transaction sheets found in order to add more");

    const latestTransactionSheetInfo = transactionSheetInfos[0];

    const doesNeedNewTransactionSheet =
      new Date() > latestTransactionSheetInfo.endDate;

    // Stop if we don't need to make any new transaction sheets
    if (!doesNeedNewTransactionSheet) return false;

    // Add 12 hours to get to the next day mid-day to avoid problems with daylight saving time
    const newStartDate = new Date(latestTransactionSheetInfo.endDate);
    newStartDate.setHours(newStartDate.getHours() + 12);
    // Add pay period days minus one plus 12 hours to get to the mid-day of the last day to
    // avoid problems with daylight saving time
    const newEndDate = new Date(latestTransactionSheetInfo.endDate);
    newEndDate.setDate(
      newEndDate.getDate() + Variables.getVariables().PayPeriodDays - 1
    );
    newEndDate.setHours(newEndDate.getHours() + 12);

    // Determine if this is a gap period (for now, strictly by 28 pay periods per year)
    // TODO: Improve gap period calculation?
    const isGapPeriod =
      transactionSheetInfos.length >= 14 &&
      transactionSheetInfos[13].sheet.getName().includes("(Gap)");

    /** New transaction sheet name MM/DD/YY - MM/DD/YY (Gap) */
    const newSheetName = `${
      newStartDate.getMonth() + 1
    }/${newStartDate.getDate()}/${newStartDate
      .getFullYear()
      .toString()
      .substring(2, 4)} - ${
      newEndDate.getMonth() + 1
    }/${newEndDate.getDate()}/${newEndDate
      .getFullYear()
      .toString()
      .substring(2, 4)}${isGapPeriod ? " (Gap)" : ""}`;

    // Duplicate the template and update its name
    const spreadsheet = SpreadsheetApp.getActiveSpreadsheet();
    const templateSheet = spreadsheet.getSheetByName(
      Variables.getVariables().TemplateName
    );
    if (!templateSheet)
      throw new Error(
        `${
          Variables.getVariables().TemplateName
        } sheet not found! Cannot duplicate to create sheet ${newSheetName}.`
      );

    const newSheet = spreadsheet.insertSheet(
      newSheetName,
      // Insert the sheet before the latest transaction sheet
      latestTransactionSheetInfo.sheet.getIndex() - 1,
      {
        template: templateSheet,
      }
    );

    // Color the tab if it's a gap period to indicate it needs review
    if (isGapPeriod) newSheet.setTabColor("#E8A9CA");
    return true;
  }

  /**
   * Adds new transaction sheets as needed until caught up to current date
   *
   * @returns number of sheets added
   */
  export function addTransactionSheets() {
    let numSheetsAdded = 0;
    while (addTransactionSheet()) {
      numSheetsAdded += 1;
    }
    return numSheetsAdded;
  }

  /**
   * Record receipt info in the given sheet
   *
   * Note: This does NOT check the receipt infos to make sure they should go on the given sheet
   *
   * @param sheet
   * @param receiptInfos
   * @returns number of receipts added to the sheet
   */
  function recordReceiptsOnSheet(
    sheet: GoogleAppsScript.Spreadsheet.Sheet,
    receiptInfos: ReceiptInfoBase[]
  ) {
    // Get a range of enough empty transaction rows to put the receipts in
    const transactionsRange = getTransactionsRange(sheet);
    const transactionsMax = Variables.getSheetVariables(sheet).TransactionsMax;
    const transactionsValues = transactionsRange.getValues();

    // Get first row after last non-empty transaction row index from first transaction row on this sheet
    let firstEmptyTransactionIndex = transactionsMax - 1;
    // Backtrack up the transaction rows until we find a row with something in it
    while (
      firstEmptyTransactionIndex >= 0 &&
      transactionsValues[firstEmptyTransactionIndex].every(
        (value) => value === ""
      )
    )
      // This row is filled, so try the previous row until there are no more available rows
      firstEmptyTransactionIndex -= 1;
    // Add one to go to the first empty row
    firstEmptyTransactionIndex += 1;

    if (
      firstEmptyTransactionIndex >= 0 &&
      firstEmptyTransactionIndex + receiptInfos.length > transactionsMax
    )
      // We didn't find enough empty rows. Throw
      throw new Error(
        `There are not enough empty transaction rows in sheet ${sheet.getName()} to record ${
          receiptInfos.length
        } receipts! First empty transaction row: ${
          transactionsRange.getRow() + firstEmptyTransactionIndex
        }. TransactionsMax: ${transactionsMax}. Last transaction row: ${transactionsRange.getLastRow()}`
      );

    const range = transactionsRange.offset(
      firstEmptyTransactionIndex,
      0,
      receiptInfos.length,
      3
    );

    // Record the receipt info in the range
    range.setValues(
      receiptInfos.map((receiptInfo) => [
        receiptInfo.date,
        receiptInfo.name,
        receiptInfo.cost,
      ])
    );

    // Mark receipt errors and notes
    receiptInfos.forEach((receiptInfo, i) => {
      // Mark empty cost
      if (!receiptInfo.cost) {
        // Mark the cost cell that there was no cost
        const costCell = range.getCell(i + 1, 3);
        costCell.setBackground("#E8A9CA");
      }

      if (!receiptInfo.errorMessage && !receiptInfo.note) return;

      const receiptHasError = !!receiptInfo.errorMessage;
      const combinedNote = `${receiptInfo.errorMessage}${
        receiptInfo.errorMessage && receiptInfo.note
          ? "\n\n>>>>>>>>>> NOTES <<<<<<<<<<\n\n"
          : ""
      }${receiptInfo.note}`;

      // Mark the name cell with the information from the note
      const nameCell = range.getCell(i + 1, 2);
      nameCell.setNote(combinedNote);
      nameCell.setBackground(receiptHasError ? "#FF0000" : "#E8A9CA");
    });

    return receiptInfos.length;
  }

  /**
   * Record receipt info in the spreadsheet in the active sheet
   *
   * @param threadList threads and receipt info for emails to record
   * @param shouldMarkProcessed set to true to mark the threads as processed by these scripts. Defaults to false
   * @returns information about how many receipts were recorded into which sheet
   */
  function recordReceipts(
    threadList: ThreadList,
    shouldMarkProcessed: boolean = false
  ) {
    const receiptInfos = threadList.getAllReceiptInfos();
    /** Map from transaction sheet name to number of receipts added on that sheet */
    const receiptInfosAdded: { [transactionSheetName: string]: number } = {};

    if (receiptInfos.length > 0) {
      const transactionSheetInfos = getTransactionSheetInfos();
      // Group receiptInfos by transaction sheet
      receiptInfos.forEach((receiptInfo) => {
        const transactionSheetInfo = transactionSheetInfos.find(
          (tSheetInfo) =>
            receiptInfo.date >= tSheetInfo.startDate &&
            receiptInfo.date < tSheetInfo.endDate
        );
        if (transactionSheetInfo) {
          transactionSheetInfo.receiptInfosToAdd.push(receiptInfo);
        } else
          throw new Error(
            `Could not find transaction sheet for receipt ${JSON.stringify(
              receiptInfo
            )}`
          );
      });

      transactionSheetInfos.forEach((transactionSheetInfo) => {
        if (transactionSheetInfo.receiptInfosToAdd.length > 0)
          receiptInfosAdded[transactionSheetInfo.sheet.getName()] =
            recordReceiptsOnSheet(
              transactionSheetInfo.sheet,
              transactionSheetInfo.receiptInfosToAdd
            );
      });
    }

    let numErrors = 0;
    // Mark the receipt threads processed and report errors
    threadList.threadInfos.forEach((threadInfo) => {
      // Report thread error - do not mark processed
      if (threadInfo.errors.length > 0) {
        numErrors += threadInfo.errors.length;
        Logger.log(threadInfo.errors);
        return;
      }

      // Mark receipt thread processed if no error
      if (shouldMarkProcessed) markThreadProcessed(threadInfo.thread);
      else
        Logger.log(`Would mark thread ${threadInfo.thread.getId()} processed`);
    });

    if (numErrors > 0) {
      SpreadsheetApp.getUi().alert(
        `There were ${numErrors} errors while processing. Please review.`
      );
    }

    return receiptInfosAdded;
  }

  /**
   * Get receipt info for all unprocessed emails in the inbox and record them in the spreadsheet in the active sheet
   *
   * @param start index of starting thread in the query to get email threads. Set this and max to null to get all. Defaults to 0
   * @param max max number of threads from which to get receipts. Set this and start to null to get all. Defaults to 10
   * @param shouldMarkProcessed set to true to mark the threads as processed by these scripts. Defaults to false
   * @returns information about how many receipts were recorded into which sheet
   */
  export function getAndRecordReceipts(
    start: number | null = 0,
    max: number | null = 10,
    shouldMarkProcessed = false
  ) {
    const threadList = getChaseReceipts(start, max);
    Logger.log(JSON.stringify(threadList));
    return recordReceipts(threadList, shouldMarkProcessed);
  }
}
