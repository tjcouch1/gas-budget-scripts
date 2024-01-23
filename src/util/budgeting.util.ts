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
   * Get range on sheet starting at top left with the provided number of columns and the
   * number of max transactions rows
   * @param sheet
   * @param a1NotationTopLeft top left in A1 notation
   * @param numColumns number of columns in the range
   * @returns range at top left with numColumns and max transaction rows
   */
  function getTransactionSizeRange(
    sheet: GoogleAppsScript.Spreadsheet.Sheet,
    a1NotationTopLeft: string,
    numColumns: number
  ) {
    // Get top left of range
    const topLeftRange = sheet.getRange(a1NotationTopLeft);
    // Expand to size of all transactions
    const transactionSizeRange = topLeftRange.offset(
      0,
      0,
      Variables.getSheetVariables(sheet).TransactionsMax,
      numColumns
    );
    return transactionSizeRange;
  }

  /**
   * Get full range of all transaction rows in a sheet. Contains only date, name, and cost columns
   * @param sheet
   */
  function getTransactionsRange(sheet: GoogleAppsScript.Spreadsheet.Sheet) {
    // Get transactions on sheet with date, name, cost columns
    return getTransactionSizeRange(
      sheet,
      Variables.getSheetVariables(sheet).TransactionsStart,
      3
    );
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
    if (Util.areDatesEqual(a.startDate, b.startDate)) return 0;
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
   * Get the next open transaction row range (with number of rows of `transactionCount`)
   * on the given sheet - date, name, cost
   *
   * First transaction row in the returned range is the first empty row after all currently
   * filled transaction rows. Skips empty rows between filled rows.
   *
   * @param sheet
   * @param transactionsCount number of transactions to get in the range (number of rows)
   */
  function getNextOpenTransactionRange(
    sheet: GoogleAppsScript.Spreadsheet.Sheet,
    transactionsCount: number
  ) {
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
      firstEmptyTransactionIndex + transactionsCount > transactionsMax
    )
      // We didn't find enough empty rows. Throw
      throw new Error(
        `There are not enough empty transaction rows in sheet ${sheet.getName()} to get ${transactionsCount} transaction rows! First empty transaction row: ${
          transactionsRange.getRow() + firstEmptyTransactionIndex
        }. TransactionsMax: ${transactionsMax}. Last transaction row: ${transactionsRange.getLastRow()}`
      );

    return transactionsRange.offset(
      firstEmptyTransactionIndex,
      0,
      transactionsCount,
      3
    );
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
    const range = getNextOpenTransactionRange(sheet, receiptInfos.length);

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

  /** Get range of all checkboxes that split a transaction into two. Contains one column */
  export function getTransactionSplitCheckboxesRange(
    sheet: GoogleAppsScript.Spreadsheet.Sheet
  ) {
    // Get checkboxes for each transaction
    return getTransactionSizeRange(
      sheet,
      Variables.getSheetVariables(sheet).SplitCheckboxesStart,
      1
    );
  }

  /** Information about a transaction row derived from a transaction row range. Could be blank */
  type TransactionRowInfo = {
    date: GoogleAppsScript.Base.Date | undefined;
    name: string | undefined;
    cost: string | number | undefined;
    category: string | undefined;
    type: string | undefined;
    /** The range for the date, name, and cost of this transaction row */
    range: GoogleAppsScript.Spreadsheet.Range;
    /** The range for the category and type */
    metadataRange: GoogleAppsScript.Spreadsheet.Range;
  };

  /**
   * Returns a range of the transaction metadata rows corresponding to the rows in the current range of transaction rows
   * @param range current transaction rows range to get metadata range for
   * @returns range with same number of rows as original transaction rows range offset to contain the metadata for those
   * transactions - category, type
   */
  function getTransactionMetadataRange(
    range: GoogleAppsScript.Spreadsheet.Range
  ) {
    return range.offset(0, 4, range.getNumRows(), 2);
  }

  /**
   * Get transaction information from a transaction row range
   * @param range transaction range - date, name, cost
   * @returns
   */
  function getTransactionRowInfo(
    range: GoogleAppsScript.Spreadsheet.Range
  ): TransactionRowInfo {
    const transactionValues = range.getValues();
    const transactionFormulas = range.getFormulas();
    // Get formula or value of each transaction cell
    const [date, name, cost] = transactionFormulas[0].map(
      (formula, i) => formula || transactionValues[0][i]
    );

    // Get the metadata range
    const metadataRange = getTransactionMetadataRange(range);
    const [[category, type]] = metadataRange.getValues();

    return { date, name, cost, category, type, range, metadataRange };
  }

  /**
   * Determine if two transaction rows seem to be equal based on their info.
   *
   * Checks dates first and then also checks names (Can use | to match just the stuff before the bar)
   */
  function doTransactionRowInfosSeemEqual(
    transactionInfo1: TransactionRowInfo,
    transactionInfo2: TransactionRowInfo
  ): boolean {
    // Check dates equal
    if (!Util.areDatesEqual(transactionInfo1.date, transactionInfo2.date))
      return false;

    // Check names or prefixes (before bar|) equal (or at least first 1/3 of the names equal)
    if (transactionInfo1.name !== transactionInfo2.name) {
      if (!transactionInfo1.name || !transactionInfo2.name) return false;

      let nameEqual = false;

      const prefix1 = transactionInfo1.name.split("|")[0].trimEnd();
      const prefix2 = transactionInfo2.name.split("|")[0].trimEnd();

      if (prefix1 === prefix2) nameEqual = true;

      if (!nameEqual) {
        const shorterThirdLength = Math.floor(
          Math.min(
            transactionInfo1.name.length / 3,
            transactionInfo2.name.length / 3
          )
        );

        if (
          transactionInfo1.name.substring(0, shorterThirdLength) ===
          transactionInfo2.name.substring(0, shorterThirdLength)
        )
          nameEqual = true;
      }

      if (!nameEqual) return false;
    }

    return true;
  }

  /**
   * Gets a 2D array of `null`s corresponding to the size of the provided range
   * @param range
   */
  function getNullRangeValues(range: GoogleAppsScript.Spreadsheet.Range) {
    return Array.from({ length: range.getNumRows() }, () =>
      Array.from(
        {
          length: range.getNumColumns(),
        },
        () => null
      )
    );
  }

  /**
   * Splits a transaction into two rows in the provided transaction sheet.
   *
   * If the checked transaction is already split, it splits it again
   *
   * @param sheet sheet on which to split transaction
   * @param transactionIndex index of transaction to split relative to first transaction (0)
   * @returns range containing new transaction rows
   */
  export function splitTransaction(
    sheet: GoogleAppsScript.Spreadsheet.Sheet,
    transactionIndex: number
  ) {
    // Get range at transaction index
    /** Top left cell of the transactions */
    const topLeftRange = sheet.getRange(
      Variables.getSheetVariables(sheet).TransactionsStart
    );

    // Get range of all transactions in this group of split transactions
    // Assume all split transactions are immediately below the transaction to split
    const transactionsMax = Variables.getSheetVariables(sheet).TransactionsMax;
    /**
     * Array of all transactions in this group of split transactions.
     *
     * First row is transaction row to split in two
     */
    let transactionsInGroup: TransactionRowInfo[] = [];
    while (transactionIndex + transactionsInGroup.length < transactionsMax) {
      // Check the next row
      const nextTransactionRange = topLeftRange.offset(
        transactionIndex + transactionsInGroup.length,
        0,
        1,
        3
      );
      const nextTransactionRangeInfo =
        getTransactionRowInfo(nextTransactionRange);

      if (
        transactionsInGroup.length === 0 &&
        !nextTransactionRangeInfo.date &&
        !nextTransactionRangeInfo.name &&
        !nextTransactionRangeInfo.cost &&
        !nextTransactionRangeInfo.category &&
        !nextTransactionRangeInfo.type
      )
        throw new Error(
          `Checked row ${nextTransactionRange.getRow()} on sheet ${sheet.getName()}, but the transaction row has no content!`
        );

      // If this is the first transaction (the one to split) or
      // this transaction matches the first one, add it to the group and look for more
      if (
        transactionsInGroup.length === 0 ||
        doTransactionRowInfosSeemEqual(
          nextTransactionRangeInfo,
          transactionsInGroup[0]
        )
      )
        transactionsInGroup.push(nextTransactionRangeInfo);
      // If it's not in this group, break. We have found all transactions in this group
      else break;
    }

    if (transactionsInGroup.length === 0)
      throw new Error(
        `Somehow there are no transactions to split on sheet ${sheet.getName()} transactionIndex ${transactionIndex}`
      );
    if (transactionIndex + transactionsInGroup.length >= transactionsMax)
      throw new Error(
        `No room on sheet ${sheet.getName()} to split group of ${
          transactionsInGroup.length
        } starting at row ${transactionsInGroup[0].range.getRow()}!`
      );

    // If this row to split looks like it's in the middle of a transaction group, reject
    if (transactionIndex > 0) {
      const previousTransactionRange = topLeftRange.offset(
        transactionIndex - 1,
        0,
        1,
        3
      );
      const nextTransactionRangeInfo = getTransactionRowInfo(
        previousTransactionRange
      );

      if (
        doTransactionRowInfosSeemEqual(
          nextTransactionRangeInfo,
          transactionsInGroup[0]
        )
      )
        throw new Error(
          `Transaction in sheet ${sheet.getName()} at row ${transactionsInGroup[0].range.getRow()} seems to be in a transaction group already!`
        );
    }

    // Check to see if the next transaction row after the transaction group is empty
    const nextTransactionRangeAfterGroup = topLeftRange.offset(
      transactionIndex + transactionsInGroup.length,
      0,
      1,
      3
    );
    const nextTransactionRangeAfterGroupInfo = getTransactionRowInfo(
      nextTransactionRangeAfterGroup
    );

    /** What the new group of transaction rows should look like */
    const transactionsInNewGroup = transactionsInGroup.map((transaction, i) => {
      // If this is the transaction to split row, prepare it to
      // subtract the new split transaction cost
      if (i === 0)
        return {
          ...transaction,
          name: `${transaction.name}${
            transaction.name?.includes("|") ? "" : " | "
          }`,
          cost: `${
            Util.isString(transaction.cost) && transaction.cost.startsWith("=")
              ? ""
              : "="
          }${transaction.cost}-R[${transactionsInGroup.length}]C[0]`,
        };
      // Otherwise just clone it and return it
      return { ...transaction };
    });
    // Add a new split transaction row with cost 0 plus tax
    transactionsInNewGroup.push({
      date: transactionsInNewGroup[0].date,
      name: transactionsInNewGroup[0].name,
      cost: `=${Variables.getVariables().TaxMultiplier}*(0)`,
      category: transactionsInNewGroup[0].category,
      type: transactionsInNewGroup[0].type,
      range: nextTransactionRangeAfterGroupInfo.range,
      metadataRange: nextTransactionRangeAfterGroupInfo.metadataRange,
    });

    // If the next transaction row after the transaction group isn't empty, move all the following transactions down one
    if (
      nextTransactionRangeAfterGroupInfo.date ||
      nextTransactionRangeAfterGroupInfo.name ||
      nextTransactionRangeAfterGroupInfo.cost ||
      nextTransactionRangeAfterGroupInfo.category ||
      nextTransactionRangeAfterGroupInfo.type
    ) {
      const nextOpenTransactionRange = getNextOpenTransactionRange(sheet, 1);

      // Get the range from the one after the transaction group to the one before the next open transaction range
      /** Transactions to move down one row */
      const transactionsToMove = nextTransactionRangeAfterGroup.offset(
        0,
        0,
        nextOpenTransactionRange.getRow() -
          nextTransactionRangeAfterGroup.getRow(),
        3
      );
      /** Transaction metadata to move down one row */
      const transactionsMetadataToMove =
        getTransactionMetadataRange(transactionsToMove);

      // Copy the transactions down one
      transactionsToMove.copyTo(transactionsToMove.offset(1, 0));
      transactionsMetadataToMove.copyTo(
        transactionsMetadataToMove.offset(1, 0)
      );

      // Clear the values out of the current row
      const nullArrayNextTransactionRangeAfterGroupSize = getNullRangeValues(
        nextTransactionRangeAfterGroupInfo.range
      );
      const nullArrayNextTransactionMetadataRangeAfterGroupSize =
        getNullRangeValues(nextTransactionRangeAfterGroupInfo.metadataRange);

      nextTransactionRangeAfterGroupInfo.range
        .clearContent()
        .setNotes(nullArrayNextTransactionRangeAfterGroupSize)
        .setBackgrounds(nullArrayNextTransactionRangeAfterGroupSize);
      nextTransactionRangeAfterGroupInfo.metadataRange
        .clearContent()
        .setNotes(nullArrayNextTransactionMetadataRangeAfterGroupSize)
        .setBackgrounds(nullArrayNextTransactionMetadataRangeAfterGroupSize);
    }

    // Add a split transaction row directly after this transaction group
    // Get the current transaction group plus the new row
    /** Range for new group of transaction rows one longer than the current group */
    const splitTransactionRange = topLeftRange.offset(
      transactionIndex,
      0,
      transactionsInNewGroup.length,
      3
    );

    // Fill in the new split transaction row
    /** New transaction info to write into the new split row */
    const newTransactionInfo =
      transactionsInNewGroup[transactionsInNewGroup.length - 1];
    nextTransactionRangeAfterGroupInfo.range.setValues([
      [
        newTransactionInfo.date,
        newTransactionInfo.name,
        newTransactionInfo.cost,
      ],
    ]);
    nextTransactionRangeAfterGroupInfo.metadataRange.setValues([
      [newTransactionInfo.category, newTransactionInfo.type],
    ]);

    // Subtract the new split transaction row from the row to split
    const newFirstTransactionInfo = transactionsInNewGroup[0];
    transactionsInGroup[0].range.setValues([
      [
        newFirstTransactionInfo.date,
        newFirstTransactionInfo.name,
        newFirstTransactionInfo.cost,
      ],
    ]);

    return splitTransactionRange;
  }
}
