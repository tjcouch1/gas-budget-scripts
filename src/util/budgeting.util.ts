namespace Budgeting {
  /** Chase Gmail search query for receipt email threads */
  const chaseGmailSearchQuery =
    "from:(no.reply.alerts@chase.com) in:inbox NOT label:receipts NOT label:receipts-cru-reimburse NOT label:receipts-tax-deductible NOT label:receipts-scripted";

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
   * Record receipt info in the spreadsheet in the active sheet
   *
   * @param threadList threads and receipt info for emails to record
   * @param shouldMarkProcessed set to true to mark the threads as processed by these scripts. Defaults to false
   */
  function recordReceipts(
    threadList: ThreadList,
    shouldMarkProcessed: boolean = false
  ) {
    const sheet = SpreadsheetApp.getActiveSheet();

    const receiptInfos = threadList.receiptInfos;

    if (receiptInfos.length > 0) {
      // Get a range after the end of the contents of the sheet
      const range = sheet.getRange(
        sheet.getLastRow() + 1,
        1,
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
  }

  /**
   * Get receipt info for all unprocessed emails in the inbox and record them in the spreadsheet in the active sheet
   *
   * @param start index of starting thread in the query to get email threads. Set this and max to null to get all. Defaults to 0
   * @param max max number of threads from which to get receipts. Set this and start to null to get all. Defaults to 10
   * @param shouldMarkProcessed set to true to mark the threads as processed by these scripts. Defaults to false
   */
  export function getAndRecordReceipts(
    start: number | null = 0,
    max: number | null = 10,
    shouldMarkProcessed = false
  ) {
    const threadList = getChaseReceipts(start, max);
    Logger.log(JSON.stringify(threadList));
    recordReceipts(threadList, shouldMarkProcessed);
  }
}
