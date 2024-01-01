/// <reference path="../models/receipt-info.model.ts" />
/// <reference path="../models/thread-info.model.ts" />
/// <reference path="../models/thread-list.model.ts" />
import ReceiptInfo = ReceiptInfoModel.ReceiptInfo;
import ReceiptThread = ThreadInfoModel.ReceiptThread;
import ReceiptStuff = ThreadListModel.ReceiptStuff;

namespace Budgeting {
  /** Chase Gmail search query for receipt email threads */
  const chaseGmailSearchQuery =
    "from:(no.reply.alerts@chase.com) in:inbox NOT label:receipts NOT label:receipts-cru-reimburse NOT label:receipts-tax-deductible NOT label:receipts-scripted";
  /**
   * RegExp pattern matching to Chase's receipt email subjects
   *
   * Named groups: `cost`, `name`
   */
  const chaseSubjectReceiptRegExp =
    /^Your \$(?<cost>.+) transaction with (?<name>.+)$/;
  /**
   * RegExp pattern matching to Chase's refund receipt email subjects
   *
   * Named groups: `cost`
   */
  const chaseSubjectRefundRegExp =
    /^You have a \$(?<cost>.+) credit pending on your credit card$/;
  /**
   * RegExp pattern matching to Chase's refund receipt email plain body
   *
   * Named groups: `name`
   */
  const chaseBodyRefundRegExp = /\nMerchant\s+(?<name>.+)\s+\n/;

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
   * Compares receiptInfos by date in ascending order
   *
   * Used in [`Array.prototype.sort`](https://developer.mozilla.org/en-US/docs/Web/JavaScript/Reference/Global_Objects/Array/sort)
   */
  function compareReceiptInfosByDateAscending(
    a: ReceiptInfoModel.ReceiptInfo,
    b: ReceiptInfoModel.ReceiptInfo
  ) {
    if (a.date > b.date) return 1;
    if (a.date === b.date) return 0;
    return -1;
  }

  /**
   * Get receipt info for all unprocessed Chase emails in the inbox
   *
   * @param start index of starting thread in the query to get email threads. Set this and max to null to get all. Defaults to 0
   * @param max max number of threads from which to get receipts. Set this and start to null to get all. Defaults to 10
   *
   * @returns receipt stuff - threads and receipt info for all unprocessed Chase emails
   */
  function getChaseReceipts(start: number | null = 0, max: number | null = 10): ThreadListModel.ReceiptStuff {
    // Query Gmail for receipt emails
    const threads =
      start !== null && max !== null
        ? GmailApp.search(chaseGmailSearchQuery, start, max)
        : GmailApp.search(chaseGmailSearchQuery);

    /**
     * Array of receipt-related thread info - GmailThread and an error if there was trouble
     * deriving receipts from the thread
     */
    const receiptThreads: ThreadInfoModel.ReceiptThread[] = [];

    // Map email info into receipt infos
    /** All receipt email info */
    const receiptInfos = threads.flatMap((thread) => {
      /** All receipt infos for this thread */
      const threadReceiptInfos: ReceiptInfoModel.ReceiptInfo[] = [];

      /** Error message for this thread */
      let errorMessage: string | undefined;

      try {
        // Try getting receipt info from each message in the thread
        const messages = thread.getMessages();

        messages.forEach((message) => {
          // Try to get the cost and name for the message
          let cost: number | undefined;
          let name: string | undefined;
          const notes: string[] = [];

          const subject = message.getSubject();

          // Test if it is a normal chase receipt
          let matches = chaseSubjectReceiptRegExp.exec(subject);
          if (matches && matches.length === 3 && matches.groups) {
            cost = parseFloat(matches.groups.cost);
            name = matches.groups.name;
          } else {
            // Test if it is a chase return receipt
            matches = chaseSubjectRefundRegExp.exec(subject);
            if (matches && matches.length === 2 && matches.groups) {
              cost = parseFloat(matches.groups.cost) * -1;
              matches = chaseBodyRefundRegExp.exec(message.getPlainBody());
              if (matches && matches.length === 2 && matches.groups)
                name = matches.groups.name;
            }
          }

          if (!cost && !name) {
            // This message is not a receipt. Add a note about it
            notes.push(
              `Message is not a receipt:\nSubject: ${subject}\nDate: ${message.getDate()}\nThread ID: ${thread.getId()}\n140 Chars of Plain Body:\n${message
                .getPlainBody()
                ?.substring(0, 140)}`
            );

            // Add note to the previous receipt if possible
            if (threadReceiptInfos.length > 0) {
              threadReceiptInfos[threadReceiptInfos.length - 1].notes.splice(
                threadReceiptInfos[threadReceiptInfos.length - 1].notes.length,
                0,
                ...notes
              );
              Logger.log(
                `${JSON.stringify(notes)} Adding note to previous receipt`
              );
              return;
            }
            // Otherwise continue to add a new blank receipt
            Logger.log(
              `${JSON.stringify(notes)} Adding note in a blank receipt`
            );
          }

          // We have a receipt (or a blank receipt with a note). Return receiptInfo
          const date = message.getDate();
          threadReceiptInfos.push({
            date,
            name,
            cost,
            notes,
          });
        });
      } catch (e) {
        errorMessage = `Error while processing thread with ID ${thread.getId()}. Skipping marking as processed. ${e}`;
      }

      // Save the thread and error message if there were any receipts in the thread or there was an error
      if (threadReceiptInfos.length > 0 || errorMessage)
        receiptThreads.push({ thread, errorMessage });

      return threadReceiptInfos;
    });

    receiptInfos.sort(compareReceiptInfosByDateAscending);

    return { threads: receiptThreads, receipts: receiptInfos };
  }

  /**
   * Record receipt info in the spreadsheet in the active sheet
   *
   * @param receiptStuff receipt stuff - threads and receipt info for emails to record
   * @param shouldMarkProcessed set to true to mark the threads as processed by these scripts. Defaults to false
   */
  function recordReceipts(receiptStuff: ThreadListModel.ReceiptStuff, shouldMarkProcessed: boolean = false) {
    const sheet = SpreadsheetApp.getActiveSheet();

    const receiptInfos = receiptStuff.receipts;

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

      // Mark receipt notes
      receiptInfos.forEach((receiptInfo, i) => {
        if (receiptInfo.notes.length > 0) {
          const nameCell = range.getCell(i + 1, 2);

          // Mark the name cell with the information from the note
          nameCell.setNote(
            receiptInfo.notes.join(
              "\n\n~~~~~~~~~~~~~~~~~ NEXT NOTE ~~~~~~~~~~~~~~~~~\n\n"
            )
          );
          nameCell.setBackground("#E8A9CA");
        }
      });
    }

    /** Error message to show in the ui for all threads */
    let combinedErrorMessage = "";

    // Mark the receipt threads processed and report errors
    receiptStuff.threads.forEach((receiptThread) => {
      // Report thread error
      if (receiptThread.errorMessage) {
        Logger.log(receiptThread.errorMessage);
        if (!combinedErrorMessage)
          combinedErrorMessage = "Error while processing receipt emails:";
        combinedErrorMessage += `\n${receiptThread.errorMessage}`;
        return;
      }

      // Mark receipt thread processed if no error
      if (shouldMarkProcessed) markThreadProcessed(receiptThread.thread);
      else
        Logger.log(
          `Would mark thread ${receiptThread.thread.getId()} processed`
        );
    });

    if (combinedErrorMessage) {
      SpreadsheetApp.getUi().alert(combinedErrorMessage);
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
    const receiptStuff = getChaseReceipts(start, max);
    Logger.log(
      JSON.stringify({
        ...receiptStuff,
        threads: receiptStuff.threads.map((receiptThread) => ({
          threadId: receiptThread.thread.getId(),
          errorMessage: receiptThread.errorMessage,
        })),
      })
    );
    recordReceipts(receiptStuff, shouldMarkProcessed);
  }
}
