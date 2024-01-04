namespace Budgeting {
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

  /**
   * Compares receiptInfos by date in ascending order
   *
   * Used in [`Array.prototype.sort`](https://developer.mozilla.org/en-US/docs/Web/JavaScript/Reference/Global_Objects/Array/sort)
   */
  function compareReceiptInfosByDateAscending(a: ReceiptInfo, b: ReceiptInfo) {
    if (a.date > b.date) return 1;
    if (a.date === b.date) return 0;
    return -1;
  }

  /** Information for threads retrieved and receipts found in those threads */
  export class ThreadList {
    /** Array of receipt-related thread info */
    public threadInfos: ThreadInfo[] = [];
    /** All receipt email info */
    public receiptInfos: ReceiptInfo[] = [];
    /**
     *
     * @param threads array of GmailThreads from which to derive receipts
     */
    constructor(threads: GoogleAppsScript.Gmail.GmailThread[]) {
      // Map email info into receipt infos
      /** All receipt email info */
      this.receiptInfos = threads.flatMap((thread) => {
        /** All receipt infos for this thread */
        const threadReceiptInfos: ReceiptInfo[] = [];

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
                  threadReceiptInfos[threadReceiptInfos.length - 1].notes
                    .length,
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
            threadReceiptInfos.push(
              new Budgeting.ReceiptInfo(message, date, cost, name, notes)
            );
          });
        } catch (e) {
          errorMessage = `Error while processing thread with ID ${thread.getId()}. Skipping marking as processed. ${e}`;
        }

        // Save the thread and error message if there were any receipts in the thread or there was an error
        if (threadReceiptInfos.length > 0 || errorMessage)
          this.threadInfos.push(new Budgeting.ThreadInfo(thread, errorMessage));

        return threadReceiptInfos;
      });

      this.receiptInfos.sort(compareReceiptInfosByDateAscending);
    }
  }
}
