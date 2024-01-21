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
   * RegExp pattern matching to Chase's gas receipt email subjects
   */
  const chaseSubjectGasRegExp = /^You used your card at a gas station$/;
  /**
   * RegExp pattern matching to Chase's refund and gas receipt email plain body
   *
   * Named groups: `name`
   */
  const chaseBodyMerchantRegExp = /\nMerchant\s+(?<name>.+)\s+\n/;

  /**
   * Compares receiptInfos by date in ascending order
   *
   * Used in [`Array.prototype.sort`](https://developer.mozilla.org/en-US/docs/Web/JavaScript/Reference/Global_Objects/Array/sort)
   */
  function compareReceiptInfosByDateAscending(
    a: ReceiptInfoBase,
    b: ReceiptInfoBase
  ) {
    if (a.date > b.date) return 1;
    if (Util.areDatesEqual(a.date, b.date)) return 0;
    return -1;
  }

  /** Information for threads retrieved and receipts found in those threads */
  export class ThreadList {
    /** Array of receipt-related thread info */
    public readonly threadInfos: ThreadInfo[] = [];

    #allReceiptInfosCache: ReceiptInfoBase[] | undefined;
    /**
     * Returns an array of all receipt email info - flat mapped from threadInfos and contains thread notes and errors baked into the receipts
     *
     * TODO: Move note and error calculation stuff into ThreadInfo?
     */
    public getAllReceiptInfos(): ReceiptInfoBase[] {
      if (this.#allReceiptInfosCache) return this.#allReceiptInfosCache;

      this.#allReceiptInfosCache = this.threadInfos.flatMap((threadInfo) => {
        const receiptInfos: ReceiptInfoBase[] = [...threadInfo.receiptInfos];

        // Add an empty receipt for an error line if the thread has no receipts
        if (receiptInfos.length === 0) {
          receiptInfos.push(
            new Budgeting.EmptyReceipt(
              threadInfo.thread.getLastMessageDate(),
              threadInfo.thread
            )
          );
        }

        // Add errors to the first receipt
        if (threadInfo.errors.length > 0)
          receiptInfos[0].errorMessage = threadInfo.errors.join(
            "\n\n~~~~~~~~~~~~~~~~~ NEXT ERROR ~~~~~~~~~~~~~~~~~\n\n"
          );

        // Add notes to the first receipt
        if (threadInfo.notes.length > 0)
          receiptInfos[0].note = threadInfo.notes.join(
            "\n\n~~~~~~~~~~~~~~~~~ NEXT NOTE ~~~~~~~~~~~~~~~~~\n\n"
          );

        return receiptInfos;
      });
      this.#allReceiptInfosCache.sort(compareReceiptInfosByDateAscending);

      return this.#allReceiptInfosCache;
    }
    /**
     *
     * @param threads array of GmailThreads from which to derive receipts
     */
    constructor(threads: GoogleAppsScript.Gmail.GmailThread[]) {
      // Map email info into receipt infos
      /** All receipt email info */
      threads.forEach((thread) => {
        const threadInfo = new Budgeting.ThreadInfo(thread);

        try {
          // Try getting receipt info from each message in the thread
          const messages = thread.getMessages();

          messages.forEach((message) => {
            const subject = message.getSubject();
            try {
              // Try to get the cost and name for the message
              let cost: number | undefined;
              let name: string | undefined;

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
                  matches = chaseBodyMerchantRegExp.exec(
                    message.getPlainBody()
                  );
                  if (matches && matches.length === 2 && matches.groups)
                    name = matches.groups.name;
                } else {
                  // Test if it is a chase gas receipt
                  matches = chaseSubjectGasRegExp.exec(subject);
                  if (matches && matches.length === 1) {
                    matches = chaseBodyMerchantRegExp.exec(
                      message.getPlainBody()
                    );
                    if (matches && matches.length === 2 && matches.groups)
                      name = matches.groups.name;
                  }
                }
              }

              if (!cost && !name) {
                // This message is not a receipt. Add a note about it
                threadInfo.notes.push(
                  `Message is not a receipt:\nSubject: ${subject}\nDate: ${message.getDate()}\nThread ID: ${thread.getId()}\n280 Chars of Plain Body:\n${message
                    .getPlainBody()
                    ?.substring(0, 280)}`
                );
              } else {
                // We have a receipt (or a blank receipt with a note). Return receiptInfo
                threadInfo.receiptInfos.push(
                  new Budgeting.ReceiptInfo(message, cost, name)
                );
              }
            } catch (e) {
              threadInfo.errors.push(
                `Error while processing message ${message.getId()} with subject ${subject} from date ${message.getDate()} on thread ${thread.getId()}. Skipping marking as processed. ${e}`
              );
            }
          });
        } catch (e) {
          threadInfo.errors.push(
            `Error while processing thread with ID ${thread.getId()}. Skipping marking as processed. ${e}`
          );
        }

        // Save the thread info if there were any receipts in the thread or there was an error
        if (threadInfo.hasInformation()) this.threadInfos.push(threadInfo);
        else if (threadInfo.notes.length > 0)
          // Log any notes we are throwing out so we know
          Logger.log(
            `Ignoring thread ${thread.getId()} that has notes but no relevant info. Notes: ${
              threadInfo.notes
            }`
          );
      });
    }
  }
}
