namespace Budgeting {
  /** Threads retrieved for receipt processing */
  export class ThreadInfo {
    /**
     *
     * @param thread gmail thread for this receipt thread
     * @param receiptInfos the receipts associated with this thread
     * @param notes Notes of some kind of abnormality related to this thread. Things like if there was a message that was not a receipt on this thread that has messages that were receipts
     * @param errors If there are errors in this array, there were errors while processing the thread. Do not mark as processed. Some receipts may still be logged for this thread
     */
    constructor(
      public thread: GoogleAppsScript.Gmail.GmailThread,
      public receiptInfos: ReceiptInfo[] = [],
      public notes: string[] = [],
      public errors: string[] = []
    ) {}

    /**
     * Returns true if this ThreadInfo has meaningful information in it.
     *
     * We should save the thread if it has receipts or errors in it. Otherwise it is fine to throw out.
     *
     * WARNING: This means we are throwing away notes that are not associated with any receipts!
     */
    public hasInformation() {
      return this.receiptInfos.length > 0 || this.errors.length > 0;
    }

    toJSON() {
      return { ...this, thread: undefined, threadId: this.thread.getId() };
    }
  }
}
