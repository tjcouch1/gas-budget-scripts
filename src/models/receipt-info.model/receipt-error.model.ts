namespace Budgeting {
  /** Receipt information from an email */
  export class EmptyReceipt extends Budgeting.ReceiptInfoBase {
    #thread: GoogleAppsScript.Gmail.GmailThread;
    /**
     *
     * @param date
     * @param thread
     */
    constructor(
      date: GoogleAppsScript.Base.Date,
      thread: GoogleAppsScript.Gmail.GmailThread
    ) {
      super();
      this.date = date;
      this.#thread = thread;
    }

    public get thread() {
      return this.#thread;
    }

    toJSON() {
      return {
        ...this,
        threadId: this.thread.getId(),
      };
    }
  }
}
