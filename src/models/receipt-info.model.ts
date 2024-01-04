namespace Budgeting {
  /** Receipt information from an email */
  export class ReceiptInfo {
    /**
     *
     * @param message the gmail message from which this receipt is derived
     * @param thread the gmail thread that contains this receipt's message
     * @param date
     * @param cost
     * @param name
     * @param notes Notes on the receipt of some abnormality. May pertain to other messages in the thread
     */
    constructor(
      public message: GoogleAppsScript.Gmail.GmailMessage,
      public date: GoogleAppsScript.Base.Date,
      public cost: number | undefined,
      public name: string | undefined,
      public notes: string[]
    ) {
    }

    public get thread() {
      return this.message.getThread();
    }

    toJSON() {
      return { ...this, message: undefined, threadId: this.thread.getId() };
    }
  }
}
