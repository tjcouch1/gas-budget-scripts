namespace Budgeting {
  // #region ReceiptInfoBase
  /** Receipt information from an email */
  export abstract class ReceiptInfoBase {
    public date: GoogleAppsScript.Base.Date;
    public cost: number | undefined;
    public name: string | undefined;
    /** Error message to display in a note on this receipt. Empty string means no error */
    public errorMessage: string = "";
    /**
     * Note to display on this receipt. Empty string means no note. If `errorMessage` is also present, it will be
     * displayed before this note.
     */
    public note: string = "";

    /** the gmail thread that contains this receipt's message */
    public abstract get thread(): GoogleAppsScript.Gmail.GmailThread;
  }
  // #endregion

  // #region ReceiptError
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
      // Get rid of errorMessage and note if they don't have any information
      return {
        ...this,
        threadId: this.thread.getId(),
        errorMessage: this.errorMessage || undefined,
        note: this.note || undefined,
      };
    }
  }
  // #endregion

  // #region ReceiptInfo
  /** Receipt information from an email */
  export class ReceiptInfo extends Budgeting.ReceiptInfoBase {
    /**
     *
     * @param message the gmail message from which this receipt is derived
     * @param cost
     * @param name
     */
    constructor(
      public message: GoogleAppsScript.Gmail.GmailMessage,
      cost: number | undefined,
      name: string | undefined
    ) {
      super();
      this.cost = cost;
      this.name = name;
      this.date = this.message.getDate();
    }

    public get thread() {
      return this.message.getThread();
    }

    toJSON() {
      return {
        ...this,
        // Get rid of Gmail objects
        message: undefined,
        messageId: this.message.getId(),
        threadId: this.thread.getId(),
        // Get rid of errorMessage and note if they don't have any information
        errorMessage: this.errorMessage || undefined,
        note: this.note || undefined,
      };
    }
  }
  // #endregion
}
