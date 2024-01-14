namespace Budgeting {
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
        message: undefined,
        messageId: this.message.getId(),
        threadId: this.thread.getId(),
      };
    }
  }
}
