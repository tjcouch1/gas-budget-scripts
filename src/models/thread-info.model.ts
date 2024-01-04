namespace Budgeting {
  /** Threads retrieved for receipt processing */
  export class ThreadInfo {
    /**
     *
     * @param thread
     * @param errorMessage If this exists, there was an error while processing the thread. Do not mark as processed. Some receipts may still be logged for this thread
     */
    constructor(
      public thread: GoogleAppsScript.Gmail.GmailThread,
      public errorMessage?: string
    ) {}

    toJSON() {
      return { ...this, thread: undefined, threadId: this.thread.getId() };
    }
  }
}
