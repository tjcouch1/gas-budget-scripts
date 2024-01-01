namespace ThreadInfoModel {
  /** Threads retrieved for receipt processing */
  export type ReceiptThread = {
    thread: GoogleAppsScript.Gmail.GmailThread;
    /** If this exists, there was an error while processing the thread. Do not mark as processed. Some receipts may still be logged for this thread */
    errorMessage?: string;
  };
}