namespace Budgeting {
  /** Information for threads retrieved and receipts found in those threads */
  export type ReceiptStuff = {
    threads: ReceiptThread[];
    receipts: ReceiptInfo[];
  };
}
