namespace ThreadListModel {
  /** Information for threads retrieved and receipts found in those threads */
  export type ReceiptStuff = {
    threads: ThreadInfoModel.ReceiptThread[];
    receipts: ReceiptInfoModel.ReceiptInfo[];
  };
}
