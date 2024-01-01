namespace ReceiptInfoModel {
  /** Receipt information from an email */
  export type ReceiptInfo = {
    date: GoogleAppsScript.Base.Date;
    cost?: number;
    name?: string;
    /** Notes on the receipt of some abnormality. May pertain to other messages in the thread */
    notes: string[];
  };
}
