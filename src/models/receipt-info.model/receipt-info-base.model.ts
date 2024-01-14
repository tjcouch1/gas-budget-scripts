namespace Budgeting {
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
}
