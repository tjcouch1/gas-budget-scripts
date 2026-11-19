import {
  ReceiptInfoBase,
  EmptyReceipt,
  ReceiptInfo,
} from "./receipt-info.model";
import { ThreadInfo } from "./thread-info.model";
import {
  applyReceiptRules,
  extractEmailAddress,
  joinThreadErrors,
  joinThreadNotes,
  planThreadImport,
  sortRowsByDate,
  threadErrorText,
  type ImportMessageInput,
  type CompiledReceiptRules,
  type RuleContext,
} from "@budget/shared";
import {
  getCompiledReceiptRules,
  getRuleContext,
  toRuleEmail,
} from "../util/receipt-rules.util";
import { getMessagesOfThreads } from "../util/gmail.util";

/** Information for threads retrieved and receipts found in those threads */
export class ThreadList {
  /** Array of receipt-related thread info */
  public readonly threadInfos: ThreadInfo[] = [];

  #ignoredThreads: GoogleAppsScript.Gmail.GmailThread[] = [];
  /**
   * Threads that have no receipts and no errors (e.g. statements, alerts, marketing). They are not in
   * `threadInfos`, so `getAllReceiptInfos()` does not record them, but they are labeled when marking
   * so later runs skip them (web-plan.md section 4.2).
   *
   * A getter (not an own property) so `JSON.stringify(threadList)` output is unchanged.
   */
  public get ignoredThreads(): GoogleAppsScript.Gmail.GmailThread[] {
    return this.#ignoredThreads;
  }

  #allReceiptInfosCache: ReceiptInfoBase[] | undefined;
  /**
   * Returns an array of all receipt email info - flat mapped from threadInfos and contains thread notes and errors baked into the receipts
   *
   * TODO: Move note and error calculation stuff into ThreadInfo?
   */
  public getAllReceiptInfos(): ReceiptInfoBase[] {
    if (this.#allReceiptInfosCache) return this.#allReceiptInfosCache;

    this.#allReceiptInfosCache = sortRowsByDate(
      this.threadInfos.flatMap((threadInfo) => {
        const receiptInfos: ReceiptInfoBase[] = [...threadInfo.receiptInfos];

        // Add an empty receipt for an error line if the thread has no receipts
        if (receiptInfos.length === 0) {
          receiptInfos.push(
            new EmptyReceipt(
              threadInfo.thread.getLastMessageDate(),
              threadInfo.thread,
            ),
          );
        }

        // Add errors to the first receipt
        if (threadInfo.errors.length > 0)
          receiptInfos[0].errorMessage = joinThreadErrors(threadInfo.errors);

        // Add notes to the first receipt
        if (threadInfo.notes.length > 0)
          receiptInfos[0].note = joinThreadNotes(threadInfo.notes);

        return receiptInfos;
      }),
    );

    return this.#allReceiptInfosCache;
  }
  /**
   * @param threads array of GmailThreads from which to derive receipts
   * @param rules compiled receipt rules (default: the ReceiptRules variable; throws here, before any
   *   thread is read, when it is invalid)
   * @param context forward address/name used by the rules (default: from the Variables sheet)
   */
  constructor(
    threads: GoogleAppsScript.Gmail.GmailThread[],
    rules: CompiledReceiptRules = getCompiledReceiptRules(),
    context: RuleContext = getRuleContext(),
  ) {
    // One batch read of every thread's messages (falls back to per-thread reads on failure)
    const messagesOfThreads = getMessagesOfThreads(threads);

    // Map email info into receipt infos
    /** All receipt email info */
    threads.forEach((thread, threadIndex) => {
      const threadInfo = new ThreadInfo(thread);

      try {
        const messages = messagesOfThreads[threadIndex]();
        const byId = new Map<string, GoogleAppsScript.Gmail.GmailMessage>();
        const inputs = messages.map((message): ImportMessageInput => {
          byId.set(message.getId(), message);
          try {
            const outcome = applyReceiptRules(
              rules,
              toRuleEmail(message, rules.usesHtmlBody),
              context,
            );
            return {
              messageId: message.getId(),
              subject: message.getSubject(),
              // Original case of the address, for the note text
              fromAddress: extractEmailAddress(message.getFrom()),
              date: message.getDate(),
              dateText: `${message.getDate()}`,
              plainBody: message.getPlainBody(),
              outcome,
              error: null,
            };
          } catch (e) {
            return {
              messageId: message.getId(),
              subject: message.getSubject(),
              fromAddress: "",
              date: message.getDate(),
              dateText: `${message.getDate()}`,
              plainBody: undefined,
              outcome: null,
              error: `${e}`,
            };
          }
        });
        // The same planner the web app's email preview uses
        const plan = planThreadImport({
          threadId: thread.getId(),
          lastMessageDate: () => thread.getLastMessageDate(),
          messages: inputs,
        });
        plan.rows.forEach((row) => {
          const message = row.messageId && byId.get(row.messageId);
          if (!message) return; // the empty row of an errored thread is made in getAllReceiptInfos
          threadInfo.receiptInfos.push(
            new ReceiptInfo(
              message,
              row.cost,
              row.name,
              row.category,
              row.type,
              thread,
            ),
          );
        });
        threadInfo.notes.push(...plan.notes);
        threadInfo.errors.push(...plan.errors);
      } catch (e) {
        threadInfo.errors.push(threadErrorText(thread.getId(), `${e}`));
      }

      // Save the thread info if there were any receipts in the thread or there was an error
      if (threadInfo.hasInformation()) this.threadInfos.push(threadInfo);
      else {
        // Keep the thread so it can be labeled as ignored when marking
        this.#ignoredThreads.push(thread);
        if (threadInfo.notes.length > 0)
          // Log any notes we are throwing out so we know
          Logger.log(
            `Ignoring thread ${thread.getId()} that has notes but no relevant info. Notes: ${
              threadInfo.notes
            }`,
          );
      }
    });
  }
}
