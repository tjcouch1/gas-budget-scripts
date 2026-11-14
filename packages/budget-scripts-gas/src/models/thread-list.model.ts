import {
  ReceiptInfoBase,
  EmptyReceipt,
  ReceiptInfo,
} from "./receipt-info.model";
import { ThreadInfo } from "./thread-info.model";
import { Util } from "../util/util";
import {
  applyReceiptRules,
  extractEmailAddress,
  type CompiledReceiptRules,
  type RuleContext,
} from "@budget/shared";
import {
  getCompiledReceiptRules,
  getRuleContext,
  toRuleEmail,
} from "../util/receipt-rules.util";

/**
 * Compares receiptInfos by date in ascending order
 *
 * Used in [`Array.prototype.sort`](https://developer.mozilla.org/en-US/docs/Web/JavaScript/Reference/Global_Objects/Array/sort)
 */
function compareReceiptInfosByDateAscending(
  a: ReceiptInfoBase,
  b: ReceiptInfoBase,
) {
  if (a.date > b.date) return 1;
  if (Util.areDatesEqual(a.date, b.date)) return 0;
  return -1;
}

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

    this.#allReceiptInfosCache = this.threadInfos.flatMap((threadInfo) => {
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
        receiptInfos[0].errorMessage = threadInfo.errors.join(
          "\n\n~~~~~~~~~~~~~~~~~ NEXT ERROR ~~~~~~~~~~~~~~~~~\n\n",
        );

      // Add notes to the first receipt
      if (threadInfo.notes.length > 0)
        receiptInfos[0].note = threadInfo.notes.join(
          "\n\n~~~~~~~~~~~~~~~~~ NEXT NOTE ~~~~~~~~~~~~~~~~~\n\n",
        );

      return receiptInfos;
    });
    this.#allReceiptInfosCache.sort(compareReceiptInfosByDateAscending);

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
    // Map email info into receipt infos
    /** All receipt email info */
    threads.forEach((thread) => {
      const threadInfo = new ThreadInfo(thread);

      try {
        // Try getting receipt info from each message in the thread
        const messages = thread.getMessages();

        messages.forEach((message) => {
          try {
            const outcome = applyReceiptRules(
              rules,
              toRuleEmail(message, rules.usesHtmlBody),
              context,
            );
            // Original case of the address, for the note text
            const from = extractEmailAddress(message.getFrom());
            const { cost, name, category, type } = outcome.receipt;
            const receipt = new ReceiptInfo(
              message,
              cost,
              name,
              category,
              type,
            );

            if (!receipt.cost && !receipt.name) {
              // This message is not a receipt. Add a note about it
              threadInfo.notes.push(
                `Message is not a receipt:\nFrom: ${from}\nSubject: ${message.getSubject()}\nDate: ${message.getDate()}\nThread ID: ${thread.getId()}\n280 Chars of Plain Body:\n${message
                  .getPlainBody()
                  ?.substring(0, 280)}`,
              );
            } else {
              // We have a receipt (or a blank receipt with a note). Return receiptInfo
              threadInfo.receiptInfos.push(receipt);
            }
          } catch (e) {
            threadInfo.errors.push(
              `Error while processing message ${message.getId()} with subject ${message.getSubject()} from date ${message.getDate()} on thread ${thread.getId()}. Skipping marking as processed. ${e}`,
            );
          }
        });
      } catch (e) {
        threadInfo.errors.push(
          `Error while processing thread with ID ${thread.getId()}. Skipping marking as processed. ${e}`,
        );
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
