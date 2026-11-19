/** Max threads the Gmail batch APIs (`addToThreads`, `moveThreadsToArchive`, ...) take per call */
export const GMAIL_BATCH_SIZE = 100;

/** Calls `fn` with consecutive chunks of at most `GMAIL_BATCH_SIZE` threads */
export function forEachThreadChunk(
  threads: GoogleAppsScript.Gmail.GmailThread[],
  fn: (chunk: GoogleAppsScript.Gmail.GmailThread[]) => void,
) {
  for (let i = 0; i < threads.length; i += GMAIL_BATCH_SIZE)
    fn(threads.slice(i, i + GMAIL_BATCH_SIZE));
}

/**
 * Gets the messages of every thread with `GmailApp.getMessagesForThreads` (one call per 100
 * threads instead of one `getMessages()` per thread). Returns one function per thread (same order)
 * that returns that thread's messages.
 *
 * If a batch call fails, the threads of that chunk fall back to `thread.getMessages()`, called when
 * their function is called, so a single unreadable thread throws only for that thread, exactly as
 * reading the threads one by one did.
 */
export function getMessagesOfThreads(
  threads: GoogleAppsScript.Gmail.GmailThread[],
): (() => GoogleAppsScript.Gmail.GmailMessage[])[] {
  const getters: (() => GoogleAppsScript.Gmail.GmailMessage[])[] = [];
  forEachThreadChunk(threads, (chunk) => {
    let batch: GoogleAppsScript.Gmail.GmailMessage[][] | undefined;
    try {
      batch = GmailApp.getMessagesForThreads(chunk);
    } catch (e) {
      Logger.log(
        `getMessagesForThreads failed; reading ${chunk.length} threads one by one. ${e}`,
      );
    }
    chunk.forEach((thread, i) =>
      getters.push(batch ? () => batch[i] : () => thread.getMessages()),
    );
  });
  return getters;
}
