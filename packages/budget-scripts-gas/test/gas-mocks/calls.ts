/**
 * Counts GAS API calls made on the mocks (e.g. "Range.setNote", "GmailThread.getMessages") so tests
 * can assert how many round trips the code makes. Counts only (no call log), so the long-running
 * demo server does not grow memory. `installGasGlobals` resets it.
 */
export class CallCounter {
  private readonly counts = new Map<string, number>();

  record(name: string) {
    this.counts.set(name, (this.counts.get(name) ?? 0) + 1);
  }

  /** Number of calls of `name` since the last reset */
  get(name: string) {
    return this.counts.get(name) ?? 0;
  }

  /** All non-zero counts, sorted by name */
  toObject(): Record<string, number> {
    return Object.fromEntries(
      [...this.counts.entries()].sort(([a], [b]) => a.localeCompare(b)),
    );
  }

  reset() {
    this.counts.clear();
  }
}

/** The shared counter for the installed environment */
export const gasCalls = new CallCounter();
