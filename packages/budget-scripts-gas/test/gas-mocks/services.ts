/**
 * Minimal in-memory mocks of PropertiesService, LockService, ContentService, and Utilities.
 * Only what src/ calls is implemented.
 */

export class MockProperties {
  readonly store = new Map<string, string>();
  getProperty(key: string) {
    return this.store.get(key) ?? null;
  }
  setProperty(key: string, value: string) {
    this.store.set(key, String(value));
    return this;
  }
  deleteProperty(key: string) {
    this.store.delete(key);
    return this;
  }
}

export class MockPropertiesService {
  readonly scriptProperties = new MockProperties();
  getScriptProperties() {
    return this.scriptProperties;
  }
}

export class MockLock {
  /** Set true to simulate another execution holding the lock */
  contended = false;
  held = false;
  /** Log of lock calls, e.g. "tryLock:20000", "waitLock:30000", "release" */
  calls: string[] = [];
  tryLock(timeoutInMillis: number) {
    this.calls.push(`tryLock:${timeoutInMillis}`);
    if (this.contended) return false;
    this.held = true;
    return true;
  }
  waitLock(timeoutInMillis: number) {
    this.calls.push(`waitLock:${timeoutInMillis}`);
    if (this.contended)
      throw new Error(
        "Lock timeout: another process was holding the lock for too long.",
      );
    this.held = true;
  }
  releaseLock() {
    this.calls.push("release");
    this.held = false;
  }
  hasLock() {
    return this.held;
  }
}

export class MockLockService {
  readonly scriptLock = new MockLock();
  getScriptLock() {
    return this.scriptLock;
  }
}

export class MockCache {
  readonly store = new Map<string, string>();
  /** Expiration in seconds passed to the last `put` of each key (entries never expire in the mock) */
  readonly expirations = new Map<string, number | undefined>();
  get(key: string) {
    return this.store.get(key) ?? null;
  }
  put(key: string, value: string, expirationInSeconds?: number) {
    this.store.set(key, String(value));
    this.expirations.set(key, expirationInSeconds);
  }
}

export class MockCacheService {
  readonly scriptCache = new MockCache();
  getScriptCache() {
    return this.scriptCache;
  }
}

export class MockTextOutput {
  mimeType: string | undefined;
  constructor(private readonly content: string) {}
  getContent() {
    return this.content;
  }
  setMimeType(mimeType: string) {
    this.mimeType = mimeType;
    return this;
  }
  getMimeType() {
    return this.mimeType;
  }
}

export class MockContentService {
  readonly MimeType = { JSON: "JSON", TEXT: "TEXT" } as const;
  createTextOutput(content = "") {
    return new MockTextOutput(content);
  }
}

let uuidCounter = 0;

export class MockUtilities {
  getUuid() {
    uuidCounter += 1;
    return `00000000-0000-4000-8000-${String(uuidCounter).padStart(12, "0")}`;
  }
  /** Supports only yyyy-MM-dd. Lenient like Java's parser: 2024-02-30 rolls over to March 1 */
  parseDate(text: string, timeZone: string, format: string) {
    if (format !== "yyyy-MM-dd")
      throw new Error(`MockUtilities.parseDate: unsupported format ${format}`);
    const m = /^(\d{4})-(\d{2})-(\d{2})$/.exec(text);
    if (!m) throw new Error(`Unparseable date: "${text}"`);
    const utcGuess = Date.UTC(Number(m[1]), Number(m[2]) - 1, Number(m[3]));
    // Offset of the time zone at that instant: format it there and read it back as UTC
    const offsetAt = (time: number) => {
      const parts = Object.fromEntries(
        new Intl.DateTimeFormat("en-US", {
          timeZone,
          hourCycle: "h23",
          year: "numeric",
          month: "2-digit",
          day: "2-digit",
          hour: "2-digit",
          minute: "2-digit",
          second: "2-digit",
        })
          .formatToParts(new Date(time))
          .map((p) => [p.type, p.value]),
      );
      const local = Date.UTC(
        Number(parts.year),
        Number(parts.month) - 1,
        Number(parts.day),
        Number(parts.hour),
        Number(parts.minute),
        Number(parts.second),
      );
      return local - time;
    };
    const first = utcGuess - offsetAt(utcGuess);
    return new Date(utcGuess - offsetAt(first));
  }
  /** Supports only the patterns src/ uses */
  formatDate(date: Date, timeZone: string, format: string) {
    if (format !== "yyyy-MM-dd")
      throw new Error(`MockUtilities.formatDate: unsupported format ${format}`);
    // en-CA formats as yyyy-MM-dd
    return new Intl.DateTimeFormat("en-CA", {
      timeZone,
      year: "numeric",
      month: "2-digit",
      day: "2-digit",
    }).format(date);
  }
}
