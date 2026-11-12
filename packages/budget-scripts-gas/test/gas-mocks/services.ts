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
