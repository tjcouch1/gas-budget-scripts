/** Minimal in-memory mock of the parts of GmailApp used by src/. */
import { gasCalls } from "./calls";

/** Max threads the Gmail batch APIs accept per call (real limit) */
export const GMAIL_BATCH_LIMIT = 100;

function assertBatchSize(method: string, threads: unknown[]) {
  if (threads.length > GMAIL_BATCH_LIMIT)
    throw new Error(
      `${method}: at most ${GMAIL_BATCH_LIMIT} threads per call, got ${threads.length}`,
    );
}

export class MockLabel {
  /** Sizes of each addToThreads call, to test chunking */
  addToThreadsCalls: number[] = [];
  constructor(readonly name: string) {}
  getName() {
    return this.name;
  }
  addToThreads(threads: MockThread[]) {
    gasCalls.record("GmailLabel.addToThreads");
    assertBatchSize("GmailLabel.addToThreads", threads);
    this.addToThreadsCalls.push(threads.length);
    threads.forEach((t) => t.labels.push(this));
    return this;
  }
}

export interface MessageInit {
  id?: string;
  from: string;
  to?: string;
  subject: string;
  body: string;
  /** HTML body returned by getBody() */
  html?: string;
  date?: Date;
}

let idCounter = 0;

export class MockMessage {
  thread!: MockThread;
  private readonly id: string;
  private readonly init: MessageInit;
  constructor(init: MessageInit) {
    this.init = init;
    this.id = init.id ?? `msg-${++idCounter}`;
  }
  getId() {
    return this.id;
  }
  getFrom() {
    return this.init.from;
  }
  getTo() {
    return this.init.to ?? "me@example.com";
  }
  getSubject() {
    return this.init.subject;
  }
  getPlainBody() {
    return this.init.body;
  }
  getBody() {
    return this.init.html ?? "";
  }
  getDate() {
    return this.init.date ?? new Date("2024-03-01T12:00:00Z");
  }
  getThread() {
    gasCalls.record("GmailMessage.getThread");
    return this.thread;
  }
}

export class MockThread {
  labels: MockLabel[] = [];
  archived = false;
  /** If set, getMessages throws it (simulates a thread-level failure) */
  messagesError: Error | undefined;
  private readonly id: string;
  constructor(
    readonly messages: MockMessage[],
    id?: string,
  ) {
    this.id = id ?? `thread-${++idCounter}`;
    messages.forEach((m) => (m.thread = this));
  }
  getId() {
    return this.id;
  }
  getMessages() {
    gasCalls.record("GmailThread.getMessages");
    if (this.messagesError) throw this.messagesError;
    return this.messages;
  }
  getLastMessageDate() {
    gasCalls.record("GmailThread.getLastMessageDate");
    return this.messages[this.messages.length - 1].getDate();
  }
  getLabels() {
    gasCalls.record("GmailThread.getLabels");
    return this.labels;
  }
  addLabel(label: MockLabel) {
    gasCalls.record("GmailThread.addLabel");
    this.labels.push(label);
    return this;
  }
  moveToArchive() {
    gasCalls.record("GmailThread.moveToArchive");
    this.archived = true;
    return this;
  }
}

/** Typed `any` so it can be passed where src/ expects the (much larger) GmailThread interface */
// eslint-disable-next-line @typescript-eslint/no-explicit-any
export function makeThread(...messages: MessageInit[]): any {
  return new MockThread(messages.map((m) => new MockMessage(m)));
}

/**
 * A thread with a fixed id whose messages get fixed ids (`<id>-m1`, ...) unless they have one, so
 * ids in error messages and snapshots do not depend on test order
 */
// eslint-disable-next-line @typescript-eslint/no-explicit-any
export function makeThreadWithId(id: string, ...messages: MessageInit[]): any {
  return new MockThread(
    messages.map(
      (m, i) => new MockMessage({ ...m, id: m.id ?? `${id}-m${i + 1}` }),
    ),
    id,
  );
}

export class MockGmailApp {
  labels = new Map<string, MockLabel>();
  /** Threads returned by search() */
  // eslint-disable-next-line @typescript-eslint/no-explicit-any
  threads: any[] = [];
  searches: { query: string; start?: number; max?: number }[] = [];
  createdLabels: string[] = [];

  search(query: string, start?: number, max?: number) {
    gasCalls.record("GmailApp.search");
    this.searches.push({ query, start, max });
    return start === undefined
      ? this.threads
      : this.threads.slice(start, start + (max ?? Infinity));
  }
  getUserLabelByName(name: string) {
    gasCalls.record("GmailApp.getUserLabelByName");
    return this.labels.get(name) ?? null;
  }
  createLabel(name: string) {
    gasCalls.record("GmailApp.createLabel");
    const label = new MockLabel(name);
    this.labels.set(name, label);
    this.createdLabels.push(name);
    return label;
  }
  /** Sizes of each moveThreadsToArchive call, to test chunking */
  moveThreadsToArchiveCalls: number[] = [];
  moveThreadsToArchive(threads: MockThread[]) {
    gasCalls.record("GmailApp.moveThreadsToArchive");
    assertBatchSize("GmailApp.moveThreadsToArchive", threads);
    this.moveThreadsToArchiveCalls.push(threads.length);
    threads.forEach((t) => (t.archived = true));
    return this;
  }
  /**
   * Messages of each thread, in thread order. Like one batch request: if any thread fails (see
   * `MockThread.messagesError`) the whole call throws
   */
  getMessagesForThreads(threads: MockThread[]) {
    gasCalls.record("GmailApp.getMessagesForThreads");
    assertBatchSize("GmailApp.getMessagesForThreads", threads);
    const failed = threads.find((t) => t.messagesError);
    if (failed) throw failed.messagesError;
    return threads.map((t) => [...t.messages]);
  }
  getThreadById(id: string) {
    return this.threads.find((t) => t.getId() === id) ?? null;
  }
}
