/** Minimal in-memory mock of the parts of GmailApp used by src/. */

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
  getDate() {
    return this.init.date ?? new Date("2024-03-01T12:00:00Z");
  }
  getThread() {
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
    if (this.messagesError) throw this.messagesError;
    return this.messages;
  }
  getLastMessageDate() {
    return this.messages[this.messages.length - 1].getDate();
  }
  addLabel(label: MockLabel) {
    this.labels.push(label);
    return this;
  }
  moveToArchive() {
    this.archived = true;
    return this;
  }
}

/** Typed `any` so it can be passed where src/ expects the (much larger) GmailThread interface */
// eslint-disable-next-line @typescript-eslint/no-explicit-any
export function makeThread(...messages: MessageInit[]): any {
  return new MockThread(messages.map((m) => new MockMessage(m)));
}

export class MockGmailApp {
  labels = new Map<string, MockLabel>();
  /** Threads returned by search() */
  // eslint-disable-next-line @typescript-eslint/no-explicit-any
  threads: any[] = [];
  searches: { query: string; start?: number; max?: number }[] = [];
  createdLabels: string[] = [];

  search(query: string, start?: number, max?: number) {
    this.searches.push({ query, start, max });
    return start === undefined
      ? this.threads
      : this.threads.slice(start, start + (max ?? Infinity));
  }
  getUserLabelByName(name: string) {
    return this.labels.get(name) ?? null;
  }
  createLabel(name: string) {
    const label = new MockLabel(name);
    this.labels.set(name, label);
    this.createdLabels.push(name);
    return label;
  }
  /** Sizes of each moveThreadsToArchive call, to test chunking */
  moveThreadsToArchiveCalls: number[] = [];
  moveThreadsToArchive(threads: MockThread[]) {
    assertBatchSize("GmailApp.moveThreadsToArchive", threads);
    this.moveThreadsToArchiveCalls.push(threads.length);
    threads.forEach((t) => (t.archived = true));
    return this;
  }
  getThreadById(id: string) {
    return this.threads.find((t) => t.getId() === id) ?? null;
  }
}
