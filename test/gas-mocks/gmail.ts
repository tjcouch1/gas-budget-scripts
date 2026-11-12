/** Minimal in-memory mock of the parts of GmailApp used by src/. */

export class MockLabel {
  constructor(readonly name: string) {}
  getName() {
    return this.name;
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
    id?: string
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

  // eslint-disable-next-line @typescript-eslint/no-explicit-any
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
  getThreadById(id: string) {
    return this.threads.find((t) => t.getId() === id) ?? null;
  }
}
