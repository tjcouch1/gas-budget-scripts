import { openDB, type IDBPDatabase } from "idb";

/** Minimal async key/value store so tests can inject an in-memory implementation. */
export interface KeyValueStore {
  get<T>(key: string): Promise<T | undefined>;
  set(key: string, value: unknown): Promise<void>;
  delete(key: string): Promise<void>;
  /** Removes every entry. */
  clear(): Promise<void>;
}

export class MemoryKeyValueStore implements KeyValueStore {
  readonly data = new Map<string, unknown>();
  get<T>(key: string): Promise<T | undefined> {
    return Promise.resolve(this.data.get(key) as T | undefined);
  }
  set(key: string, value: unknown): Promise<void> {
    this.data.set(key, value);
    return Promise.resolve();
  }
  delete(key: string): Promise<void> {
    this.data.delete(key);
    return Promise.resolve();
  }
  clear(): Promise<void> {
    this.data.clear();
    return Promise.resolve();
  }
}

const STORE_NAME = "kv";

/** IndexedDB-backed store. CryptoKey objects are structured-clonable, so they can be stored here. */
export class IdbKeyValueStore implements KeyValueStore {
  private dbPromise: Promise<IDBPDatabase> | null = null;

  constructor(private readonly dbName: string) {}

  private db(): Promise<IDBPDatabase> {
    this.dbPromise ??= openDB(this.dbName, 1, {
      upgrade(db) {
        db.createObjectStore(STORE_NAME);
      },
    });
    return this.dbPromise;
  }

  async get<T>(key: string): Promise<T | undefined> {
    return (await (await this.db()).get(STORE_NAME, key)) as T | undefined;
  }
  async set(key: string, value: unknown): Promise<void> {
    await (await this.db()).put(STORE_NAME, value, key);
  }
  async delete(key: string): Promise<void> {
    await (await this.db()).delete(STORE_NAME, key);
  }
  async clear(): Promise<void> {
    await (await this.db()).clear(STORE_NAME);
  }
}
