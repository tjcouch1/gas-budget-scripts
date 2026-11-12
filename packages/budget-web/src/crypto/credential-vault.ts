import type { KeyValueStore } from "./kv-store.ts";

/**
 * Credential vault (AGENTS.md section 1.2, web-plan.md section 4.3).
 *
 * The connection credentials are encrypted with a non-extractable AES-GCM-256
 * CryptoKey kept in IndexedDB. Only { iv, ciphertext } is persisted. Plaintext
 * exists only in the scope of whoever calls `load()` (the API client's request
 * function); it is never held in module state. This protects copied storage
 * dumps, not script running in the page.
 */

export interface Credentials {
  gasWebAppUrl: string;
  apiKey: string;
}

/** Anything that can supply credentials for one request. */
export interface CredentialSource {
  load(): Promise<Credentials | null>;
}

interface EncryptedBlob {
  iv: Uint8Array<ArrayBuffer>;
  ciphertext: ArrayBuffer;
}

const KEY_ID = "credential-key";
const BLOB_ID = "credential-blob";

export const GAS_URL_PATTERN =
  /^https:\/\/script\.google\.com\/macros\/s\/[A-Za-z0-9_-]+\/exec$/;

export function validateGasUrl(url: string): boolean {
  return GAS_URL_PATTERN.test(url.trim());
}

export function validateCredentials(credentials: Credentials): void {
  if (!validateGasUrl(credentials.gasWebAppUrl)) {
    throw new Error(
      "The URL must look like https://script.google.com/macros/s/<deployment-id>/exec",
    );
  }
  if (!credentials.apiKey.trim()) throw new Error("The API key is required.");
}

export class CredentialVault implements CredentialSource {
  constructor(private readonly store: KeyValueStore) {}

  private async getOrCreateKey(): Promise<CryptoKey> {
    const existing = await this.store.get<CryptoKey>(KEY_ID);
    if (existing) return existing;
    const key = await crypto.subtle.generateKey(
      { name: "AES-GCM", length: 256 },
      false, // non-extractable
      ["encrypt", "decrypt"],
    );
    await this.store.set(KEY_ID, key);
    return key;
  }

  async save(credentials: Credentials): Promise<void> {
    validateCredentials(credentials);
    const key = await this.getOrCreateKey();
    const iv = crypto.getRandomValues(new Uint8Array(12));
    const plaintext = new TextEncoder().encode(
      JSON.stringify({
        gasWebAppUrl: credentials.gasWebAppUrl.trim(),
        apiKey: credentials.apiKey.trim(),
      }),
    );
    const ciphertext = await crypto.subtle.encrypt(
      { name: "AES-GCM", iv },
      key,
      plaintext,
    );
    const blob: EncryptedBlob = { iv, ciphertext };
    await this.store.set(BLOB_ID, blob);
  }

  async hasCredentials(): Promise<boolean> {
    return (await this.store.get(BLOB_ID)) !== undefined;
  }

  /** Decrypts in memory. Callers must not retain the result beyond their request. */
  async load(): Promise<Credentials | null> {
    const blob = await this.store.get<EncryptedBlob>(BLOB_ID);
    const key = await this.store.get<CryptoKey>(KEY_ID);
    if (!blob || !key) return null;
    try {
      const plaintext = await crypto.subtle.decrypt(
        { name: "AES-GCM", iv: blob.iv },
        key,
        blob.ciphertext,
      );
      const parsed = JSON.parse(
        new TextDecoder().decode(plaintext),
      ) as Partial<Credentials>;
      if (
        typeof parsed.gasWebAppUrl !== "string" ||
        typeof parsed.apiKey !== "string"
      ) {
        return null;
      }
      return { gasWebAppUrl: parsed.gasWebAppUrl, apiKey: parsed.apiKey };
    } catch {
      return null;
    }
  }

  async forget(): Promise<void> {
    await this.store.delete(BLOB_ID);
    await this.store.delete(KEY_ID);
  }
}
