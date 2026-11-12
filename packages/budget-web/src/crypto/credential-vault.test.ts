import { describe, expect, it } from "vitest";
import { CredentialVault, validateGasUrl } from "./credential-vault.ts";
import { MemoryKeyValueStore } from "./kv-store.ts";

const URL = "https://script.google.com/macros/s/TEST_DEPLOYMENT-id_1/exec";
const KEY = "super-secret-api-key-value";

interface Blob {
  iv: Uint8Array;
  ciphertext: ArrayBuffer;
}

describe("CredentialVault", () => {
  it("round-trips credentials", async () => {
    const vault = new CredentialVault(new MemoryKeyValueStore());
    expect(await vault.hasCredentials()).toBe(false);
    await vault.save({ gasWebAppUrl: URL, apiKey: KEY });
    expect(await vault.hasCredentials()).toBe(true);
    expect(await vault.load()).toEqual({ gasWebAppUrl: URL, apiKey: KEY });
  });

  it("persists only ciphertext and a non-extractable key", async () => {
    const store = new MemoryKeyValueStore();
    await new CredentialVault(store).save({ gasWebAppUrl: URL, apiKey: KEY });
    const blob = store.data.get("credential-blob") as Blob;
    const bytes = new TextDecoder("latin1").decode(blob.ciphertext);
    expect(bytes).not.toContain(KEY);
    expect(bytes).not.toContain("script.google.com");
    expect(Object.keys(blob).sort()).toEqual(["ciphertext", "iv"]);
    expect(blob.iv.byteLength).toBe(12);

    const key = store.data.get("credential-key") as CryptoKey;
    expect(key.extractable).toBe(false);
    expect(key.algorithm).toMatchObject({ name: "AES-GCM", length: 256 });
    await expect(crypto.subtle.exportKey("raw", key)).rejects.toThrow();
  });

  it("uses a fresh IV for each save", async () => {
    const store = new MemoryKeyValueStore();
    const vault = new CredentialVault(store);
    await vault.save({ gasWebAppUrl: URL, apiKey: KEY });
    const first = Array.from((store.data.get("credential-blob") as Blob).iv);
    await vault.save({ gasWebAppUrl: URL, apiKey: KEY });
    const second = Array.from((store.data.get("credential-blob") as Blob).iv);
    expect(first).not.toEqual(second);
  });

  it("rejects invalid URLs and empty keys", async () => {
    const vault = new CredentialVault(new MemoryKeyValueStore());
    await expect(
      vault.save({
        gasWebAppUrl: "http://script.google.com/macros/s/x/exec",
        apiKey: KEY,
      }),
    ).rejects.toThrow(/URL/);
    await expect(
      vault.save({
        gasWebAppUrl: "https://evil.example.com/macros/s/x/exec",
        apiKey: KEY,
      }),
    ).rejects.toThrow(/URL/);
    await expect(
      vault.save({ gasWebAppUrl: URL, apiKey: "  " }),
    ).rejects.toThrow(/key/);
    expect(await vault.hasCredentials()).toBe(false);
  });

  it("returns null when the ciphertext is tampered with", async () => {
    const store = new MemoryKeyValueStore();
    const vault = new CredentialVault(store);
    await vault.save({ gasWebAppUrl: URL, apiKey: KEY });
    const blob = store.data.get("credential-blob") as Blob;
    new Uint8Array(blob.ciphertext)[0] ^= 0xff;
    expect(await vault.load()).toBeNull();
  });

  it("forget removes both the key and the ciphertext", async () => {
    const store = new MemoryKeyValueStore();
    const vault = new CredentialVault(store);
    await vault.save({ gasWebAppUrl: URL, apiKey: KEY });
    await vault.forget();
    expect(store.data.size).toBe(0);
    expect(await vault.load()).toBeNull();
  });

  it("validates the deployment URL shape", () => {
    expect(validateGasUrl(URL)).toBe(true);
    expect(validateGasUrl(`${URL}?key=1`)).toBe(false);
    expect(validateGasUrl("https://script.google.com/macros/s/abc/dev")).toBe(
      false,
    );
  });

  it("accepts the local demo backend URL in dev builds only", () => {
    // Vitest runs with import.meta.env.DEV = true, like `vite`; `vite build` has it false
    expect(import.meta.env.DEV).toBe(true);
    expect(validateGasUrl("http://localhost:8787/macros/s/demo/exec")).toBe(
      true,
    );
    expect(validateGasUrl("http://127.0.0.1:8787/macros/s/demo/exec")).toBe(
      true,
    );
    expect(validateGasUrl("http://evil.example.com/macros/s/demo/exec")).toBe(
      false,
    );
    expect(validateGasUrl("http://localhost:8787/other/exec")).toBe(false);
  });
});
