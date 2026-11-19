/**
 * A random id for idempotency keys. `crypto.randomUUID()` exists only in secure contexts, so this
 * falls back to `getRandomValues` (available everywhere) on plain http.
 */
export function randomId(): string {
  const c = globalThis.crypto;
  if (typeof c.randomUUID === "function") return c.randomUUID();
  const bytes = c.getRandomValues(new Uint8Array(16));
  return Array.from(bytes, (b) => b.toString(16).padStart(2, "0")).join("");
}
