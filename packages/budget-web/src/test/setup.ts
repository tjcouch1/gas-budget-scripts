import "@testing-library/jest-dom/vitest";
import "fake-indexeddb/auto";
import { webcrypto } from "node:crypto";
import { afterEach } from "vitest";
import { cleanup } from "@testing-library/react";

// jsdom may not expose SubtleCrypto; Node's WebCrypto is spec-compatible.
if (!globalThis.crypto?.subtle) {
  Object.defineProperty(globalThis, "crypto", {
    value: webcrypto,
    configurable: true,
  });
}

afterEach(() => cleanup());
