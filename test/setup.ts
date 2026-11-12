import { beforeEach, vi } from "vitest";
import { installGasGlobals } from "./gas-mocks";

// src/ keeps module-level caches (Variables, Gmail labels), so give every test a fresh module
// registry and fresh GAS globals.
beforeEach(() => {
  vi.resetModules();
  installGasGlobals();
});
