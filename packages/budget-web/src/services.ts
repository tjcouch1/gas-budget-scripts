import type { PayPeriodStateDto } from "@budget/shared";
import { createApiClient } from "./api/client.ts";
import {
  CredentialVault,
  type Credentials,
  validateCredentials,
} from "./crypto/credential-vault.ts";
import { IdbKeyValueStore, type KeyValueStore } from "./crypto/kv-store.ts";
import { SyncEngine } from "./sync/engine.ts";

export interface Services {
  engine: SyncEngine;
  vault: CredentialVault;
  /** Calls getState with credentials that have not been saved yet. */
  testConnection(candidate: Credentials): Promise<PayPeriodStateDto>;
  /** Called after credentials are saved or forgotten so the app can re-sync. */
  clearLocalData(): Promise<void>;
}

export function createServices(options?: {
  vaultStore?: KeyValueStore;
  cacheStore?: KeyValueStore;
  fetchImpl?: typeof fetch;
}): Services {
  const vaultStore =
    options?.vaultStore ?? new IdbKeyValueStore("budget-vault");
  const cacheStore =
    options?.cacheStore ?? new IdbKeyValueStore("budget-cache");
  const vault = new CredentialVault(vaultStore);
  const callApi = createApiClient({
    credentials: vault,
    fetchImpl: options?.fetchImpl,
  });
  const engine = new SyncEngine({ callApi, store: cacheStore });
  return {
    engine,
    vault,
    async testConnection(candidate) {
      validateCredentials(candidate);
      const trial = createApiClient({
        credentials: { load: () => Promise.resolve(candidate) },
        fetchImpl: options?.fetchImpl,
      });
      const { data } = await trial<PayPeriodStateDto>({ action: "getState" });
      return data;
    },
    clearLocalData: () => engine.clearLocalData(),
  };
}
