/**
 * Script-lock helpers. Every execution that writes transactions or records receipts (web API
 * mutations, the scheduled/menu receipt imports, and the in-sheet split checkbox) runs under the
 * same script lock, so they cannot interleave (e.g. two imports recording the same threads twice).
 */

/** How long existing entry points (menu, triggers) wait for the lock before failing loudly */
export const ENTRY_POINT_LOCK_WAIT_MS = 30000;

/** How long web API mutations try for the lock before reporting `busy` */
export const API_LOCK_TRY_MS = 20000;

/**
 * Runs `fn` while holding the script lock, waiting up to `timeoutMs` for it. Throws (like
 * `LockService`'s `waitLock`) if the lock cannot be obtained, so time-driven runs fail loudly
 * (Apps Script emails failures) instead of silently skipping. Always releases the lock.
 */
export function withScriptLock<T>(
  fn: () => T,
  timeoutMs = ENTRY_POINT_LOCK_WAIT_MS,
): T {
  const lock = LockService.getScriptLock();
  lock.waitLock(timeoutMs);
  try {
    return fn();
  } finally {
    lock.releaseLock();
  }
}

/**
 * Runs `fn` while holding the script lock if it can be obtained within `timeoutMs`. Calls
 * `onBusy` instead (without running `fn`) if not. Always releases the lock.
 */
export function tryWithScriptLock<T>(
  fn: () => T,
  onBusy: () => T,
  timeoutMs = API_LOCK_TRY_MS,
): T {
  const lock = LockService.getScriptLock();
  if (!lock.tryLock(timeoutMs)) return onBusy();
  try {
    return fn();
  } finally {
    lock.releaseLock();
  }
}

export const Lock = {
  withScriptLock,
  tryWithScriptLock,
};
