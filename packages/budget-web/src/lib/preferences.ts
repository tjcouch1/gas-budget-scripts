/** Per-device preference (not a secret): whether Envelope choosers start expanded. */
export const ENVELOPE_EXPANDED_KEY = "budget.envelopeExpanded";

export function getEnvelopeExpanded(): boolean {
  try {
    return localStorage.getItem(ENVELOPE_EXPANDED_KEY) === "1";
  } catch {
    return false;
  }
}

export function setEnvelopeExpanded(expanded: boolean): void {
  try {
    localStorage.setItem(ENVELOPE_EXPANDED_KEY, expanded ? "1" : "0");
  } catch {
    // storage unavailable: the preference just does not persist
  }
}

export const DETAILS_OPEN_KEY = "budget.detailsOpen";
export const PAID_WITH_EXPANDED_KEY = "budget.paidWithExpanded";

function readFlag(key: string): boolean {
  try {
    return localStorage.getItem(key) === "1";
  } catch {
    return false;
  }
}
function writeFlag(key: string, on: boolean): void {
  try {
    localStorage.setItem(key, on ? "1" : "0");
  } catch {
    // storage unavailable: the preference just does not persist
  }
}

export const getDetailsOpen = () => readFlag(DETAILS_OPEN_KEY);
export const setDetailsOpen = (open: boolean) =>
  writeFlag(DETAILS_OPEN_KEY, open);
export const getPaidWithExpanded = () => readFlag(PAID_WITH_EXPANDED_KEY);
export const setPaidWithExpanded = (expanded: boolean) =>
  writeFlag(PAID_WITH_EXPANDED_KEY, expanded);

/** Per-device recently chosen envelopes, most recent first (kept short). */
export const RECENT_ENVELOPES_KEY = "budget.recentEnvelopes";
const RECENT_ENVELOPES_MAX = 30;

export function getRecentEnvelopes(): string[] {
  try {
    const parsed: unknown = JSON.parse(
      localStorage.getItem(RECENT_ENVELOPES_KEY) ?? "[]",
    );
    return Array.isArray(parsed)
      ? parsed.filter((x): x is string => typeof x === "string")
      : [];
  } catch {
    return [];
  }
}

/** Moves the envelope to the front of the recently-used list. */
export function recordRecentEnvelope(name: string): void {
  if (!name) return;
  try {
    const next = [name, ...getRecentEnvelopes().filter((n) => n !== name)];
    localStorage.setItem(
      RECENT_ENVELOPES_KEY,
      JSON.stringify(next.slice(0, RECENT_ENVELOPES_MAX)),
    );
  } catch {
    // storage unavailable: the order just does not persist
  }
}
