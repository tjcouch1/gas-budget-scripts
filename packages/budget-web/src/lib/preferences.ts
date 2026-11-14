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
