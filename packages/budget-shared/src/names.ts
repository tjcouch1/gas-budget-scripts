/**
 * Split-name helpers. These mirror `splitTransaction` in
 * packages/budget-scripts-gas/src/util/budgeting.util.ts (see AGENTS.md section 3.3).
 */

const SPLIT_SEPARATOR = " | ";

/** Parent name after a split: " | " is appended only if the name has no "|" yet. */
export function buildSplitParentName(parentName: string): string {
  return `${parentName}${parentName.includes("|") ? "" : SPLIT_SEPARATOR}`;
}

/** The part of a name before the first " | " (the merchant). */
export function getSplitNamePrefix(name: string): string {
  return name.split(SPLIT_SEPARATOR)[0];
}

/** Child name: "<parent name before ' | '> | <item note>". */
export function buildSplitChildName(
  parentName: string,
  itemNote: string,
): string {
  return `${getSplitNamePrefix(parentName)}${SPLIT_SEPARATOR}${itemNote}`;
}

/**
 * True when free text would be interpreted as a formula by Sheets (starts with
 * `=`, `+`, `-`, or `@` after left-trimming). Such text must be rejected
 * server-side (AGENTS.md section 1.4).
 */
export function isUnsafeFreeText(text: string): boolean {
  return /^\s*[=+\-@]/.test(text);
}
