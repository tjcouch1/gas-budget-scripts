import { buildSplitChildName, getSplitNamePrefix } from "@budget/shared";

const SEPARATOR = " | ";

/** True when the name belongs to a split group (it contains "|"). */
export const isGroupName = (name: string): boolean => name.includes("|");

/**
 * Splits "<merchant> | <item note>" into its two parts so only the note is edited
 * (the merchant prefix is what groups split rows; AGENTS.md section 3.3).
 */
export function splitNameParts(name: string): { prefix: string; note: string } {
  const prefix = getSplitNamePrefix(name);
  const note = name.startsWith(prefix + SEPARATOR)
    ? name.slice(prefix.length + SEPARATOR.length)
    : "";
  return { prefix, note };
}

/** The full name for a split row with the edited item note. */
export const joinNameParts = (prefix: string, note: string): string =>
  buildSplitChildName(prefix, note);
