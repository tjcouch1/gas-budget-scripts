/**
 * Tooltip text for the receipt-rule fields. One place, keyed by field. The wording is copied from
 * the TSDoc of the matching type/property in packages/budget-shared/src/receipt-rules.ts (keep the
 * two in step), lightly reflowed to be self-contained.
 */
export const FIELD_HELP = {
  // ReceiptRulesConfig
  defaultTypePrefix:
    'Prefix for "recipient" types when the email was not for the forwarder. Today "TJ". 1-20 chars.',
  // SenderRules
  senderAddress:
    "Exact sender address (compared case-insensitively). Strict syntax, see STRICT_EMAIL_REGEX.",
  senderType:
    'Base of the "Paid with" type: "Credit", "Paypal", "Venmo". 1-40 chars, no "=" "+" "-" "@" start.',
  typePrefix:
    '"fixed": type is `type` verbatim. "recipient": type is "<prefix> <type>" where prefix is ForwardName if the email was forwarded by the forwarder or was sent To ForwardEmailAddress (and ForwardName is non-empty), else `defaultTypePrefix`.',
  forwardable:
    "true: emails From ForwardEmailAddress whose body says they were forwarded from this sender are parsed with these rules (subject must start with a forwarding subject prefix). false: such forwards are not receipts (they only get a note).",
  // ReceiptRule
  ruleId:
    'Unique within its sender. Shown in the web app\'s preview ("matched rule: chase-purchase"). 1-40 chars, [A-Za-z0-9_-].',
  subjectPattern:
    'Required anchor: run against the subject. The first rule (in order) whose subject matches wins. For a forwarded email the subject must start with one of `forwarding.subjectPrefixes` followed by this pattern without its leading "^". The pattern is wrapped in a group, so a top-level alternation such as `^a|b` needs the prefix for every alternative. RegExp source, e.g. "^You paid (?<name>.+) \\$(?<cost>.+)$". Named groups are what templates reference.',
  subjectFlags:
    'Any of "i", "m", "s", "u" (each at most once). Never "g", "y" or "d". Default "".',
  bodyEnabled:
    'Optional second match, run once after the subject matched. Its groups are available to templates as ${body.<group>}. For a forwarded email it runs on the original (unwrapped) body, see ForwardingConfig. When it does not match, every ${body.*} reference is "missing" (see Name).',
  bodySource: "Which body the body pattern runs on: plainBody or htmlBody.",
  bodyPattern:
    'RegExp source for the optional second match, run on the chosen body after the subject matched. Its named groups are available to templates as ${body.<group>}. RegExp source, e.g. "^You paid (?<name>.+) \\$(?<cost>.+)$".',
  bodyFlags:
    'Any of "i", "m", "s", "u" (each at most once). Never "g", "y" or "d". Default "".',
  name: "Template for the receipt name. Placeholders: ${subject.<group>} and ${body.<group>}. Any other text is literal. If ANY referenced group is missing (regex did not match, group did not participate, or body not available) the template yields nothing and the name fallback is tried the same way. If both fail, name is undefined. Omit for rules that have no name.",
  nameFallback:
    "Template tried the same way as the name when the name template yields nothing (a referenced group is missing). If both fail, name is undefined.",
  costEnabled: "Where the cost comes from. Leave off for no cost.",
  costSource: "Where the cost comes from: the subject or the body.",
  costGroup:
    'The named group whose text is the cost. The group\'s text goes through parseFloat (exactly like today, so "1,234.56" is 1). An empty or missing group gives no cost.',
  costNegate: "Multiplies the cost by -1 (refunds, money received).",
  category:
    "Envelope/category written to the sheet when this rule yields a receipt (cost or name present).",
  // ForwardingConfig
  subjectPrefixes:
    'Literal subject prefixes of forwarded emails (regex-escaped when used). 1-5 entries of 1-20 chars, each letters/digits plus an optional trailing ":" and whitespace (that word is the Gmail query\'s `subject:` token). Optional; default ["Fwd: "]. A message from ForwardEmailAddress whose subject starts with none of them is "not a receipt" (not-a-forward).',
  fromPatterns:
    "Regexes with a named group `email`, run in order on the forwarded email's whole plain body to find the original sender (\"From: PayPal <service@paypal.com>\"). The group may hold a formal address; the address is extracted. None matching (the subject has a forward prefix) => error \"Could not get 'from' email address from forwarded email with subject '<subject>'\".",
  bodyMarkers:
    'Markers that start the forwarded content. The first marker found in the plain body wins: the original body is everything from the marker on, split on "\\n", with the first `skipLines` lines dropped. None found => whole body.',
} as const;

export type FieldHelpKey = keyof typeof FIELD_HELP;
