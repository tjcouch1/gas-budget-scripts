import type { ReceiptRulesConfig } from "./receipt-rules.ts";

/** Longest allowed `JSON.stringify(config, null, 2)` (a Sheets cell holds 50,000 characters) */
export const MAX_RECEIPT_RULES_CHARS = 40_000;

/**
 * The built-in receipt rules: exactly the parsing the script did before rules were configurable
 * (Chase, PayPal, Venmo, and forwarded PayPal/Venmo emails). Used when the `ReceiptRules` variable
 * is missing or empty, and written into the cell by the setup helper.
 */
export const DEFAULT_RECEIPT_RULES: ReceiptRulesConfig = {
  version: 1,
  defaultTypePrefix: "TJ",
  forwarding: {
    subjectPrefixes: ["Fwd: "],
    fromPatterns: [
      {
        pattern:
          "[Ff]orwarded message.*\\s*\\r?\\n\\r?\\n?\\*?From:\\*?[ \\u00a0](?<email>.+)\\r?\\n",
      },
      {
        pattern:
          "転送されたメッセージ.*\\s*\\r?\\n\\r?\\n?\\*?差出人:\\*?[ \\u00a0](?<email>.+)\\r?\\n",
      },
    ],
    bodyMarkers: [
      { marker: "---------- Forwarded message ---------", skipLines: 5 },
      { marker: "Begin forwarded message:", skipLines: 6 },
      { marker: "転送されたメッセージ:", skipLines: 6 },
    ],
  },
  senders: [
    {
      address: "no.reply.alerts@chase.com",
      type: "Credit",
      typePrefix: "fixed",
      forwardable: false,
      rules: [
        {
          id: "chase-purchase",
          subject: {
            pattern: "^You made a \\$(?<cost>.+) transaction with (?<name>.+)$",
          },
          name: "${subject.name}",
          cost: { source: "subject", group: "cost" },
        },
        {
          id: "chase-refund",
          subject: {
            pattern:
              "^You have a \\$(?<cost>.+) credit pending on your credit card$",
          },
          body: {
            source: "plainBody",
            pattern: "\\r?\\nMerchant\\s+(?<name>.*\\S)\\s*\\r?\\n",
          },
          name: "${body.name}",
          cost: { source: "subject", group: "cost", negate: true },
        },
        {
          id: "chase-gas",
          subject: { pattern: "^You used your card at a gas station$" },
          body: {
            source: "plainBody",
            pattern: "\\r?\\nMerchant\\s+(?<name>.*\\S)\\s*\\r?\\n",
          },
          name: "${body.name} (Gas)",
          category: "Gas",
        },
      ],
    },
    {
      address: "service@paypal.com",
      type: "Paypal",
      typePrefix: "recipient",
      forwardable: true,
      rules: [
        {
          id: "paypal-sent",
          subject: {
            pattern: "^You sent a \\$(?<cost>.+)[\\s\\u00a0]USD payment$",
          },
          body: {
            source: "plainBody",
            pattern:
              "You sent \\$.+ to (?<name>.+)\\r?\\n\\r?\\n.*\\r?\\n.*\\r?\\n\\r?\\n(?<details>.+)\\r?\\n",
          },
          name: "${body.name} - ${body.details}",
          cost: { source: "subject", group: "cost" },
        },
        {
          id: "paypal-received",
          subject: { pattern: "^Money is waiting for you$" },
          body: {
            source: "plainBody",
            pattern:
              "Accept your \\$(?<cost>.+)[\\s\\u00a0]USD from (?<name>.+)\\r?\\n\\r?\\n.*\\r?\\n.*\\r?\\n\\r?\\n(?<details>.+)\\r?\\n",
          },
          name: "${body.name} - ${body.details}",
          cost: { source: "body", group: "cost", negate: true },
        },
        // Money received ("<Payer> sent you $8.07 USD"). NEEDS A REAL SAMPLE: the body regex for the
        // note is a guess and optional; without a match the name is the payer from the subject.
        {
          id: "paypal-sent-you",
          subject: {
            pattern: "^(?<name>.+) sent you \\$(?<cost>.+)[\\s\\u00a0]USD$",
          },
          body: {
            source: "plainBody",
            pattern:
              "[Nn]ote from [^\\r\\n]+\\r?\\n(?:[ \\t]*\\r?\\n)*(?<details>[^\\r\\n]*\\S)",
          },
          name: "${subject.name} - ${body.details}",
          nameFallback: "${subject.name}",
          cost: { source: "subject", group: "cost", negate: true },
        },
      ],
    },
    {
      address: "venmo@venmo.com",
      type: "Venmo",
      typePrefix: "recipient",
      forwardable: true,
      rules: [
        {
          id: "venmo-paid",
          subject: { pattern: "^You paid (?<name>.+) \\$(?<cost>.+)$" },
          body: {
            source: "plainBody",
            pattern: "paid\\s*.*\\s*\\s*.*\\s*\\r?\\n(?<details>.+)\\r?\\n",
          },
          name: "${subject.name} - ${body.details}",
          nameFallback: "${subject.name}",
          cost: { source: "subject", group: "cost" },
        },
        {
          id: "venmo-charge",
          subject: {
            pattern:
              "^You completed (?<name>.+)'s \\$(?<cost>.+) charge request$",
          },
          body: {
            source: "plainBody",
            pattern: "charged\\s*.*\\s*\\s*.*\\s*\\r?\\n(?<details>.+)\\r?\\n",
          },
          name: "${subject.name} - ${body.details}",
          nameFallback: "${subject.name}",
          cost: { source: "subject", group: "cost" },
        },
        {
          id: "venmo-received",
          subject: { pattern: "^(?<name>.+) paid you \\$(?<cost>.+)$" },
          body: {
            source: "plainBody",
            pattern: "paid\\s*.*\\s*\\s*.*\\s*\\r?\\n(?<details>.+)\\r?\\n",
          },
          name: "${subject.name} - ${body.details}",
          nameFallback: "${subject.name}",
          cost: { source: "subject", group: "cost", negate: true },
        },
      ],
    },
  ],
};

/** Cell note for the `ReceiptRules` key on the Variables sheet (plain text, <= 2,500 chars) */
export const RECEIPT_RULES_NOTE = `ReceiptRules (optional): how receipt emails are found and parsed, as JSON.
Empty or missing = built-in defaults (Chase, PayPal, Venmo). Easiest: edit it in the web app (Settings > Customize variables), which validates and lets you test the rules on real emails before saving. Clear the cell to go back to the built-in defaults.

Shape:
version: 1
defaultTypePrefix: e.g. "TJ" (used in "TJ Venmo" when the email was not for ForwardName)
forwarding: subjectPrefixes (e.g. ["Fwd: "]), fromPatterns (regexes with a named group (?<email>...) that find the original sender in a forwarded body), bodyMarkers ({marker, skipLines}: where the forwarded content starts)
senders[]: address, type ("Paid with" base), typePrefix ("fixed" or "recipient"), forwardable (true/false), rules[]
rules[]: id, subject {pattern, flags}, optional body {source: "plainBody" or "htmlBody", pattern, flags}, name, nameFallback, cost {source: "subject" or "body", group, negate}, category

Patterns are JavaScript regular expressions without slashes; flags may use i, m, s, u. Templates use \${subject.<group>} and \${body.<group>} for named groups; if any referenced group is missing, nameFallback is tried. Rules are tried in order and the first rule whose subject matches wins. Forwarded emails (from ForwardEmailAddress) must start with a subject prefix (letters or digits, then an optional ":" and spaces); other messages from that address, and unknown senders are not receipts (only a note). A forward whose original sender cannot be found is an error.

The Gmail search is built from senders[].address plus ForwardEmailAddress (only with a subject prefix word such as Fwd), always with in:inbox and the processed-label exclusions. Regexes run in Apps Script, so slow patterns slow every import. An invalid value stops receipt imports with an error until it is fixed or cleared. A saved value does not pick up new built-in rules after an upgrade: use Customize variables > "Reset to built-in defaults" > Save, or clear the cell.

Keep column A empty below the last variable: the web app and the setup helper only add ReceiptRules in an empty row (both columns empty) of Variables!A1:B30 below the last non-empty cell in column A.`;
