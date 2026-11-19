import { useId, useState } from "react";
import type { ReceiptRule, RuleIssue, SenderRules } from "@budget/shared";
import { Button, ConfirmDialog } from "../ui.tsx";
import { CheckboxField, issuesFor, TextField } from "./fields.tsx";
import { HelpTip } from "./HelpTip.tsx";
import { RuleCard } from "./RuleCard.tsx";

function newRule(existing: ReceiptRule[]): ReceiptRule {
  const ids = new Set(existing.map((rule) => rule.id));
  let n = existing.length + 1;
  while (ids.has(`rule-${n}`)) n++;
  return { id: `rule-${n}`, subject: { pattern: "^" } };
}

/** Editor for one sender. `path` is like "senders[0]". */
export function SenderCard({
  sender,
  index,
  path,
  issues,
  onChange,
  onRemove,
}: {
  sender: SenderRules;
  index: number;
  path: string;
  issues: RuleIssue[];
  onChange: (sender: SenderRules) => void;
  onRemove: () => void;
}) {
  const [confirming, setConfirming] = useState(false);
  const radioName = useId();
  const at = (sub: string) => issuesFor(issues, `${path}.${sub}`);
  const setRules = (rules: ReceiptRule[]) => onChange({ ...sender, rules });
  const label = `Sender ${sender.address || index + 1}`;
  return (
    <section
      aria-label={label}
      className="space-y-3 rounded-xl bg-white p-3 ring-1 ring-slate-200 dark:bg-slate-800/60 dark:ring-slate-700"
    >
      <div className="grid gap-3 sm:grid-cols-2">
        <TextField
          label="Sender address"
          help="senderAddress"
          type="email"
          value={sender.address}
          issues={at("address")}
          mono
          placeholder="alerts@example.com"
          onChange={(address) => onChange({ ...sender, address })}
        />
        <TextField
          label="Paid with type"
          help="senderType"
          value={sender.type}
          issues={at("type")}
          placeholder="Credit"
          onChange={(type) => onChange({ ...sender, type })}
        />
      </div>
      <fieldset>
        <legend className="flex items-center text-sm font-medium">
          Type prefix
          <HelpTip help="typePrefix" label="Type prefix" />
        </legend>
        <div className="flex flex-wrap gap-x-4">
          {(
            [
              ["fixed", "Fixed"],
              ["recipient", "Prefix by recipient"],
            ] as const
          ).map(([value, text]) => (
            <label
              key={value}
              className="flex min-h-11 items-center gap-2 text-sm"
            >
              <input
                type="radio"
                name={radioName}
                className="h-5 w-5"
                checked={sender.typePrefix === value}
                onChange={() => onChange({ ...sender, typePrefix: value })}
              />
              {text}
            </label>
          ))}
        </div>
      </fieldset>
      <CheckboxField
        label="Forwardable"
        help="forwardable"
        checked={sender.forwardable}
        onChange={(forwardable) => onChange({ ...sender, forwardable })}
      />
      {at("rules").length > 0 && (
        <p className="text-sm text-red-600 dark:text-red-400">
          {at("rules")
            .map((issue) => issue.message)
            .join("; ")}
        </p>
      )}
      <div className="space-y-3">
        {sender.rules.map((rule, i) => (
          <RuleCard
            // Index keys: ids are editable and may be duplicated while typing.
            key={i}
            rule={rule}
            index={i}
            path={`${path}.rules[${i}]`}
            issues={issues}
            onChange={(next) =>
              setRules(sender.rules.map((r, j) => (j === i ? next : r)))
            }
            onRemove={() => setRules(sender.rules.filter((_, j) => j !== i))}
          />
        ))}
      </div>
      <div className="flex flex-wrap gap-2">
        <Button
          onClick={() => setRules([...sender.rules, newRule(sender.rules)])}
        >
          Add rule
        </Button>
        <Button variant="ghost" onClick={() => setConfirming(true)}>
          Remove sender
        </Button>
      </div>
      {confirming && (
        <ConfirmDialog
          title={`Remove ${sender.address || "this sender"}?`}
          confirmLabel="Remove sender"
          onConfirm={() => {
            setConfirming(false);
            onRemove();
          }}
          onCancel={() => setConfirming(false)}
        >
          Its {sender.rules.length} rule(s) are removed from the draft. Nothing
          changes on the sheet until you save.
        </ConfirmDialog>
      )}
    </section>
  );
}
