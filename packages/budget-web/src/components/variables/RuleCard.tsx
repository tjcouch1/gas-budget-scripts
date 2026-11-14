import type { BodySource, ReceiptRule, RuleIssue } from "@budget/shared";
import { Button } from "../ui.tsx";
import {
  CheckboxField,
  issuesFor,
  SelectField,
  TextField,
  withOptional,
} from "./fields.tsx";

const BODY_SOURCES: { value: BodySource; label: string }[] = [
  { value: "plainBody", label: "Plain body" },
  { value: "htmlBody", label: "HTML body" },
];

const COST_SOURCES: { value: "subject" | "body"; label: string }[] = [
  { value: "subject", label: "Subject" },
  { value: "body", label: "Body" },
];

/** Editor for one rule of a sender. `path` is like "senders[0].rules[1]". */
export function RuleCard({
  rule,
  index,
  path,
  issues,
  onChange,
  onRemove,
}: {
  rule: ReceiptRule;
  index: number;
  path: string;
  issues: RuleIssue[];
  onChange: (rule: ReceiptRule) => void;
  onRemove: () => void;
}) {
  const at = (sub: string) => issuesFor(issues, `${path}.${sub}`);
  const { body, cost } = rule;
  return (
    <fieldset
      aria-label={`Rule ${rule.id || index + 1}`}
      className="space-y-3 rounded-lg border border-slate-200 p-3 dark:border-slate-700"
    >
      <div className="flex items-end gap-2">
        <TextField
          className="flex-1"
          label="Rule id"
          value={rule.id}
          issues={at("id")}
          mono
          onChange={(id) => onChange({ ...rule, id })}
        />
        <Button variant="ghost" onClick={onRemove}>
          Remove rule
        </Button>
      </div>
      {issuesFor(issues, path).length > 0 && (
        <p className="text-sm text-red-600 dark:text-red-400">
          {issuesFor(issues, path)
            .map((issue) => issue.message)
            .join("; ")}
        </p>
      )}
      <div className="grid gap-3 sm:grid-cols-[1fr_6rem]">
        <TextField
          label="Subject pattern"
          value={rule.subject.pattern}
          issues={[...at("subject"), ...at("subject.pattern")]}
          mono
          onChange={(pattern) =>
            onChange({ ...rule, subject: { ...rule.subject, pattern } })
          }
        />
        <TextField
          label="Subject flags"
          value={rule.subject.flags ?? ""}
          issues={at("subject.flags")}
          mono
          placeholder="i"
          onChange={(flags) =>
            onChange({
              ...rule,
              subject: withOptional(rule.subject, "flags", flags),
            })
          }
        />
      </div>

      <CheckboxField
        label="Also match the body"
        checked={body !== undefined}
        onChange={(checked) =>
          onChange(
            withOptional(
              rule,
              "body",
              checked ? { source: "plainBody", pattern: "" } : undefined,
            ),
          )
        }
      />
      {body && (
        <div className="grid gap-3 sm:grid-cols-[10rem_1fr_6rem]">
          <SelectField
            label="Body source"
            value={body.source}
            options={BODY_SOURCES}
            issues={at("body.source")}
            onChange={(source) =>
              onChange({ ...rule, body: { ...body, source } })
            }
          />
          <TextField
            label="Body pattern"
            value={body.pattern}
            issues={[...at("body"), ...at("body.pattern")]}
            mono
            onChange={(pattern) =>
              onChange({ ...rule, body: { ...body, pattern } })
            }
          />
          <TextField
            label="Body flags"
            value={body.flags ?? ""}
            issues={at("body.flags")}
            mono
            onChange={(flags) =>
              onChange({ ...rule, body: withOptional(body, "flags", flags) })
            }
          />
        </div>
      )}

      <div className="grid gap-3 sm:grid-cols-2">
        <TextField
          label="Name"
          value={rule.name ?? ""}
          issues={at("name")}
          mono
          placeholder="${subject.name}"
          onChange={(name) => onChange(withOptional(rule, "name", name))}
        />
        <TextField
          label="Name fallback"
          value={rule.nameFallback ?? ""}
          issues={at("nameFallback")}
          mono
          onChange={(nameFallback) =>
            onChange(withOptional(rule, "nameFallback", nameFallback))
          }
        />
      </div>

      <CheckboxField
        label="Has a cost"
        checked={cost !== undefined}
        onChange={(checked) =>
          onChange(
            withOptional(
              rule,
              "cost",
              checked ? { source: "subject", group: "cost" } : undefined,
            ),
          )
        }
      />
      {cost && (
        <div className="grid gap-3 sm:grid-cols-[10rem_1fr_auto] sm:items-end">
          <SelectField
            label="Cost source"
            value={cost.source}
            options={COST_SOURCES}
            issues={at("cost.source")}
            onChange={(source) =>
              onChange({ ...rule, cost: { ...cost, source } })
            }
          />
          <TextField
            label="Cost group"
            value={cost.group}
            issues={[...at("cost"), ...at("cost.group")]}
            mono
            onChange={(group) =>
              onChange({ ...rule, cost: { ...cost, group } })
            }
          />
          <CheckboxField
            label="Negate (refund)"
            checked={cost.negate === true}
            onChange={(negate) =>
              onChange({
                ...rule,
                cost: withOptional(cost, "negate", negate || undefined),
              })
            }
          />
        </div>
      )}

      <TextField
        label="Category"
        value={rule.category ?? ""}
        issues={at("category")}
        onChange={(category) =>
          onChange(withOptional(rule, "category", category))
        }
      />
    </fieldset>
  );
}
