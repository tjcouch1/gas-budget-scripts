import { useId, useMemo, useState, type KeyboardEvent } from "react";
import type {
  ReceiptRulesConfig,
  RuleIssue,
  SenderRules,
} from "@budget/shared";
import { Button } from "../ui.tsx";
import { issuesFor, monoInputClass, TextField } from "./fields.tsx";
import { ForwardingEditor } from "./ForwardingEditor.tsx";
import { SenderCard } from "./SenderCard.tsx";

/** Roving-tabindex keyboard handling for a tablist (ArrowLeft/Right wrap, Home, End). */
export function tabKeyDown<T extends string>(
  e: KeyboardEvent<HTMLElement>,
  ids: T[],
  current: T,
  select: (id: T) => void,
  idOf: (id: T) => string,
) {
  const i = ids.indexOf(current);
  let next: number;
  if (e.key === "ArrowRight") next = (i + 1) % ids.length;
  else if (e.key === "ArrowLeft") next = (i - 1 + ids.length) % ids.length;
  else if (e.key === "Home") next = 0;
  else if (e.key === "End") next = ids.length - 1;
  else return;
  e.preventDefault();
  const id = ids[next]!;
  select(id);
  document.getElementById(idOf(id))?.focus();
}

const isObject = (value: unknown): value is Record<string, unknown> =>
  typeof value === "object" && value !== null && !Array.isArray(value);

/**
 * Returns a reason when the parsed JSON has a shape the form cannot render (so it is not applied),
 * or null. Everything else, including invalid values, is applied and reported by the live validation.
 */
export function formShapeProblem(value: unknown): string | null {
  if (!isObject(value)) return "The JSON must be an object.";
  const { forwarding, senders } = value;
  if (!isObject(forwarding)) return "forwarding must be an object.";
  for (const key of ["fromPatterns", "bodyMarkers"] as const)
    if (!Array.isArray(forwarding[key]) || !forwarding[key].every(isObject))
      return `forwarding.${key} must be an array of objects.`;
  if (
    forwarding.subjectPrefixes !== undefined &&
    !Array.isArray(forwarding.subjectPrefixes)
  )
    return "forwarding.subjectPrefixes must be an array.";
  if (!Array.isArray(senders)) return "senders must be an array.";
  for (const [i, sender] of senders.entries()) {
    if (!isObject(sender)) return `senders[${i}] must be an object.`;
    if (!Array.isArray(sender.rules))
      return `senders[${i}].rules must be an array.`;
    for (const [j, rule] of sender.rules.entries()) {
      const at = `senders[${i}].rules[${j}]`;
      if (!isObject(rule) || !isObject(rule.subject))
        return `${at} must be an object with a subject object.`;
      for (const key of ["body", "cost"] as const)
        if (rule[key] !== undefined && !isObject(rule[key]))
          return `${at}.${key} must be an object.`;
    }
  }
  return null;
}

type EditorTab = "form" | "json";

export function RulesEditor({
  config,
  onChange,
  issues,
  forwardEmailAddress,
  forwardName,
  jsonSeed,
}: {
  config: ReceiptRulesConfig;
  onChange: (config: ReceiptRulesConfig) => void;
  issues: RuleIssue[];
  forwardEmailAddress: string;
  forwardName: string;
  /** Initial JSON text (e.g. the sheet's invalid cell); opens the JSON tab. */
  jsonSeed?: string;
}) {
  const ids = useId();
  const [tab, setTab] = useState<EditorTab>(jsonSeed ? "json" : "form");
  const pretty = useMemo(() => JSON.stringify(config, null, 2), [config]);
  const [jsonBase, setJsonBase] = useState(pretty);
  const [jsonText, setJsonText] = useState(jsonSeed ?? pretty);
  const [jsonError, setJsonError] = useState<string | null>(null);
  // The JSON text follows the config whenever the config changes (form edit, reset, apply).
  if (jsonBase !== pretty) {
    setJsonBase(pretty);
    setJsonText(pretty);
    setJsonError(null);
  }

  const setSenders = (senders: SenderRules[]) =>
    onChange({ ...config, senders });

  const applyJson = () => {
    let parsed: unknown;
    try {
      parsed = JSON.parse(jsonText);
    } catch (error) {
      setJsonError(error instanceof Error ? error.message : String(error));
      return;
    }
    const problem = formShapeProblem(parsed);
    if (problem) {
      setJsonError(problem);
      return;
    }
    setJsonError(null);
    onChange(parsed as ReceiptRulesConfig);
  };

  const tabs: { id: EditorTab; label: string }[] = [
    { id: "form", label: "Form" },
    { id: "json", label: "JSON" },
  ];

  return (
    <div className="space-y-3">
      <section
        aria-label="Forwarding variables"
        className="space-y-3 rounded-xl bg-slate-100 p-3 text-sm dark:bg-slate-800"
      >
        <dl className="grid grid-cols-[auto_1fr] gap-x-3 gap-y-1">
          <dt className="text-slate-500 dark:text-slate-400">
            ForwardEmailAddress
          </dt>
          <dd className="font-mono break-all">
            {forwardEmailAddress || "(not set)"}
          </dd>
          <dt className="text-slate-500 dark:text-slate-400">ForwardName</dt>
          <dd>{forwardName || "(not set)"}</dd>
        </dl>
        <p className="text-xs text-slate-500 dark:text-slate-400">
          Edit these two in the sheet&apos;s Variables tab.
        </p>
        <TextField
          label="Default type prefix"
          value={config.defaultTypePrefix}
          issues={issuesFor(issues, "defaultTypePrefix")}
          hint='Used for "Prefix by recipient" senders when the email was not for the forwarder.'
          onChange={(defaultTypePrefix) =>
            onChange({ ...config, defaultTypePrefix })
          }
        />
      </section>

      {issues.length > 0 && (
        <div
          className="rounded-lg bg-red-50 p-3 text-sm text-red-800 dark:bg-red-950/50 dark:text-red-200"
          aria-label="Rule problems"
          role="group"
        >
          <p className="font-medium">
            {issues.length} problem{issues.length === 1 ? "" : "s"} to fix
            before saving:
          </p>
          <ul className="mt-1 list-disc space-y-0.5 pl-5">
            {issues.map((issue, i) => (
              <li key={i} className="break-words">
                <code>{issue.path || "(root)"}</code>: {issue.message}
              </li>
            ))}
          </ul>
        </div>
      )}

      <div role="tablist" aria-label="Rules editor mode" className="flex gap-1">
        {tabs.map((t) => (
          <button
            key={t.id}
            type="button"
            role="tab"
            id={`${ids}-tab-${t.id}`}
            aria-controls={`${ids}-panel-${t.id}`}
            aria-selected={tab === t.id}
            tabIndex={tab === t.id ? 0 : -1}
            onKeyDown={(e) =>
              tabKeyDown(
                e,
                tabs.map((x) => x.id),
                tab,
                setTab,
                (id) => `${ids}-tab-${id}`,
              )
            }
            onClick={() => setTab(t.id)}
            className={`min-h-11 rounded-lg px-4 text-sm font-medium ${
              tab === t.id
                ? "bg-teal-700 text-white"
                : "text-slate-700 hover:bg-slate-200/60 dark:text-slate-200 dark:hover:bg-slate-700/60"
            }`}
          >
            {t.label}
          </button>
        ))}
      </div>

      {tab === "form" ? (
        <div
          role="tabpanel"
          id={`${ids}-panel-form`}
          aria-labelledby={`${ids}-tab-form`}
          className="space-y-3"
        >
          {config.senders.map((sender, i) => (
            <SenderCard
              key={i}
              sender={sender}
              index={i}
              path={`senders[${i}]`}
              issues={issues}
              onChange={(next) =>
                setSenders(config.senders.map((s, j) => (j === i ? next : s)))
              }
              onRemove={() =>
                setSenders(config.senders.filter((_, j) => j !== i))
              }
            />
          ))}
          <Button
            onClick={() =>
              setSenders([
                ...config.senders,
                {
                  address: "",
                  type: "",
                  typePrefix: "fixed",
                  forwardable: false,
                  rules: [],
                },
              ])
            }
          >
            Add sender
          </Button>
          <ForwardingEditor
            forwarding={config.forwarding}
            issues={issues}
            onChange={(forwarding) => onChange({ ...config, forwarding })}
          />
        </div>
      ) : (
        <div
          role="tabpanel"
          id={`${ids}-panel-json`}
          aria-labelledby={`${ids}-tab-json`}
          className="space-y-2"
        >
          <textarea
            aria-label="Rules JSON"
            value={jsonText}
            spellCheck={false}
            rows={18}
            onChange={(e) => setJsonText(e.target.value)}
            className={`${monoInputClass} min-h-64 whitespace-pre`}
          />
          {jsonError && (
            <p role="alert" className="text-sm text-red-600 dark:text-red-400">
              {jsonError}
            </p>
          )}
          <div className="flex flex-wrap gap-2">
            <Button variant="primary" onClick={applyJson}>
              Apply JSON
            </Button>
            <Button
              onClick={() => {
                setJsonText(pretty);
                setJsonError(null);
              }}
            >
              Revert text
            </Button>
          </div>
        </div>
      )}
    </div>
  );
}
