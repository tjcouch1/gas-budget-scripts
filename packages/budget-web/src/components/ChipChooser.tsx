import { useLayoutEffect, useRef, useState, type ReactNode } from "react";

/**
 * A labelled row of pills. Collapsed it is one pill tall and scrolls
 * sideways (the selected pill is scrolled into view); the header toggle
 * expands it so every pill wraps into view at once.
 */
export function ChipChooser({
  label,
  options,
  expandedOptions,
  value,
  onSelect,
  renderExtra,
  disabled = false,
  defaultExpanded = false,
  onExpandedChange,
}: {
  /** Header text and the accessible name of the pill group. */
  label: string;
  options: string[];
  /** Order used while expanded (e.g. the sheet's layout); defaults to `options`. */
  expandedOptions?: string[];
  value: string;
  /** Called with the tapped option, or "" for Clear. */
  onSelect: (value: string) => void;
  /** Extra content inside a pill (e.g. an envelope's balance). */
  renderExtra?: (option: string, selected: boolean) => ReactNode;
  disabled?: boolean;
  /** Initial state only; the chooser manages its own toggling afterwards. */
  defaultExpanded?: boolean;
  /** Called when the user toggles the header. */
  onExpandedChange?: (expanded: boolean) => void;
}) {
  const [expanded, setExpanded] = useState(defaultExpanded);
  const scroller = useRef<HTMLDivElement>(null);
  // The current value stays choosable even when it is not in the list.
  const shown = expanded && expandedOptions ? expandedOptions : options;
  const all = value && !shown.includes(value) ? [...shown, value] : shown;

  useLayoutEffect(() => {
    // scrollIntoView would also scroll the drawer vertically, so set scrollLeft directly.
    const el = scroller.current;
    const selected = el?.querySelector<HTMLElement>('[aria-pressed="true"]');
    if (el && selected) el.scrollLeft = Math.max(0, selected.offsetLeft - 8);
    // Only on mount: later taps should not jump the row.
  }, []);

  return (
    <div role="group" aria-label={label}>
      <button
        type="button"
        aria-expanded={expanded}
        onClick={() => {
          setExpanded(!expanded);
          onExpandedChange?.(!expanded);
        }}
        className="flex min-h-11 items-center gap-1 text-sm font-medium"
      >
        {label}
        <span
          aria-hidden="true"
          className={`text-slate-400 transition-transform ${expanded ? "rotate-90" : ""}`}
        >
          ›
        </span>
      </button>
      <div
        ref={scroller}
        data-testid="chip-scroller"
        className={`relative flex gap-2 py-0.5 ${
          expanded ? "flex-wrap" : "no-scrollbar flex-nowrap overflow-x-auto"
        }`}
      >
        {all.map((option) => {
          const selected = option === value;
          return (
            <button
              key={option}
              type="button"
              aria-pressed={selected}
              disabled={disabled}
              onClick={() => onSelect(option)}
              className={`min-h-9 shrink-0 whitespace-nowrap rounded-full px-3 text-xs ring-1 transition-colors disabled:opacity-50 ${
                selected
                  ? "bg-teal-700 text-white ring-teal-700"
                  : "bg-white ring-slate-300 hover:bg-teal-50 dark:bg-slate-800 dark:ring-slate-600 dark:hover:bg-slate-700"
              }`}
            >
              <span className="inline-flex flex-col items-end">
                <span>{option}</span>
                {renderExtra?.(option, selected)}
              </span>
            </button>
          );
        })}
        {value && (
          <button
            type="button"
            disabled={disabled}
            onClick={() => onSelect("")}
            className="min-h-9 shrink-0 rounded-full px-3 text-xs text-slate-500 ring-1 ring-dashed ring-slate-300 disabled:opacity-50 dark:ring-slate-600"
          >
            Clear
          </button>
        )}
      </div>
    </div>
  );
}
