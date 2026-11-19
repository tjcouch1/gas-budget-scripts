import type { EnvelopeRowDto } from "@budget/shared";
import { envelopeRemain } from "../lib/rows.ts";
import {
  getEnvelopeExpanded,
  recordRecentEnvelope,
  setEnvelopeExpanded,
} from "../lib/preferences.ts";
import { ChipChooser } from "./ChipChooser.tsx";

/**
 * The Envelope chooser shared by the transaction sheet and the new-expense
 * sheet: each pill shows the envelope's remaining amount under its name,
 * choices are recorded as recently used, and (with `follow`) the expanded
 * state is the saved per-device preference.
 */
export function EnvelopeChipChooser({
  value,
  options,
  layoutOptions,
  envelopes,
  onSelect,
  disabled,
  follow = false,
  inline = false,
}: {
  value: string;
  /** Recently-used order, shown while collapsed. */
  options: string[];
  /** The sheet's layout order, shown while expanded. */
  layoutOptions?: string[] | null;
  envelopes: EnvelopeRowDto[] | null;
  onSelect: (value: string) => void;
  disabled?: boolean;
  /** Start in, and write toggles to, the saved expanded preference. */
  follow?: boolean;
  /** Titleless and always collapsed (the expanded preference is ignored), for use inside a list row. */
  inline?: boolean;
}) {
  return (
    <ChipChooser
      label="Envelope"
      options={options}
      expandedOptions={layoutOptions ?? undefined}
      value={value}
      onSelect={(v) => {
        recordRecentEnvelope(v);
        onSelect(v);
      }}
      disabled={disabled}
      hideHeader={inline}
      defaultExpanded={!inline && follow && getEnvelopeExpanded()}
      onExpandedChange={follow && !inline ? setEnvelopeExpanded : undefined}
      renderExtra={(option, selected) => {
        const remain = envelopeRemain(envelopes, option);
        if (!remain?.display) return null;
        return (
          <span
            className={`text-2xs ${
              remain.value !== null && remain.value < 0
                ? selected
                  ? "text-red-200"
                  : "text-red-600"
                : "opacity-70"
            }`}
          >
            {remain.display}
          </span>
        );
      }}
    />
  );
}
