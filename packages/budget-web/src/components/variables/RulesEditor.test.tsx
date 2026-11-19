import { describe, expect, it, vi } from "vitest";
import { fireEvent, render, screen, within } from "@testing-library/react";
import userEvent from "@testing-library/user-event";
import {
  DEFAULT_RECEIPT_RULES,
  type ReceiptRulesConfig,
  type RuleIssue,
} from "@budget/shared";
import { RulesEditor } from "./RulesEditor.tsx";

function setup(
  options: {
    config?: ReceiptRulesConfig;
    issues?: RuleIssue[];
    jsonSeed?: string;
  } = {},
) {
  const onChange = vi.fn<(config: ReceiptRulesConfig) => void>();
  const config = options.config ?? DEFAULT_RECEIPT_RULES;
  const view = render(
    <RulesEditor
      config={config}
      onChange={onChange}
      issues={options.issues ?? []}
      forwardEmailAddress="fwd@example.com"
      forwardName="Sam"
      jsonSeed={options.jsonSeed}
    />,
  );
  return { onChange, view, user: userEvent.setup() };
}

const lastConfig = (onChange: ReturnType<typeof setup>["onChange"]) =>
  onChange.mock.calls.at(-1)![0];

const rule = (id: string) => screen.getByRole("group", { name: `Rule ${id}` });

describe("RulesEditor", () => {
  it("renders the defaults: three sender cards and the rule ids", () => {
    setup();
    expect(screen.getAllByRole("region", { name: /^Sender / })).toHaveLength(3);
    for (const id of ["chase-purchase", "chase-refund"])
      expect(rule(id)).toBeInTheDocument();
    expect(screen.getByText("fwd@example.com")).toBeInTheDocument();
    expect(screen.getByText("Sam")).toBeInTheDocument();
  });

  it("edits the Chase purchase name template", () => {
    const { onChange } = setup();
    fireEvent.change(within(rule("chase-purchase")).getByLabelText("Name"), {
      target: { value: "${subject.name}!" },
    });
    const config = lastConfig(onChange);
    expect(config.senders[0].rules[0].name).toBe("${subject.name}!");
    expect(DEFAULT_RECEIPT_RULES.senders[0].rules[0].name).toBe(
      "${subject.name}",
    );
  });

  it("edits the default type prefix", () => {
    const { onChange } = setup();
    fireEvent.change(screen.getByLabelText("Default type prefix"), {
      target: { value: "AB" },
    });
    expect(lastConfig(onChange).defaultTypePrefix).toBe("AB");
  });

  it("toggles 'Also match the body' to add and remove body", async () => {
    const { onChange, user } = setup();
    const purchase = within(rule("chase-purchase"));
    await user.click(purchase.getByLabelText("Also match the body"));
    expect(lastConfig(onChange).senders[0].rules[0].body).toEqual({
      source: "plainBody",
      pattern: "",
    });
    await user.click(
      within(rule("chase-refund")).getByLabelText("Also match the body"),
    );
    expect(lastConfig(onChange).senders[0].rules[1].body).toBeUndefined();
  });

  it("toggles 'Has a cost'", async () => {
    const { onChange, user } = setup();
    await user.click(
      within(rule("chase-purchase")).getByLabelText("Has a cost"),
    );
    expect(lastConfig(onChange).senders[0].rules[0].cost).toBeUndefined();
  });

  it("shows an issue under its field, linked by aria-describedby", () => {
    setup({
      issues: [
        {
          path: "senders[0].rules[0].subject.pattern",
          message: "Invalid regular expression",
        },
      ],
    });
    const input = within(rule("chase-purchase")).getByLabelText(
      "Subject pattern",
    );
    const id = input.getAttribute("aria-describedby");
    expect(id).toBeTruthy();
    expect(document.getElementById(id!)).toHaveTextContent(
      "Invalid regular expression",
    );
    expect(input).toHaveAttribute("aria-invalid", "true");
  });

  it("shows pretty JSON in the JSON tab", async () => {
    const { user } = setup();
    await user.click(screen.getByRole("tab", { name: "JSON" }));
    expect(screen.getByLabelText("Rules JSON")).toHaveValue(
      JSON.stringify(DEFAULT_RECEIPT_RULES, null, 2),
    );
  });

  it("Apply JSON with invalid JSON shows the SyntaxError and does not call onChange", async () => {
    const { onChange, user } = setup();
    await user.click(screen.getByRole("tab", { name: "JSON" }));
    fireEvent.change(screen.getByLabelText("Rules JSON"), {
      target: { value: "{ nope" },
    });
    await user.click(screen.getByRole("button", { name: "Apply JSON" }));
    let message = "";
    try {
      JSON.parse("{ nope");
    } catch (error) {
      message = (error as Error).message;
    }
    expect(screen.getByRole("alert")).toHaveTextContent(message);
    expect(onChange).not.toHaveBeenCalled();
  });

  it("Apply JSON with valid JSON calls onChange", async () => {
    const { onChange, user } = setup();
    await user.click(screen.getByRole("tab", { name: "JSON" }));
    const next = { ...DEFAULT_RECEIPT_RULES, defaultTypePrefix: "XY" };
    fireEvent.change(screen.getByLabelText("Rules JSON"), {
      target: { value: JSON.stringify(next) },
    });
    await user.click(screen.getByRole("button", { name: "Apply JSON" }));
    expect(lastConfig(onChange)).toEqual(next);
  });

  it("Apply JSON refuses a shape the form cannot show", async () => {
    const { onChange, user } = setup();
    await user.click(screen.getByRole("tab", { name: "JSON" }));
    fireEvent.change(screen.getByLabelText("Rules JSON"), {
      target: {
        value:
          '{"forwarding":{"fromPatterns":[],"bodyMarkers":[]},"senders":3}',
      },
    });
    await user.click(screen.getByRole("button", { name: "Apply JSON" }));
    expect(screen.getByRole("alert")).toHaveTextContent(/senders/);
    expect(onChange).not.toHaveBeenCalled();
  });

  it("starts on the JSON tab with seeded text", () => {
    setup({ jsonSeed: '{"broken": true}' });
    expect(screen.getByLabelText("Rules JSON")).toHaveValue('{"broken": true}');
  });

  it("lists all issues in a summary", () => {
    setup({ issues: [{ path: "zzz", message: "unknown key" }] });
    expect(
      screen.getByRole("group", { name: "Rule problems" }),
    ).toHaveTextContent("zzz: unknown key");
  });

  it("adds a sender and removes one after confirming", async () => {
    const { onChange, user } = setup();
    await user.click(screen.getByRole("button", { name: "Add sender" }));
    expect(lastConfig(onChange).senders).toHaveLength(4);

    const chase = screen.getByRole("region", {
      name: "Sender no.reply.alerts@chase.com",
    });
    await user.click(
      within(chase).getByRole("button", { name: "Remove sender" }),
    );
    const confirm = screen.getByRole("alertdialog");
    await user.click(
      within(confirm).getByRole("button", { name: "Remove sender" }),
    );
    const senders = lastConfig(onChange).senders;
    expect(senders).toHaveLength(2);
    expect(senders.map((s) => s.address)).not.toContain(
      "no.reply.alerts@chase.com",
    );
  });

  it("adds a rule with a unique id and removes a rule", async () => {
    const { onChange, user } = setup();
    const chase = within(
      screen.getByRole("region", { name: "Sender no.reply.alerts@chase.com" }),
    );
    await user.click(chase.getByRole("button", { name: "Add rule" }));
    const rules = lastConfig(onChange).senders[0].rules;
    expect(rules).toHaveLength(
      DEFAULT_RECEIPT_RULES.senders[0].rules.length + 1,
    );
    expect(new Set(rules.map((r) => r.id)).size).toBe(rules.length);

    await user.click(
      within(rule("chase-refund")).getByRole("button", { name: "Remove rule" }),
    );
    expect(
      lastConfig(onChange).senders[0].rules.map((r) => r.id),
    ).not.toContain("chase-refund");
  });

  it("edits sender fields: type prefix radio and forwardable", async () => {
    const { onChange, user } = setup();
    const chase = within(
      screen.getByRole("region", { name: "Sender no.reply.alerts@chase.com" }),
    );
    await user.click(chase.getByLabelText("Prefix by recipient"));
    expect(lastConfig(onChange).senders[0].typePrefix).toBe("recipient");
    await user.click(chase.getByLabelText("Forwardable"));
    expect(lastConfig(onChange).senders[0].forwardable).toBe(true);
  });

  it("edits forwarding settings", () => {
    const { onChange } = setup();
    fireEvent.change(screen.getByLabelText("Body marker 1"), {
      target: { value: "-- fwd --" },
    });
    expect(lastConfig(onChange).forwarding.bodyMarkers[0].marker).toBe(
      "-- fwd --",
    );
    fireEvent.change(screen.getByLabelText("Forward subject prefix 1"), {
      target: { value: "FW: " },
    });
    expect(lastConfig(onChange).forwarding.subjectPrefixes).toEqual(["FW: "]);
  });

  it("labels every form control", () => {
    setup();
    const controls = document.querySelectorAll("input, select, textarea");
    expect(controls.length).toBeGreaterThan(20);
    for (const control of controls) {
      const labelled =
        control.getAttribute("aria-label") ||
        (control.id &&
          document.querySelector(`label[for="${CSS.escape(control.id)}"]`)) ||
        control.closest("label");
      expect(labelled, control.outerHTML).toBeTruthy();
    }
  });

  it("moves between mode tabs with the arrow keys", async () => {
    const { user } = setup();
    const form = screen.getByRole("tab", { name: "Form" });
    form.focus();
    await user.keyboard("{ArrowRight}");
    const json = screen.getByRole("tab", { name: "JSON" });
    expect(json).toHaveAttribute("aria-selected", "true");
    expect(json).toHaveFocus();
    expect(form).toHaveAttribute("tabindex", "-1");
  });
});
