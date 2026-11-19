import { describe, expect, it } from "vitest";
import { render, screen } from "@testing-library/react";
import userEvent from "@testing-library/user-event";
import { DEFAULT_RECEIPT_RULES } from "@budget/shared";
import { RulesEditor } from "./RulesEditor.tsx";
import { FIELD_HELP } from "./help.ts";

describe("field tooltips", () => {
  it("shows a help button with title text on every receipt-rule field", () => {
    render(
      <RulesEditor
        config={DEFAULT_RECEIPT_RULES}
        onChange={() => undefined}
        issues={[]}
        forwardEmailAddress="fwd@example.com"
        forwardName="Sam"
      />,
    );
    const help = (name: string) =>
      screen.getAllByRole("button", { name: `Help for ${name}` })[0];
    expect(help("Default type prefix")).toHaveAttribute(
      "title",
      FIELD_HELP.defaultTypePrefix,
    );
    expect(help("Sender address")).toHaveAttribute(
      "title",
      FIELD_HELP.senderAddress,
    );
    expect(help("Type prefix")).toHaveAttribute("title", FIELD_HELP.typePrefix);
    expect(help("Forwardable")).toHaveAttribute(
      "title",
      FIELD_HELP.forwardable,
    );
    expect(help("Rule id")).toHaveAttribute("title", FIELD_HELP.ruleId);
    expect(help("Subject pattern")).toHaveAttribute(
      "title",
      FIELD_HELP.subjectPattern,
    );
    expect(help("Category")).toHaveAttribute("title", FIELD_HELP.category);
  });

  it("toggles the same text inline when tapped, linked by aria-describedby", async () => {
    const user = userEvent.setup();
    render(
      <RulesEditor
        config={DEFAULT_RECEIPT_RULES}
        onChange={() => undefined}
        issues={[]}
        forwardEmailAddress="fwd@example.com"
        forwardName="Sam"
      />,
    );
    const button = screen.getByRole("button", {
      name: "Help for Default type prefix",
    });
    expect(screen.queryByRole("note")).toBeNull();
    await user.click(button);
    const note = screen.getByRole("note");
    expect(note).toHaveTextContent(FIELD_HELP.defaultTypePrefix);
    expect(button).toHaveAttribute("aria-describedby", note.id);
    await user.click(button);
    expect(screen.queryByRole("note")).toBeNull();
  });
});
