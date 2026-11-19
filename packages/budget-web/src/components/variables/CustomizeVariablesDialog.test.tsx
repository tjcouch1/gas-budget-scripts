import { describe, expect, it, vi } from "vitest";
import { fireEvent, render, screen, within } from "@testing-library/react";
import userEvent from "@testing-library/user-event";
import {
  DEFAULT_RECEIPT_RULES,
  type FetchReceiptEmailsResultDto,
  type ReceiptEmailDto,
  type ReceiptRulesConfig,
  type ReceiptRulesStateDto,
} from "@budget/shared";
import { ApiError, type ApiRequestBody } from "../../api/client.ts";
import type { Services } from "../../services.ts";
import { CustomizeVariablesDialog } from "./CustomizeVariablesDialog.tsx";

const FWD = "fwd@example.com";

function state(
  overrides: Partial<ReceiptRulesStateDto> = {},
): ReceiptRulesStateDto {
  return {
    raw: "",
    hasVariable: true,
    config: DEFAULT_RECEIPT_RULES,
    isDefault: true,
    issues: [],
    forwardEmailAddress: FWD,
    forwardName: "Sam",
    query: "q",
    ...overrides,
  };
}

const cafe: ReceiptEmailDto = {
  threadId: "t1",
  messageId: "m1",
  from: "Chase <no.reply.alerts@chase.com>",
  to: "me@example.com",
  subject: "You made a $4.50 transaction with Demo Cafe",
  date: "2026-01-02T03:04:05.000Z",
  plainBody: "body",
  plainBodyTruncated: false,
  labels: [],
};

type Handler = (request: ApiRequestBody) => unknown;

function setup(
  initial: ReceiptRulesStateDto = state(),
  handlers: Partial<Record<ApiRequestBody["action"], Handler>> = {},
) {
  const api = vi.fn((request: ApiRequestBody) => {
    const handler = handlers[request.action];
    if (handler) {
      try {
        return Promise.resolve({ data: handler(request), warnings: [] });
      } catch (error) {
        return Promise.reject(error as Error);
      }
    }
    if (request.action === "getReceiptRules")
      return Promise.resolve({ data: initial, warnings: [] });
    if (request.action === "fetchReceiptEmails")
      return Promise.resolve({
        data: {
          query: "q",
          threadCount: 1,
          emails: [cafe],
          truncated: false,
        } satisfies FetchReceiptEmailsResultDto,
        warnings: [],
      });
    return Promise.reject(new Error(`unexpected ${request.action}`));
  });
  const onClose = vi.fn();
  render(
    <CustomizeVariablesDialog
      services={{ api } as unknown as Services}
      onClose={onClose}
    />,
  );
  return { api, onClose, user: userEvent.setup() };
}

const saveButton = () => screen.getByRole("button", { name: "Save to sheet" });
const nameField = () =>
  within(
    screen.getByRole("group", { name: "Rule chase-purchase" }),
  ).getByLabelText("Name");
const lastRequest = (api: ReturnType<typeof setup>["api"]) =>
  api.mock.calls.at(-1)![0];

describe("CustomizeVariablesDialog", () => {
  it("loads the rules via getReceiptRules", async () => {
    const { api } = setup();
    expect(
      screen.getByRole("dialog", { name: "Customize variables" }),
    ).toBeInTheDocument();
    expect(await screen.findByLabelText("Default type prefix")).toHaveValue(
      "TJ",
    );
    expect(api).toHaveBeenCalledWith({ action: "getReceiptRules" });
    expect(saveButton()).toBeDisabled();
  });

  it("shows a load error with Retry", async () => {
    let calls = 0;
    const { user } = setup(state(), {
      getReceiptRules: () => {
        calls++;
        if (calls === 1) throw new ApiError("network", "offline");
        return state();
      },
    });
    expect(await screen.findByRole("alert")).toHaveTextContent("offline");
    await user.click(screen.getByRole("button", { name: "Retry" }));
    expect(
      await screen.findByLabelText("Default type prefix"),
    ).toBeInTheDocument();
  });

  it("shows the server's issues and prefills the JSON for an invalid cell", async () => {
    setup(
      state({
        raw: '{"version": 2}',
        config: null,
        isDefault: false,
        issues: [{ path: "version", message: "must be 1" }],
        query: null,
      }),
    );
    expect(await screen.findByLabelText("Rules JSON")).toHaveValue(
      '{"version": 2}',
    );
    expect(
      screen.getByText(/The ReceiptRules value on the sheet is invalid/),
    ).toBeInTheDocument();
    expect(screen.getByText(/must be 1/)).toBeInTheDocument();
  });

  it("updates the live preview when a template changes", async () => {
    const { user } = setup();
    await screen.findByLabelText("Default type prefix");
    await user.click(screen.getByRole("tab", { name: "Emails" }));
    await user.click(screen.getByRole("button", { name: "Fetch emails" }));
    await user.click(screen.getByRole("tab", { name: /Results/ }));
    const card = await screen.findByRole("article", { name: cafe.subject });
    expect(within(card).getByTestId("outcome-badge")).toHaveTextContent(
      "Receipt",
    );
    expect(within(card).getByTestId("receipt-name")).toHaveTextContent(
      /^Demo Cafe$/,
    );

    await user.click(screen.getByRole("tab", { name: "Rules" }));
    fireEvent.change(nameField(), { target: { value: "${subject.name}!" } });
    expect(
      within(screen.getByRole("article", { name: cafe.subject })).getByTestId(
        "receipt-name",
      ),
    ).toHaveTextContent("Demo Cafe!");
  });

  it("disables Save while issues exist", async () => {
    setup();
    await screen.findByLabelText("Default type prefix");
    fireEvent.change(screen.getByLabelText("Default type prefix"), {
      target: { value: "" },
    });
    expect(saveButton()).toBeDisabled();
    fireEvent.change(screen.getByLabelText("Default type prefix"), {
      target: { value: "AB" },
    });
    expect(saveButton()).toBeEnabled();
  });

  it("saves with expectedRaw and shows Saved", async () => {
    const raw = JSON.stringify(DEFAULT_RECEIPT_RULES, null, 2);
    let saved: ReceiptRulesConfig | null = null;
    const { api, user } = setup(state({ raw, isDefault: false }), {
      saveReceiptRules: (request) => {
        saved = (request as { rules: ReceiptRulesConfig }).rules;
        return state({ raw: JSON.stringify(saved, null, 2), config: saved });
      },
    });
    await screen.findByLabelText("Default type prefix");
    fireEvent.change(nameField(), { target: { value: "${subject.name}!" } });
    await user.click(saveButton());
    const request = lastRequest(api);
    expect(request).toMatchObject({
      action: "saveReceiptRules",
      expectedRaw: raw,
    });
    expect(saved!.senders[0].rules[0].name).toBe("${subject.name}!");
    expect(
      await screen.findByText("Saved to the Variables sheet"),
    ).toBeInTheDocument();
    expect(saveButton()).toBeDisabled();
  });

  it("shows a conflict with Reload", async () => {
    let loads = 0;
    const { user, api } = setup(state(), {
      getReceiptRules: () => {
        loads++;
        return state({ raw: loads > 1 ? "{}" : "" });
      },
      saveReceiptRules: () => {
        throw new ApiError(
          "conflict",
          "The rules on the sheet changed since you loaded them. Reload and try again.",
        );
      },
    });
    await screen.findByLabelText("Default type prefix");
    fireEvent.change(screen.getByLabelText("Default type prefix"), {
      target: { value: "AB" },
    });
    await user.click(saveButton());
    expect(
      await screen.findByText(/changed since you loaded them/),
    ).toBeInTheDocument();
    await user.click(screen.getByRole("button", { name: "Reload" }));
    expect(await screen.findByLabelText("Default type prefix")).toHaveValue(
      "TJ",
    );
    expect(
      api.mock.calls.filter(([r]) => r.action === "getReceiptRules"),
    ).toHaveLength(2);
  });

  it("lists server validation issues from a rejected save", async () => {
    const { user } = setup(state(), {
      saveReceiptRules: () => {
        throw new ApiError(
          "bad_request",
          "ReceiptRules is invalid",
          undefined,
          [{ path: "senders[0].address", message: "server says no" }],
        );
      },
    });
    await screen.findByLabelText("Default type prefix");
    fireEvent.change(screen.getByLabelText("Default type prefix"), {
      target: { value: "AB" },
    });
    await user.click(saveButton());
    expect(await screen.findByText(/server says no/)).toBeInTheDocument();
    fireEvent.change(screen.getByLabelText("Default type prefix"), {
      target: { value: "AC" },
    });
    expect(screen.queryByText(/server says no/)).toBeNull();
  });

  it("moves between section tabs with the arrow keys", async () => {
    const { user } = setup(state());
    await screen.findByLabelText("Default type prefix");
    const rules = screen.getByRole("tab", { name: "Rules" });
    expect(rules).toHaveAttribute("tabindex", "0");
    expect(screen.getByRole("tab", { name: "Emails" })).toHaveAttribute(
      "tabindex",
      "-1",
    );
    rules.focus();
    await user.keyboard("{ArrowRight}");
    const emails = screen.getByRole("tab", { name: "Emails" });
    expect(emails).toHaveAttribute("aria-selected", "true");
    expect(emails).toHaveFocus();
  });

  it("Reset to defaults then Save sends the defaults; Clear variable sends null", async () => {
    const custom = { ...DEFAULT_RECEIPT_RULES, defaultTypePrefix: "ZZ" };
    const raw = JSON.stringify(custom, null, 2);
    const { api, user } = setup(
      state({ raw, config: custom, isDefault: false }),
      {
        saveReceiptRules: (request) => {
          const rules = (request as { rules: ReceiptRulesConfig | null }).rules;
          return rules
            ? state({
                raw: JSON.stringify(rules, null, 2),
                config: rules,
                isDefault: false,
              })
            : state();
        },
      },
    );
    expect(await screen.findByLabelText("Default type prefix")).toHaveValue(
      "ZZ",
    );
    await user.click(
      screen.getByRole("button", { name: "Reset to built-in defaults" }),
    );
    await user.click(saveButton());
    expect(lastRequest(api)).toEqual({
      action: "saveReceiptRules",
      rules: DEFAULT_RECEIPT_RULES,
      expectedRaw: raw,
    });
    await screen.findByText("Saved to the Variables sheet");

    await user.click(
      screen.getByRole("button", {
        name: "Clear variable (use built-in defaults)",
      }),
    );
    await user.click(
      within(screen.getByRole("alertdialog")).getByRole("button", {
        name: "Clear variable",
      }),
    );
    expect(lastRequest(api)).toMatchObject({
      action: "saveReceiptRules",
      rules: null,
    });
  });

  it("prompts before closing with unsaved changes", async () => {
    const { onClose, user } = setup();
    await screen.findByLabelText("Default type prefix");
    await user.click(screen.getByRole("button", { name: "Close" }));
    expect(onClose).toHaveBeenCalledTimes(1);
    onClose.mockClear();

    fireEvent.change(screen.getByLabelText("Default type prefix"), {
      target: { value: "AB" },
    });
    await user.click(screen.getByRole("button", { name: "Close" }));
    expect(onClose).not.toHaveBeenCalled();
    const prompt = screen.getByRole("alertdialog", { name: "Unsaved changes" });
    await user.click(within(prompt).getByRole("button", { name: "Discard" }));
    expect(onClose).toHaveBeenCalledTimes(1);
  });

  it("has Rules / Emails / Results tabs that switch the visible panel", async () => {
    const { user } = setup();
    await screen.findByLabelText("Default type prefix");
    const tablist = screen.getByRole("tablist", {
      name: "Customize variables sections",
    });
    const tabs = within(tablist).getAllByRole("tab");
    expect(tabs.map((t) => t.textContent)).toEqual([
      "Rules",
      "Emails",
      "Results (0)",
    ]);
    expect(tabs[0]).toHaveAttribute("aria-selected", "true");
    const panel = (name: string) => screen.getByTestId(`panel-${name}`);
    expect(panel("rules")).not.toHaveClass("hidden");
    expect(panel("emails")).toHaveClass("hidden");

    await user.click(tabs[1]);
    expect(tabs[1]).toHaveAttribute("aria-selected", "true");
    expect(panel("rules")).toHaveClass("hidden");
    expect(panel("emails")).not.toHaveClass("hidden");
  });
});
