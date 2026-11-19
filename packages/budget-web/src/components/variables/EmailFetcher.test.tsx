import { describe, expect, it, vi } from "vitest";
import { fireEvent, render, screen } from "@testing-library/react";
import userEvent from "@testing-library/user-event";
import {
  buildReceiptSearchQuery,
  type FetchReceiptEmailsResultDto,
  type ReceiptEmailDto,
} from "@budget/shared";
import { ApiError, type CallApi } from "../../api/client.ts";
import { EmailFetcher } from "./EmailFetcher.tsx";

const SENDERS = ["no.reply.alerts@chase.com", "service@paypal.com"];
const FWD = "fwd@example.com";

const email: ReceiptEmailDto = {
  threadId: "t1",
  messageId: "m1",
  from: "no.reply.alerts@chase.com",
  to: "me@example.com",
  subject: "You made a $1.00 transaction with Demo Cafe",
  date: "2026-01-02T03:04:05.000Z",
  plainBody: "hi",
  plainBodyTruncated: false,
  labels: [],
};

function setup(
  api?: CallApi,
  senders = SENDERS,
  forwardSubjectPrefixes?: string[],
) {
  const result: FetchReceiptEmailsResultDto = {
    query: "server query",
    threadCount: 1,
    emails: [email, { ...email, messageId: "m2" }],
    truncated: false,
  };
  const callApi =
    api ??
    (vi.fn(() => Promise.resolve({ data: result, warnings: [] })) as CallApi);
  const onFetched = vi.fn();
  render(
    <EmailFetcher
      senders={senders}
      forwardEmailAddress={FWD}
      forwardSubjectPrefixes={forwardSubjectPrefixes}
      onFetched={onFetched}
      api={callApi}
    />,
  );
  return { api: callApi, onFetched, user: userEvent.setup() };
}

describe("EmailFetcher", () => {
  it("previews the query built by buildReceiptSearchQuery", () => {
    setup();
    expect(screen.getByTestId("query-preview")).toHaveTextContent(
      buildReceiptSearchQuery(SENDERS, FWD).trim(),
    );
  });

  it("uses the candidate forward subject prefixes in the preview and the request", async () => {
    const { api, user } = setup(undefined, SENDERS, ["Fwd: ", "FW: "]);
    expect(screen.getByTestId("query-preview")).toHaveTextContent(
      "OR (from:fwd@example.com subject:(Fwd OR FW)))",
    );
    await user.click(screen.getByRole("button", { name: "Fetch emails" }));
    expect(api).toHaveBeenCalledWith(
      expect.objectContaining({ forwardSubjectPrefixes: ["Fwd: ", "FW: "] }),
    );
  });

  it("updates the preview for processed mail, subjects, and extras", async () => {
    const { user } = setup();
    await user.click(screen.getByLabelText(/Include already-processed emails/));
    fireEvent.change(screen.getByLabelText("Subject contains"), {
      target: { value: "transaction, credit" },
    });
    fireEvent.change(screen.getByLabelText("Extra sender addresses"), {
      target: { value: "x@example.com" },
    });
    fireEvent.change(screen.getByLabelText("Newer than (days)"), {
      target: { value: "30" },
    });
    expect(screen.getByTestId("query-preview")).toHaveTextContent(
      buildReceiptSearchQuery([...SENDERS, "x@example.com"], FWD, {
        includeProcessed: true,
        subjects: ["transaction", "credit"],
        newerThanDays: 30,
      }).trim(),
    );
  });

  it("states that fetching is read-only and processed mail is off by default", () => {
    setup();
    expect(
      screen.getByText(/Fetching never labels or archives anything/),
    ).toBeInTheDocument();
    expect(
      screen.getByLabelText(/Include already-processed emails/),
    ).not.toBeChecked();
    expect(screen.getByLabelText("Max emails")).toHaveValue(10);
  });

  it("shows an inline error for an invalid extra sender and disables Fetch", () => {
    setup();
    const input = screen.getByLabelText("Extra sender addresses");
    fireEvent.change(input, { target: { value: "ok@example.com, nope" } });
    const describedBy = input.getAttribute("aria-describedby");
    expect(document.getElementById(describedBy!)).toHaveTextContent(
      /nope is not a valid email address/,
    );
    expect(screen.getByRole("button", { name: "Fetch emails" })).toBeDisabled();
  });

  it("rejects unsafe subject filters and an out-of-range max", () => {
    setup();
    fireEvent.change(screen.getByLabelText("Subject contains"), {
      target: { value: "a(b" },
    });
    expect(screen.getByText(/a\(b contains characters/)).toBeInTheDocument();
    expect(screen.getByRole("button", { name: "Fetch emails" })).toBeDisabled();
    fireEvent.change(screen.getByLabelText("Subject contains"), {
      target: { value: "" },
    });
    fireEvent.change(screen.getByLabelText("Max emails"), {
      target: { value: "51" },
    });
    expect(screen.getByText(/whole number from 1 to 50/)).toBeInTheDocument();
    expect(screen.getByRole("button", { name: "Fetch emails" })).toBeDisabled();
  });

  it("shows a friendly message for a sender with no address yet", () => {
    setup(undefined, ["", "service@paypal.com"]);
    expect(screen.getByRole("alert")).toHaveTextContent(
      "A sender in the rules has no address yet.",
    );
  });

  it("keeps Fetch enabled for 25 senders plus one extra", () => {
    const many = Array.from({ length: 25 }, (_, i) => `s${i}@example.com`);
    setup(undefined, many);
    fireEvent.change(screen.getByLabelText("Extra sender addresses"), {
      target: { value: "extra@example.com" },
    });
    expect(screen.getByRole("button", { name: "Fetch emails" })).toBeEnabled();
  });

  it("disables Fetch when a candidate sender is invalid", () => {
    setup(undefined, ["not an address", "service@paypal.com"]);
    expect(screen.getByRole("alert")).toHaveTextContent(/Invalid sender/);
    expect(screen.getByRole("button", { name: "Fetch emails" })).toBeDisabled();
  });

  it("fetches with candidate + extra senders and reports the result", async () => {
    const { api, onFetched, user } = setup();
    fireEvent.change(screen.getByLabelText("Extra sender addresses"), {
      target: { value: "x@example.com" },
    });
    await user.click(screen.getByLabelText("Include HTML bodies"));
    fireEvent.change(screen.getByLabelText("Max emails"), {
      target: { value: "5" },
    });
    await user.click(screen.getByRole("button", { name: "Fetch emails" }));
    expect(api).toHaveBeenCalledWith({
      action: "fetchReceiptEmails",
      senders: [...SENDERS, "x@example.com"],
      includeHtml: true,
      includeProcessed: false,
      maxThreads: 5,
    });
    expect(await screen.findByRole("status")).toHaveTextContent(
      "Fetched 2 emails from 1 thread",
    );
    expect(onFetched).toHaveBeenCalledWith(
      [email, { ...email, messageId: "m2" }],
      "server query",
    );
    expect(screen.getByTestId("server-query")).toHaveTextContent(
      "server query",
    );
  });

  it("shows API errors", async () => {
    const api = vi.fn(() =>
      Promise.reject(new ApiError("invalid_config", "ReceiptRules is bad")),
    ) as unknown as CallApi;
    const { user } = setup(api);
    await user.click(screen.getByRole("button", { name: "Fetch emails" }));
    expect(await screen.findByRole("alert")).toHaveTextContent(
      "ReceiptRules is bad",
    );
  });

  it("passes the transaction sheets from the server to onFetched", async () => {
    const sheets = [
      {
        name: "1/1/2026 - 1/14/2026",
        startDate: "2026-01-01T06:00:00.000Z",
        endDate: "2026-01-15T05:59:59.999Z",
      },
    ];
    const api = vi.fn(() =>
      Promise.resolve({
        data: {
          query: "q",
          threadCount: 0,
          emails: [],
          truncated: false,
          transactionSheets: sheets,
        },
      }),
    ) as unknown as CallApi;
    const { onFetched, user } = setup(api);
    await user.click(screen.getByRole("button", { name: "Fetch emails" }));
    await screen.findByRole("status");
    expect(onFetched).toHaveBeenCalledWith([], "q", sheets);
  });
});
