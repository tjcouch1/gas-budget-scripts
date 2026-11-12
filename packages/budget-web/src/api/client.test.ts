import { describe, expect, it, vi } from "vitest";
import { ApiError, createApiClient } from "./client.ts";

const credentials = {
  load: () =>
    Promise.resolve({
      gasWebAppUrl: "https://script.google.com/macros/s/TEST/exec",
      apiKey: "KEY123",
    }),
};

const jsonResponse = (body: unknown) =>
  Promise.resolve({
    text: () => Promise.resolve(JSON.stringify(body)),
  } as Response);

const asFetch = (fn: unknown) => fn as typeof fetch;

describe("createApiClient", () => {
  it("sends a text/plain POST with the key in the body and follows redirects", async () => {
    const fetchImpl = vi.fn(() =>
      jsonResponse({ ok: true, data: { x: 1 }, warnings: ["w"] }),
    );
    const call = createApiClient({
      credentials,
      fetchImpl: asFetch(fetchImpl),
    });
    const result = await call<{ x: number }>({
      action: "getState",
      sheetName: "S",
    });
    expect(result).toEqual({ data: { x: 1 }, warnings: ["w"] });
    const [url, init] = fetchImpl.mock.calls[0] as unknown as [
      string,
      RequestInit,
    ];
    expect(url).toBe("https://script.google.com/macros/s/TEST/exec");
    expect(url).not.toContain("KEY123");
    expect(init.method).toBe("POST");
    expect(init.redirect).toBe("follow");
    expect(init.headers).toEqual({
      "Content-Type": "text/plain;charset=utf-8",
    });
    expect(JSON.parse(init.body as string)).toEqual({
      action: "getState",
      sheetName: "S",
      apiVersion: 1,
      apiKey: "KEY123",
    });
  });

  it("maps error responses to typed errors, including the conflict row", async () => {
    const current = { rowIndex: 3 };
    const fetchImpl = vi.fn(() =>
      jsonResponse({
        ok: false,
        error: { code: "conflict", message: "changed" },
        conflict: { current },
      }),
    );
    const call = createApiClient({
      credentials,
      fetchImpl: asFetch(fetchImpl),
    });
    const error = await call({ action: "syncReceipts" }).catch(
      (e: ApiError) => e,
    );
    expect(error).toBeInstanceOf(ApiError);
    expect((error as ApiError).code).toBe("conflict");
    expect((error as ApiError).conflictCurrent).toEqual(current);
  });

  it("explains a non-JSON (login page) response", async () => {
    const fetchImpl = vi.fn(() =>
      Promise.resolve({
        text: () => Promise.resolve("<html>Sign in</html>"),
      } as Response),
    );
    const call = createApiClient({
      credentials,
      fetchImpl: asFetch(fetchImpl),
    });
    await expect(call({ action: "syncReceipts" })).rejects.toMatchObject({
      code: "bad_response",
      message: expect.stringContaining("Anyone"),
    });
  });

  it("reports network failures as transient", async () => {
    const fetchImpl = vi.fn(() =>
      Promise.reject(new TypeError("Failed to fetch")),
    );
    const call = createApiClient({
      credentials,
      fetchImpl: asFetch(fetchImpl),
    });
    const error = (await call({ action: "syncReceipts" }).catch(
      (e) => e,
    )) as ApiError;
    expect(error.code).toBe("network");
    expect(error.isTransient).toBe(true);
  });

  it("times out", async () => {
    const fetchImpl = vi.fn(
      (_url: string, init: RequestInit) =>
        new Promise<Response>((_resolve, reject) => {
          init.signal?.addEventListener("abort", () =>
            reject(new DOMException("aborted", "AbortError")),
          );
        }),
    );
    const call = createApiClient({
      credentials,
      fetchImpl: asFetch(fetchImpl),
      timeoutMs: 10,
    });
    await expect(call({ action: "syncReceipts" })).rejects.toMatchObject({
      code: "timeout",
    });
  });

  it("requires credentials", async () => {
    const call = createApiClient({
      credentials: { load: () => Promise.resolve(null) },
      fetchImpl: asFetch(vi.fn()),
    });
    await expect(call({ action: "syncReceipts" })).rejects.toMatchObject({
      code: "not_configured",
    });
  });
});
