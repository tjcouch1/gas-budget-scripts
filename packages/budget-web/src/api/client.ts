import {
  API_VERSION,
  type ApiErrorCode,
  type ApiRequest,
  type ApiResponse,
  type TransactionRowDto,
} from "@budget/shared";
import type { CredentialSource } from "../crypto/credential-vault.ts";

export type ApiErrorKind =
  ApiErrorCode | "network" | "timeout" | "bad_response" | "not_configured";

export class ApiError extends Error {
  constructor(
    readonly code: ApiErrorKind,
    message: string,
    /** The server's current row, present on `conflict`. */
    readonly conflictCurrent?: TransactionRowDto,
  ) {
    super(message);
    this.name = "ApiError";
  }
  /** Worth retrying later without user action. */
  get isTransient(): boolean {
    return (
      this.code === "network" || this.code === "timeout" || this.code === "busy"
    );
  }
}

/** A request without the envelope fields, which the client adds from the vault. */
export type ApiRequestBody = ApiRequest extends infer R
  ? R extends ApiRequest
    ? Omit<R, "apiVersion" | "apiKey">
    : never
  : never;

export interface ApiResult<T> {
  data: T;
  warnings: string[];
}

export type CallApi = <T>(request: ApiRequestBody) => Promise<ApiResult<T>>;

export interface ApiClientOptions {
  credentials: CredentialSource;
  fetchImpl?: typeof fetch;
  timeoutMs?: number;
}

export function createApiClient(options: ApiClientOptions): CallApi {
  const timeoutMs = options.timeoutMs ?? 60_000;
  return async function callApi<T>(request: ApiRequestBody) {
    // Credentials are decrypted only for the duration of this request.
    const credentials = await options.credentials.load();
    if (!credentials) {
      throw new ApiError(
        "not_configured",
        "Open Settings and enter the web app URL and API key.",
      );
    }
    const fetchFn = options.fetchImpl ?? globalThis.fetch.bind(globalThis);
    const controller = new AbortController();
    const timer = setTimeout(() => controller.abort(), timeoutMs);
    let text: string;
    try {
      const response = await fetchFn(credentials.gasWebAppUrl, {
        method: "POST",
        // text/plain avoids a CORS preflight, which script.google.com does not answer.
        headers: { "Content-Type": "text/plain;charset=utf-8" },
        // The API key travels only in the body, never in the URL.
        body: JSON.stringify({
          ...request,
          apiVersion: API_VERSION,
          apiKey: credentials.apiKey,
        }),
        redirect: "follow",
        signal: controller.signal,
      });
      text = await response.text();
    } catch (error) {
      if (controller.signal.aborted) {
        throw new ApiError("timeout", "The request timed out.");
      }
      throw new ApiError(
        "network",
        error instanceof Error
          ? `Network error: ${error.message}`
          : "Network error.",
      );
    } finally {
      clearTimeout(timer);
    }
    let parsed: ApiResponse<T>;
    try {
      parsed = JSON.parse(text) as ApiResponse<T>;
    } catch {
      throw new ApiError(
        "bad_response",
        "The server did not return JSON. Check the URL, and that the web app is deployed with access set to Anyone.",
      );
    }
    if (
      typeof parsed !== "object" ||
      parsed === null ||
      typeof parsed.ok !== "boolean"
    ) {
      throw new ApiError(
        "bad_response",
        "Unexpected response from the server.",
      );
    }
    if (!parsed.ok) {
      throw new ApiError(
        parsed.error.code,
        parsed.error.message,
        parsed.conflict?.current,
      );
    }
    return { data: parsed.data, warnings: parsed.warnings ?? [] };
  };
}
