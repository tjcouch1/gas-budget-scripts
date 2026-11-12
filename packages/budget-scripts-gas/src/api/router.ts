/**
 * Web App entry points (web-plan.md section 4.4). GAS cannot read request headers or set HTTP
 * status codes, so every response is JSON with `ok` and errors are reported in the body.
 */
import {
  API_VERSION,
  type ApiFailure,
  type ApiResponse,
  type ServiceInfoDto,
} from "@budget/shared";
import { ApiError, getErrorMessage } from "./api-error";
import { isAuthorized } from "./auth";
import {
  handleGetState,
  handleSplitTransaction,
  handleSyncReceipts,
  handleUpdateTransaction,
  type HandlerResult,
} from "./handlers";

function toJsonOutput(response: ApiResponse<unknown>) {
  return ContentService.createTextOutput(JSON.stringify(response)).setMimeType(
    ContentService.MimeType.JSON,
  );
}

function toFailure(e: unknown): ApiFailure {
  if (e instanceof ApiError) {
    const failure: ApiFailure = {
      ok: false,
      error: { code: e.code, message: e.message },
    };
    if (e.conflict) failure.conflict = e.conflict;
    return failure;
  }
  // Unexpected error: report only the message, never a stack trace
  const message = getErrorMessage(e);
  Logger.log(`API server error: ${message}`);
  return { ok: false, error: { code: "server_error", message } };
}

/** Parses, authenticates, and dispatches a POST body. Never throws */
export function handleApiRequest(body: unknown): ApiResponse<unknown> {
  try {
    if (typeof body !== "string")
      throw new ApiError("bad_request", "Missing request body");
    let request: unknown;
    try {
      request = JSON.parse(body);
    } catch {
      throw new ApiError("bad_request", "Request body must be JSON");
    }
    if (
      typeof request !== "object" ||
      request === null ||
      Array.isArray(request)
    )
      throw new ApiError("bad_request", "Request body must be a JSON object");
    const req = request as Record<string, unknown>;

    // Authenticate before looking at anything else in the request
    if (!isAuthorized(req.apiKey))
      throw new ApiError("unauthorized", "Invalid or missing API key");

    if (req.apiVersion !== API_VERSION)
      throw new ApiError(
        "bad_request",
        `Unsupported apiVersion; expected ${API_VERSION}`,
      );

    let result: HandlerResult<unknown>;
    switch (req.action) {
      case "getState":
        result = handleGetState(req);
        break;
      case "updateTransaction":
        result = handleUpdateTransaction(req);
        break;
      case "splitTransaction":
        result = handleSplitTransaction(req);
        break;
      case "syncReceipts":
        result = handleSyncReceipts();
        break;
      default:
        throw new ApiError("bad_request", "Unknown action");
    }

    return result.warnings && result.warnings.length > 0
      ? { ok: true, data: result.data, warnings: result.warnings }
      : { ok: true, data: result.data };
  } catch (e) {
    return toFailure(e);
  }
}

/** Web App POST entry point: every API call (body = JSON request, see `@budget/shared` ApiRequest) */
export function doPost(e: GoogleAppsScript.Events.DoPost) {
  return toJsonOutput(handleApiRequest(e?.postData?.contents));
}

/** Web App GET entry point: unauthenticated connectivity check only. Returns no data */
export function doGet(_e?: GoogleAppsScript.Events.DoGet) {
  const data: ServiceInfoDto = {
    service: "budget-api",
    apiVersion: API_VERSION,
  };
  return toJsonOutput({ ok: true, data });
}
