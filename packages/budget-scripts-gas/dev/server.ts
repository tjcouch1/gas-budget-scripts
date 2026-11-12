/**
 * Node HTTP server that exposes the demo backend like a deployed GAS Web App, plus `/__debug/*`
 * endpoints for tests (these do not exist on real GAS). Permissive CORS, text/plain JSON bodies.
 */
import {
  createServer,
  type IncomingMessage,
  type Server,
  type ServerResponse,
} from "node:http";
import { DemoBackend } from "./backend.ts";

const CORS_HEADERS = {
  "Access-Control-Allow-Origin": "*",
  "Access-Control-Allow-Methods": "GET, POST, OPTIONS",
  "Access-Control-Allow-Headers": "Content-Type",
};

/** Path that stands in for the GAS deployment URL path (`/macros/s/<id>/exec`) */
export const EXEC_PATH = "/macros/s/demo/exec";

function readBody(req: IncomingMessage): Promise<string> {
  return new Promise((resolve, reject) => {
    const chunks: Buffer[] = [];
    req.on("data", (chunk: Buffer) => chunks.push(chunk));
    req.on("end", () => resolve(Buffer.concat(chunks).toString("utf8")));
    req.on("error", reject);
  });
}

function send(
  res: ServerResponse,
  status: number,
  body: string,
  contentType = "application/json; charset=utf-8",
) {
  res.writeHead(status, { ...CORS_HEADERS, "Content-Type": contentType });
  res.end(body);
}

export function createDemoServer(backend: DemoBackend): Server {
  return createServer((req, res) => {
    void handle(backend, req, res).catch((error: unknown) => {
      send(
        res,
        500,
        JSON.stringify({
          error: error instanceof Error ? error.message : String(error),
        }),
      );
    });
  });
}

async function handle(
  backend: DemoBackend,
  req: IncomingMessage,
  res: ServerResponse,
) {
  const url = new URL(req.url ?? "/", "http://localhost");
  const method = req.method ?? "GET";

  if (method === "OPTIONS") {
    res.writeHead(204, CORS_HEADERS);
    res.end();
    return;
  }

  if (url.pathname.endsWith("/exec")) {
    if (method === "POST") send(res, 200, backend.post(await readBody(req)));
    else send(res, 200, backend.get());
    return;
  }

  if (url.pathname.startsWith("/__debug/")) {
    const route = url.pathname.slice("/__debug/".length);
    const json = (data: unknown) => send(res, 200, JSON.stringify(data));
    const body = async () =>
      JSON.parse((await readBody(req)) || "{}") as Record<string, unknown>;
    switch (`${method} ${route}`) {
      case "GET health":
        return json({ ok: true });
      case "POST reset":
        backend.reset();
        return json({ ok: true });
      case "GET sheets":
        return json(backend.listSheets());
      case "GET sheet":
        return json(backend.dumpSheet(url.searchParams.get("name") ?? ""));
      case "POST mutate":
        backend.mutateRow(await body());
        return json({ ok: true });
      case "POST receipts":
        backend.addThreads(
          ((await body()).threads as Record<string, unknown>[]) ?? [],
        );
        return json({ ok: true });
      case "GET threads":
        return json(backend.dumpThreads());
      case "POST lock":
        backend.setLockContended(Boolean((await body()).contended));
        return json({ ok: true });
    }
  }

  send(res, 404, JSON.stringify({ error: "Not found" }));
}
