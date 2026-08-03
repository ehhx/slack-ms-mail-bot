import { getConfigAsync } from "./config.ts";
import type { WaitUntilLike } from "./runtime.ts";
import { runBackground } from "./runtime.ts";
import {
  completeOAuthCallback,
  processGraphNotifications,
  processQueuedSyncs,
} from "./mail/service.ts";
import {
  getGraphValidationToken,
  parseGraphWebhookBody,
} from "./microsoft/webhook.ts";
import { handleWebRequest } from "./web/handler.ts";
import { handleLarkEvent } from "./lark/handler.ts";
import { getPersistenceStatus } from "./store/kv.ts";
import {
  importLegacyDenoKvPage,
  isLegacyDenoKvMigrationKind,
  persistenceBackendSupportsLegacyMigration,
} from "./store/migration.ts";

function htmlPage(title: string, body: string): Response {
  return new Response(
    `<!doctype html><html><head><meta charset="utf-8"><title>${title}</title></head><body><h1>${title}</h1><p>${body}</p></body></html>`,
    { status: 200, headers: { "content-type": "text/html; charset=utf-8" } },
  );
}

export async function handleRequest(
  request: Request,
  ctx?: WaitUntilLike,
): Promise<Response> {
  const url = new URL(request.url);

  const webResponse = await handleWebRequest(request);
  if (webResponse) {
    return webResponse;
  }

  if (url.pathname === "/healthz") {
    const persistence = await getPersistenceStatus();
    if (!persistence.available) {
      console.error("Persistence health check failed", persistence.error);
    }
    return Response.json(
      {
        ok: persistence.available,
        persistence: persistence.backend,
      },
      { status: persistence.available ? 200 : 503 },
    );
  }

  if (
    request.method === "GET" && url.pathname === "/oauth/microsoft/callback"
  ) {
    const code = url.searchParams.get("code");
    const state = url.searchParams.get("state");
    const error = url.searchParams.get("error");
    if (error) {
      return htmlPage(
        "Microsoft OAuth failed",
        `Authorization failed: ${error}`,
      );
    }
    if (!code || !state) {
      return new Response("Missing OAuth code/state", { status: 400 });
    }

    try {
      const bundle = await completeOAuthCallback(code, state);
      return htmlPage(
        "Mailbox connected",
        `Connected ${bundle.connection.emailAddress}. You can return to Lark or open /app after logging in.`,
      );
    } catch (oauthError) {
      console.error("OAuth callback failed", oauthError);
      return new Response("OAuth callback failed", { status: 500 });
    }
  }

  if (url.pathname === "/admin/migrate-deno-kv") {
    if (request.method !== "POST") {
      return new Response("Method Not Allowed", {
        status: 405,
        headers: { Allow: "POST" },
      });
    }
    const config = await getConfigAsync();
    const expectedToken = config.legacyDenoKvMigrationToken;
    if (
      !expectedToken ||
      request.headers.get("x-kv-migration-token") !== expectedToken
    ) {
      return new Response("Forbidden", { status: 403 });
    }
    if (!persistenceBackendSupportsLegacyMigration(config.persistenceBackend)) {
      return Response.json(
        {
          ok: false,
          error: "PERSISTENCE_BACKEND must be dual_write or supabase",
        },
        { status: 409 },
      );
    }

    let body: { kind?: unknown; cursor?: unknown };
    try {
      body = await request.json();
    } catch {
      return Response.json(
        { ok: false, error: "Invalid JSON body" },
        { status: 400 },
      );
    }
    if (!isLegacyDenoKvMigrationKind(body.kind)) {
      return Response.json(
        { ok: false, error: "Invalid migration kind" },
        { status: 400 },
      );
    }
    if (body.cursor != null && typeof body.cursor !== "string") {
      return Response.json(
        { ok: false, error: "cursor must be a string" },
        { status: 400 },
      );
    }

    try {
      const result = await importLegacyDenoKvPage(
        {
          kind: body.kind,
          cursor: body.cursor as string | undefined,
        },
        config.legacyDenoKvMigrationBatchSize,
      );
      return Response.json({ ok: true, ...result });
    } catch (error) {
      console.error("Legacy Deno KV migration failed", error);
      return Response.json(
        { ok: false, error: "Legacy Deno KV migration failed" },
        { status: 500 },
      );
    }
  }

  if (request.method !== "POST") {
    return new Response("Not found", { status: 404 });
  }

  if (url.pathname === "/lark/events") {
    return await handleLarkEvent(request, ctx);
  }

  if (url.pathname === "/graph/webhook") {
    const validationToken = getGraphValidationToken(request);
    if (validationToken) {
      return new Response(validationToken, {
        status: 200,
        headers: { "content-type": "text/plain; charset=utf-8" },
      });
    }

    const bodyText = await request.text();
    let envelope;
    try {
      envelope = parseGraphWebhookBody(bodyText || "{}");
    } catch {
      return new Response("Invalid webhook payload", { status: 400 });
    }

    runBackground(
      ctx,
      (async () => {
        const result = await processGraphNotifications(envelope.value);
        console.log("Graph notifications queued", result);
        await processQueuedSyncs();
      })(),
    );

    return new Response("accepted", { status: 202 });
  }

  return new Response("Not found", { status: 404 });
}
