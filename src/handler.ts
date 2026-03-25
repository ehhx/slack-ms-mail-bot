import { getConfigAsync } from "./config.ts";
import { handleSlackCommands } from "./slack/handlers/commands.ts";
import { handleSlackInteractivity } from "./slack/handlers/interactivity.ts";
import { verifySlackRequest } from "./slack/verify.ts";
import type { WaitUntilLike } from "./runtime.ts";
import { runBackground } from "./runtime.ts";
import { completeOAuthCallback, processGraphNotifications, processQueuedSyncs } from "./mail/service.ts";
import { getGraphValidationToken, parseGraphWebhookBody } from "./microsoft/webhook.ts";

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

  if (url.pathname === "/healthz") {
    return new Response("ok", { status: 200 });
  }

  if (request.method === "GET" && url.pathname === "/oauth/microsoft/callback") {
    const code = url.searchParams.get("code");
    const state = url.searchParams.get("state");
    const error = url.searchParams.get("error");
    if (error) {
      return htmlPage("Microsoft OAuth failed", `Authorization failed: ${error}`);
    }
    if (!code || !state) {
      return new Response("Missing OAuth code/state", { status: 400 });
    }

    try {
      const bundle = await completeOAuthCallback(code, state);
      return htmlPage(
        "Mailbox connected",
        `Connected ${bundle.connection.emailAddress}. You can return to Slack.`,
      );
    } catch (oauthError) {
      console.error("OAuth callback failed", oauthError);
      return new Response("OAuth callback failed", { status: 500 });
    }
  }

  if (request.method !== "POST") {
    return new Response("Not found", { status: 404 });
  }

  if (url.pathname === "/slack/commands") {
    const bodyText = await request.text();
    const config = await getConfigAsync();
    const verified = await verifySlackRequest(request, bodyText, config.slackSigningSecret);
    if (!verified.ok) return new Response(verified.error, { status: 401 });
    return await handleSlackCommands(bodyText, ctx);
  }

  if (url.pathname === "/slack/interactivity") {
    const bodyText = await request.text();
    const config = await getConfigAsync();
    const verified = await verifySlackRequest(request, bodyText, config.slackSigningSecret);
    if (!verified.ok) return new Response(verified.error, { status: 401 });
    return await handleSlackInteractivity(bodyText, ctx);
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

    runBackground(ctx, (async () => {
      const result = await processGraphNotifications(envelope.value);
      console.log("Graph notifications queued", result);
      await processQueuedSyncs();
    })());

    return new Response("accepted", { status: 202 });
  }

  return new Response("Not found", { status: 404 });
}
