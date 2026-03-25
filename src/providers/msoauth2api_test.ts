import { assertEquals, assertRejects } from "@std/assert";
import type { AppConfig } from "../config.ts";
import { fetchMsOauth2ApiMessages, MsOauth2ApiError } from "./msoauth2api.ts";

const config: AppConfig = {
  slackSigningSecret: "secret",
  slackBotToken: "xoxb-test",
  appBaseUrl: "https://example.com",
  kvPath: null,
  slackApiTimeoutMs: 15000,
  mailPreviewMaxChars: 220,
  graphApiBaseUrl: "https://graph.microsoft.com/v1.0",
  microsoftClientId: "client-id",
  microsoftClientSecret: "client-secret",
  microsoftRedirectUri: "https://example.com/oauth/microsoft/callback",
  microsoftAuthTenant: "common",
  tokenEncryptionKey: "encryption-key",
  webhookClientState: "client-state",
  graphSubscriptionRenewalWindowMinutes: 180,
  graphSubscriptionMaxMinutes: 4230,
  mailSyncPollIntervalMinutes: 15,
  mailProviderDefault: "graph_native",
  msOauth2apiBaseUrl: "https://ms-oauth2api.example.com",
  msOauth2apiPassword: "password",
  msOauth2apiMailbox: "INBOX",
};

Deno.test("fetchMsOauth2ApiMessages maps response and sorts by received time", async () => {
  const messages = await fetchMsOauth2ApiMessages({
    config,
    refreshToken: "refresh-token",
    emailAddress: "mailbox@example.com",
    fetchImpl: async () =>
      new Response(JSON.stringify([
        {
          send: "later@example.com",
          subject: "Later",
          text: "later body",
          date: "2026-03-25T02:00:00.000Z",
        },
        {
          send: "earlier@example.com",
          subject: "Earlier",
          text: "earlier body",
          date: "2026-03-25T01:00:00.000Z",
        },
      ])),
  });

  assertEquals(messages.length, 2);
  assertEquals(messages[0].subject, "Earlier");
  assertEquals(messages[1].subject, "Later");
  assertEquals(Boolean(messages[0].messageId), true);
  assertEquals(messages[0].messageId === messages[1].messageId, false);
});

Deno.test("fetchMsOauth2ApiMessages throws typed error on http failure", async () => {
  await assertRejects(
    () =>
      fetchMsOauth2ApiMessages({
        config,
        refreshToken: "refresh-token",
        emailAddress: "mailbox@example.com",
        fetchImpl: async () => new Response("forbidden", { status: 403 }),
      }),
    MsOauth2ApiError,
    "HTTP 403",
  );
});
