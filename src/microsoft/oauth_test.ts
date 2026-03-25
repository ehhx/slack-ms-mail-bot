import { assertEquals, assertStringIncludes } from "@std/assert";
import { buildMicrosoftAuthorizeUrl } from "./oauth.ts";
import type { AppConfig } from "../config.ts";

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
  msOauth2apiBaseUrl: null,
  msOauth2apiPassword: null,
  msOauth2apiMailbox: "INBOX",
};

Deno.test("buildMicrosoftAuthorizeUrl includes OAuth params", () => {
  const url = new URL(buildMicrosoftAuthorizeUrl(config, "state-123"));
  assertEquals(url.searchParams.get("client_id"), "client-id");
  assertEquals(url.searchParams.get("response_type"), "code");
  assertEquals(url.searchParams.get("state"), "state-123");
  assertStringIncludes(url.searchParams.get("scope") ?? "", "Mail.Read");
});
