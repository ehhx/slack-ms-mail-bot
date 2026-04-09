import { assertEquals } from "@std/assert";
import type { AppConfig } from "../config.ts";
import {
  buildWebAdminSessionCookie,
  isWebAdminAuthenticated,
  verifyWebAdminPassword,
} from "./auth.ts";

const config: AppConfig = {
  slackSigningSecret: "secret",
  slackBotToken: "xoxb-test",
  appBaseUrl: "https://example.com",
  kvPath: null,
  slackApiTimeoutMs: 15000,
  mailPreviewMaxChars: 800,
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
  msOauth2apiMailboxes: ["INBOX", "Junk"],
  webAdminPassword: "top-secret",
  webSessionSecret: "web-session-secret",
};

Deno.test("verifyWebAdminPassword validates configured password", async () => {
  assertEquals(await verifyWebAdminPassword("top-secret", config), true);
  assertEquals(await verifyWebAdminPassword("bad-password", config), false);
});

Deno.test("isWebAdminAuthenticated validates signed session cookie", async () => {
  const setCookie = await buildWebAdminSessionCookie(config);
  const cookieHeader = setCookie.split(";")[0];
  const request = new Request("https://example.com/app", {
    headers: { cookie: cookieHeader },
  });

  assertEquals(await isWebAdminAuthenticated(request, config), true);
});
