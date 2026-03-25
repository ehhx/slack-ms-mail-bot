import { assertEquals } from "@std/assert";
import { clearConfigCache } from "../config.ts";
import { setKvForTesting } from "../store/kv.ts";
import { getOAuthState, listSyncJobs, saveMailboxBundle } from "../store/mailbox.ts";
import {
  createConnectUrl,
  processGraphNotifications,
  queueMailboxSyncByMailboxRef,
} from "./service.ts";
import type { MailboxBundle } from "./types.ts";

function setEnv(): void {
  Deno.env.set("SLACK_SIGNING_SECRET", "secret");
  Deno.env.set("SLACK_BOT_TOKEN", "xoxb-test");
  Deno.env.set("APP_BASE_URL", "https://example.com");
  Deno.env.set("MICROSOFT_CLIENT_ID", "client-id");
  Deno.env.set("MICROSOFT_CLIENT_SECRET", "client-secret");
  Deno.env.set("MICROSOFT_REDIRECT_URI", "https://example.com/oauth/microsoft/callback");
  Deno.env.set("TOKEN_ENCRYPTION_KEY", "super-secret");
}

function sampleBundle(): MailboxBundle {
  return {
    connection: {
      mailboxId: "mailbox-1",
      teamId: "T1",
      authorizedByUserId: "U1",
      graphUserId: "G1",
      emailAddress: "mailbox@example.com",
      displayName: "Mailbox",
      inboxFolderId: "inbox-1",
      encryptedRefreshToken: "encrypted",
      accessTokenExpiresAt: new Date().toISOString(),
      providerType: "graph_native",
      createdAt: new Date().toISOString(),
      updatedAt: new Date().toISOString(),
      status: "active",
    },
    route: {
      mailboxId: "mailbox-1",
      slackChannelId: "C1",
      updatedAt: new Date().toISOString(),
    },
    syncState: {
      mailboxId: "mailbox-1",
      deltaLink: "delta-1",
      lastSyncAt: new Date().toISOString(),
      updatedAt: new Date().toISOString(),
    },
    lease: {
      mailboxId: "mailbox-1",
      resource: "/me/mailFolders/inbox/messages",
      clientState: "expected-client-state",
      subscriptionId: "sub-1",
      expiresAt: new Date(Date.now() + 3600_000).toISOString(),
      status: "active",
      updatedAt: new Date().toISOString(),
    },
  };
}

Deno.test("queueMailboxSyncByMailboxRef resolves email to mailbox id", async () => {
  setEnv();
  Deno.env.set("GRAPH_WEBHOOK_CLIENT_STATE", "expected-client-state");
  clearConfigCache();
  const dir = await Deno.makeTempDir();
  const kv = await Deno.openKv(`${dir}/kv.sqlite`);
  setKvForTesting(kv);
  await saveMailboxBundle(kv, sampleBundle());

  const bundle = await queueMailboxSyncByMailboxRef({
    teamId: "T1",
    mailbox: "mailbox@example.com",
    reason: "test",
  });
  assertEquals(bundle.connection.mailboxId, "mailbox-1");
  const jobs = await listSyncJobs(kv);
  assertEquals(jobs.length, 1);
  assertEquals(jobs[0].mailboxId, "mailbox-1");

  setKvForTesting(null);
  (kv as { close?: () => void }).close?.();
});

Deno.test("createConnectUrl stores requested provider type in OAuth state", async () => {
  setEnv();
  clearConfigCache();
  const dir = await Deno.makeTempDir();
  const kv = await Deno.openKv(`${dir}/kv.sqlite`);
  setKvForTesting(kv);

  const result = await createConnectUrl({
    teamId: "T1",
    userId: "U1",
    channelId: "C1",
    providerType: "ms_oauth2api",
  });
  const authorizeUrl = new URL(result.authorizeUrl);
  const state = authorizeUrl.searchParams.get("state");

  assertEquals(result.providerType, "ms_oauth2api");
  assertEquals(Boolean(state), true);
  const oauthState = await getOAuthState(kv, state!);
  assertEquals(oauthState?.providerType, "ms_oauth2api");

  setKvForTesting(null);
  (kv as { close?: () => void }).close?.();
});

Deno.test("processGraphNotifications queues valid subscriptions", async () => {
  setEnv();
  Deno.env.set("GRAPH_WEBHOOK_CLIENT_STATE", "expected-client-state");
  clearConfigCache();
  const dir = await Deno.makeTempDir();
  const kv = await Deno.openKv(`${dir}/kv.sqlite`);
  setKvForTesting(kv);
  await saveMailboxBundle(kv, sampleBundle());

  const result = await processGraphNotifications([
    { subscriptionId: "sub-1", clientState: "expected-client-state", changeType: "created" },
    { subscriptionId: "sub-1", clientState: "wrong", changeType: "created" },
  ]);

  assertEquals(result.queued, 1);
  assertEquals(result.ignored, 1);
  const jobs = await listSyncJobs(kv);
  assertEquals(jobs.length, 1);

  setKvForTesting(null);
  (kv as { close?: () => void }).close?.();
});

Deno.test("processGraphNotifications ignores non-graph provider mailboxes", async () => {
  setEnv();
  Deno.env.set("GRAPH_WEBHOOK_CLIENT_STATE", "expected-client-state");
  clearConfigCache();
  const dir = await Deno.makeTempDir();
  const kv = await Deno.openKv(`${dir}/kv.sqlite`);
  setKvForTesting(kv);

  const bundle = sampleBundle();
  bundle.connection.providerType = "ms_oauth2api";
  await saveMailboxBundle(kv, bundle);

  const result = await processGraphNotifications([
    { subscriptionId: "sub-1", clientState: "expected-client-state", changeType: "created" },
  ]);

  assertEquals(result.queued, 0);
  assertEquals(result.ignored, 1);
  assertEquals((await listSyncJobs(kv)).length, 0);

  setKvForTesting(null);
  (kv as { close?: () => void }).close?.();
});
