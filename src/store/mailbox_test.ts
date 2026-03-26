import { assert, assertEquals } from "@std/assert";
import {
  deleteMailbox,
  getMailboxBundle,
  hasDeliveredRecord,
  listMailboxBundles,
  resolveMailboxBundle,
  saveDeliveredRecord,
  saveMailboxBundle,
} from "./mailbox.ts";
import type { MailboxBundle } from "../mail/types.ts";

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
      junkFolderId: "junk-1",
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
      lastMessageReceivedAt: new Date().toISOString(),
      folderStates: {
        inbox: {
          folderId: "inbox-1",
          folderName: "Inbox",
          deltaLink: "delta-1",
          lastMessageReceivedAt: new Date().toISOString(),
        },
        junk: {
          folderId: "junk-1",
          folderName: "Junk",
          deltaLink: "delta-junk-1",
          lastMessageReceivedAt: new Date().toISOString(),
        },
      },
      lastSyncAt: new Date().toISOString(),
      updatedAt: new Date().toISOString(),
    },
    lease: {
      mailboxId: "mailbox-1",
      resource: "/me/messages",
      clientState: "client-state",
      subscriptionId: "sub-1",
      expiresAt: new Date(Date.now() + 3600_000).toISOString(),
      status: "active",
      updatedAt: new Date().toISOString(),
    },
  };
}

Deno.test("mailbox repository saves and resolves bundles", async () => {
  const dir = await Deno.makeTempDir();
  const kv = await Deno.openKv(`${dir}/kv.sqlite`);
  const bundle = sampleBundle();

  await saveMailboxBundle(kv, bundle);

  const listed = await listMailboxBundles(kv, "T1");
  assertEquals(listed.length, 1);

  const resolved = await resolveMailboxBundle(kv, "T1", "mailbox@example.com");
  assert(resolved);
  assertEquals(resolved.connection.mailboxId, "mailbox-1");

  await saveDeliveredRecord(kv, {
    mailboxId: "mailbox-1",
    dedupeKey: "dedupe-1",
    messageId: "msg-1",
    subject: "Hello",
    slackChannelId: "C1",
    deliveredAt: new Date().toISOString(),
  });
  assertEquals(await hasDeliveredRecord(kv, "mailbox-1", "dedupe-1"), true);

  const fetched = await getMailboxBundle(kv, "mailbox-1");
  assert(fetched);
  assertEquals(fetched.route?.slackChannelId, "C1");

  await deleteMailbox(kv, "mailbox-1");
  assertEquals(await getMailboxBundle(kv, "mailbox-1"), null);

  (kv as { close?: () => void }).close?.();
});
