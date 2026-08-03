import { assert, assertEquals } from "@std/assert";
import { importLegacyDenoKvPageFromStore } from "./migration.ts";
import { SupabaseKvStore } from "./supabase_kv.ts";

Deno.test("legacy KV migration is cursor-based and preserves TTL", async () => {
  const directory = await Deno.makeTempDir({ prefix: "mail-kv-migration-" });
  const legacyKv = await Deno.openKv(`${directory}/legacy.sqlite`);
  const calls: Array<Record<string, unknown>> = [];
  const fetchImpl: typeof fetch = async (input, init) => {
    const request = new Request(input, init);
    const body = await request.json() as Record<string, unknown>;
    calls.push(body);
    return new Response("true", {
      status: 200,
      headers: { "content-type": "application/json" },
    });
  };
  const target = new SupabaseKvStore(
    "https://mail.supabase.co",
    "service-role-test-key",
    1_000,
    fetchImpl,
  );

  try {
    await legacyKv.set(["mailbox_connection", "one"], { mailboxId: "one" });
    await legacyKv.set(["mailbox_connection", "two"], { mailboxId: "two" });
    const first = await importLegacyDenoKvPageFromStore(
      legacyKv,
      target,
      "mailbox_connection",
      undefined,
      1,
    );
    assertEquals(first.imported, 1);
    assert(first.nextCursor !== null);

    const second = await importLegacyDenoKvPageFromStore(
      legacyKv,
      target,
      "mailbox_connection",
      first.nextCursor ?? undefined,
      1,
    );
    assertEquals(second.imported, 1);
    assert(second.nextCursor !== null);

    const completed = await importLegacyDenoKvPageFromStore(
      legacyKv,
      target,
      "mailbox_connection",
      second.nextCursor ?? undefined,
      1,
    );
    assertEquals(completed.imported, 0);
    assertEquals(completed.nextCursor, null);

    const future = new Date(Date.now() + 60_000).toISOString();
    await legacyKv.set(["oauth_state", "future"], { expiresAt: future });
    const oauth = await importLegacyDenoKvPageFromStore(
      legacyKv,
      target,
      "oauth_state",
      undefined,
      10,
    );
    assertEquals(oauth.imported, 1);
    const oauthCall = calls.find((body) =>
      body.p_state_key === "s:oauth_state/s:future"
    );
    assert(oauthCall);
    assertEquals(typeof oauthCall.p_expires_at, "string");

    await legacyKv.set(["delivered_mail", "old"], {
      deliveredAt: new Date(Date.now() - 91 * 24 * 60 * 60 * 1000)
        .toISOString(),
    });
    const delivered = await importLegacyDenoKvPageFromStore(
      legacyKv,
      target,
      "delivered_mail",
      undefined,
      10,
    );
    assertEquals(delivered.imported, 0);
    assertEquals(delivered.skipped, 1);
  } finally {
    legacyKv.close();
    await Deno.remove(directory, { recursive: true });
  }
});
