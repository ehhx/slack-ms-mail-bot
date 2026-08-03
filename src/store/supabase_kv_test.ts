import { assertEquals, assertRejects, assertStringIncludes } from "@std/assert";
import { SupabaseKvStore } from "./supabase_kv.ts";

interface RpcCall {
  name: string;
  body: Record<string, unknown>;
}

function createStore(
  handler: (name: string, body: Record<string, unknown>) => unknown,
): { store: SupabaseKvStore; calls: RpcCall[] } {
  const calls: RpcCall[] = [];
  const fetchImpl: typeof fetch = async (input, init) => {
    const request = new Request(input, init);
    const url = new URL(request.url);
    const name = url.pathname.split("/").pop() ?? "";
    const body = await request.json() as Record<string, unknown>;
    calls.push({ name, body });
    const response = handler(name, body);
    return new Response(JSON.stringify(response), {
      status: 200,
      headers: { "content-type": "application/json" },
    });
  };
  return {
    store: new SupabaseKvStore(
      "https://mail.supabase.co/",
      "service-role-test-key",
      1_000,
      fetchImpl,
    ),
    calls,
  };
}

Deno.test("Supabase KV encodes keys and preserves RPC result shapes", async () => {
  const { store, calls } = createStore((name, body) => {
    switch (name) {
      case "mail_state_get":
        return { value: { connected: true }, version: 7 };
      case "mail_state_get_many":
        return (body.p_state_keys as string[]).map((stateKey, index) => ({
          state_key: stateKey,
          state_value: { index },
          state_version: index + 1,
        }));
      case "mail_state_set_if_absent":
        return true;
      case "mail_state_list":
        return body.p_after === null
          ? [{
            state_key: "s:mailbox_connection/s:one%2Ftwo",
            state_value: { mailboxId: "one" },
            state_version: 2,
          }]
          : [];
      case "mail_state_prune_expired":
        return 3;
      default:
        return null;
    }
  });

  const single = await store.get<{ connected: boolean }>([
    "mailbox_connection",
    "one/two",
  ]);
  assertEquals(single.value, { connected: true });
  assertEquals(single.versionstamp, "7");
  assertEquals(calls[0].body.p_state_key, "s:mailbox_connection/s:one%2Ftwo");

  const many = await store.getMany([["a"], ["b"]]);
  assertEquals(many.map((entry) => entry.value), [{ index: 0 }, { index: 1 }]);

  const inserted = await store.setIfAbsent(
    ["oauth_state", "state-1"],
    { expiresAt: "future" },
    { expireIn: 5_000 },
  );
  assertEquals(inserted, true);
  const setIfAbsentCall = calls.find((call) =>
    call.name === "mail_state_set_if_absent"
  );
  assertStringIncludes(String(setIfAbsentCall?.body.p_expires_at), "T");

  const listed = [];
  for await (
    const entry of store.list<{ mailboxId: string }>({
      prefix: ["mailbox_connection"],
    })
  ) {
    listed.push(entry);
  }
  assertEquals(listed[0].key, ["mailbox_connection", "one/two"]);
  assertEquals(listed[0].value, { mailboxId: "one" });

  const atomic = store.atomic()
    .set(["mailbox_connection", "one"], { mailboxId: "one" })
    .delete(["mailbox_connection", "old"]);
  assertEquals(await atomic.commit(), { ok: true, versionstamp: null });
  assertEquals(
    calls.slice(-2).map((call) => call.name),
    ["mail_state_set", "mail_state_delete"],
  );
  assertEquals(await store.pruneExpired(3), 3);
});

Deno.test("Supabase KV reports RPC failures with operation context", async () => {
  const fetchImpl: typeof fetch = () =>
    Promise.resolve(new Response("database unavailable", { status: 503 }));
  const store = new SupabaseKvStore(
    "https://mail.supabase.co",
    "service-role-test-key",
    1_000,
    fetchImpl,
  );

  await assertRejects(
    () => store.get(["mailbox_connection", "one"]),
    Error,
    "Supabase state RPC mail_state_get failed: HTTP 503",
  );
});
