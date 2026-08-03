import { assertEquals, assertStringIncludes } from "@std/assert";
import { clearConfigCache } from "../config.ts";
import { clearLarkTokenCache, postLarkCard } from "./api.ts";

function setEnv(): void {
  Deno.env.set("LARK_APP_ID", "cli_test");
  Deno.env.set("LARK_APP_SECRET", "app-secret");
  Deno.env.set("LARK_VERIFICATION_TOKEN", "verification-token");
  Deno.env.set("LARK_API_BASE_URL", "https://open.larksuite.test/open-apis");
  Deno.env.set("APP_BASE_URL", "https://mail.example.com");
  Deno.env.set("MICROSOFT_CLIENT_ID", "client-id");
  Deno.env.set("MICROSOFT_CLIENT_SECRET", "client-secret");
  Deno.env.set(
    "MICROSOFT_REDIRECT_URI",
    "https://mail.example.com/oauth/microsoft/callback",
  );
  Deno.env.set("TOKEN_ENCRYPTION_KEY", "encryption-key");
  clearConfigCache();
  clearLarkTokenCache();
}

Deno.test("postLarkCard obtains a tenant token and sends an interactive message", async () => {
  setEnv();
  const requests: Array<
    { url: string; body: Record<string, unknown>; authorization?: string }
  > = [];
  const result = await postLarkCard({
    chatId: "oc_chat_1",
    card: { header: { title: "Mail" }, elements: [] },
    idempotencyKey: "event-1",
    fetchImpl: async (input, init) => {
      const request = new Request(input, init);
      const url = request.url;
      const body = JSON.parse(await request.text()) as Record<
        string,
        unknown
      >;
      requests.push({
        url,
        body,
        authorization: request.headers.get("authorization") ??
          undefined,
      });
      if (url.endsWith("/auth/v3/tenant_access_token/internal")) {
        return Promise.resolve(
          new Response(JSON.stringify({
            code: 0,
            tenant_access_token: "tenant-token",
            expire: 7200,
          })),
        );
      }
      return Promise.resolve(
        new Response(JSON.stringify({
          code: 0,
          data: { message_id: "om_message_1" },
        })),
      );
    },
  });

  assertEquals(result.messageId, "om_message_1");
  assertEquals(requests.length, 2);
  assertEquals(requests[0].body.app_id, "cli_test");
  assertEquals(requests[1].authorization, "Bearer tenant-token");
  assertEquals(requests[1].body.receive_id, "oc_chat_1");
  assertEquals(requests[1].body.msg_type, "interactive");
  assertEquals(requests[1].body.uuid, "event-1");
  assertStringIncludes(String(requests[1].body.content), "Mail");
});
