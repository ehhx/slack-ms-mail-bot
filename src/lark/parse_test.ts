import { assertEquals, assertThrows } from "@std/assert";
import {
  getLarkUrlVerificationChallenge,
  getLarkVerificationToken,
  parseLarkMailCommand,
  parseLarkPayload,
  parseLarkTextMessageEvent,
} from "./parse.ts";

Deno.test("parseLarkMailCommand parses management commands", () => {
  assertEquals(parseLarkMailCommand("mail connect graph"), {
    kind: "connect",
    providerType: "graph_native",
  });
  assertEquals(parseLarkMailCommand("@_user_1 mail route a@outlook.com"), {
    kind: "route",
    mailbox: "a@outlook.com",
  });
  assertEquals(
    parseLarkMailCommand("/mail provider a@outlook.com msoauth2api"),
    {
      kind: "provider",
      mailbox: "a@outlook.com",
      providerType: "ms_oauth2api",
    },
  );
  assertEquals(parseLarkMailCommand("hello"), null);
});

Deno.test("parseLarkTextMessageEvent reads Event V2 text messages", () => {
  const payload = parseLarkPayload(JSON.stringify({
    header: {
      event_type: "im.message.receive_v1",
      tenant_key: "tenant-1",
      token: "verify-me",
    },
    event: {
      sender: {
        sender_type: "user",
        sender_id: { open_id: "ou_user_1" },
      },
      message: {
        chat_id: "oc_chat_1",
        chat_type: "group",
        message_type: "text",
        content: JSON.stringify({ text: "mail status" }),
      },
    },
  }));

  assertEquals(getLarkVerificationToken(payload), "verify-me");
  assertEquals(parseLarkTextMessageEvent(payload), {
    tenantId: "tenant-1",
    senderOpenId: "ou_user_1",
    chatId: "oc_chat_1",
    chatType: "group",
    text: "mail status",
  });
});

Deno.test("Lark URL verification and encrypted payload handling", () => {
  const payload = parseLarkPayload(JSON.stringify({
    type: "url_verification",
    token: "verify-me",
    challenge: "challenge-value",
  }));
  assertEquals(getLarkUrlVerificationChallenge(payload), "challenge-value");
  assertThrows(
    () => parseLarkPayload(JSON.stringify({ encrypt: "ciphertext" })),
    Error,
    "Encrypted Lark events",
  );
});
