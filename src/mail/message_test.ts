import { assertEquals } from "@std/assert";
import { buildDedupeKey, toPreviewText } from "./message.ts";

Deno.test("buildDedupeKey prefers internetMessageId", () => {
  assertEquals(
    buildDedupeKey("mailbox-1", {
      messageId: "graph-id",
      internetMessageId: "internet-id",
      subject: "Hello",
    }),
    "internet-id",
  );
});

Deno.test("buildDedupeKey falls back to mailbox/message id", () => {
  assertEquals(
    buildDedupeKey("mailbox-1", {
      messageId: "graph-id",
      subject: "Hello",
    }),
    "mailbox-1:graph-id",
  );
});

Deno.test("toPreviewText normalizes whitespace and truncates", () => {
  assertEquals(toPreviewText("hello\n\nworld", 10), "hello wor…");
});
