import { assertEquals } from "@std/assert";
import { getGraphValidationToken, isGraphClientStateValid, parseGraphWebhookBody } from "./webhook.ts";

Deno.test("getGraphValidationToken reads validation query", () => {
  const req = new Request("https://example.com/graph/webhook?validationToken=abc123");
  assertEquals(getGraphValidationToken(req), "abc123");
});

Deno.test("parseGraphWebhookBody returns notifications array", () => {
  const body = JSON.stringify({ value: [{ subscriptionId: "sub-1", clientState: "state" }] });
  const parsed = parseGraphWebhookBody(body);
  assertEquals(parsed.value.length, 1);
  assertEquals(parsed.value[0].subscriptionId, "sub-1");
});

Deno.test("isGraphClientStateValid checks exact match", () => {
  assertEquals(
    isGraphClientStateValid({ subscriptionId: "sub-1", clientState: "expected" }, "expected"),
    true,
  );
  assertEquals(
    isGraphClientStateValid({ subscriptionId: "sub-1", clientState: "wrong" }, "expected"),
    false,
  );
});
