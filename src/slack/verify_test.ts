import { assert, assertEquals } from "@std/assert";
import { computeSlackSignatureV0, verifySlackRequest } from "./verify.ts";

Deno.test("verifySlackRequest accepts valid signature", async () => {
  const secret = "test_secret";
  const body = "a=1&b=2";
  const timestamp = `${Math.floor(Date.now() / 1000)}`;
  const sig = await computeSlackSignatureV0(secret, timestamp, body);

  const req = new Request("http://localhost/slack/commands", {
    method: "POST",
    headers: {
      "x-slack-request-timestamp": timestamp,
      "x-slack-signature": sig,
    },
    body,
  });

  const res = await verifySlackRequest(req, body, secret, {
    nowMs: Number(timestamp) * 1000,
  });
  assert(res.ok);
});

Deno.test("verifySlackRequest rejects invalid signature", async () => {
  const req = new Request("http://localhost/slack/commands", {
    method: "POST",
    headers: {
      "x-slack-request-timestamp": "1700000000",
      "x-slack-signature": "v0=badbad",
    },
    body: "x=1",
  });

  const res = await verifySlackRequest(req, "x=1", "secret", {
    nowMs: 1700000000 * 1000,
  });
  assertEquals(res.ok, false);
});
