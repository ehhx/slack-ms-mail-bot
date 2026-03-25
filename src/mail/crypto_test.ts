import { assertEquals } from "@std/assert";
import { decryptSecret, encryptSecret } from "./crypto.ts";

Deno.test("encryptSecret roundtrips", async () => {
  const secret = await encryptSecret("refresh-token", "super-secret");
  const plain = await decryptSecret(secret, "super-secret");
  assertEquals(plain, "refresh-token");
});
