import { assertEquals } from "@std/assert";
import { handleRequest } from "./handler.ts";

Deno.test("migration endpoint requires POST", async () => {
  const response = await handleRequest(
    new Request("https://mail.example.com/admin/migrate-deno-kv"),
  );
  assertEquals(response.status, 405);
  assertEquals(response.headers.get("allow"), "POST");
});
