import { assertEquals } from "@std/assert";
import { parseMailCommand } from "./parse.ts";

Deno.test("parseMailCommand returns help for empty input", () => {
  assertEquals(parseMailCommand("").kind, "help");
});

Deno.test("parseMailCommand parses route with slack channel mention", () => {
  const cmd = parseMailCommand("route mailbox@example.com <#C123456|alerts>");
  assertEquals(cmd.kind, "route");
  if (cmd.kind === "route") {
    assertEquals(cmd.mailbox, "mailbox@example.com");
    assertEquals(cmd.channelId, "C123456");
    assertEquals(cmd.channelName, "alerts");
  }
});

Deno.test("parseMailCommand parses test action", () => {
  const cmd = parseMailCommand("test mailbox@example.com");
  assertEquals(cmd.kind, "test");
  if (cmd.kind === "test") {
    assertEquals(cmd.mailbox, "mailbox@example.com");
  }
});

Deno.test("parseMailCommand parses connect provider option", () => {
  const cmd = parseMailCommand("connect msoauth2api");
  assertEquals(cmd.kind, "connect");
  if (cmd.kind === "connect") {
    assertEquals(cmd.providerType, "ms_oauth2api");
  }
});

Deno.test("parseMailCommand parses provider switch command", () => {
  const cmd = parseMailCommand("provider mailbox@example.com graph");
  assertEquals(cmd.kind, "provider");
  if (cmd.kind === "provider") {
    assertEquals(cmd.mailbox, "mailbox@example.com");
    assertEquals(cmd.providerType, "graph_native");
  }
});
