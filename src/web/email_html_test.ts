import { assertEquals, assertStringIncludes } from "@std/assert";
import { buildReaderDocumentHtml, sanitizeEmailHtml } from "./email_html.ts";

Deno.test("sanitizeEmailHtml removes active content and resolves cid images", () => {
  const result = sanitizeEmailHtml(
    '<script>alert(1)</script><img src="cid:logo"><a href="javascript:alert(1)">open</a>',
    [{
      attachmentId: "a1",
      name: "logo.png",
      contentType: "image/png",
      contentId: "logo",
      dataBase64: "YWJj",
    }],
  );

  assertEquals(result.includes("<script"), false);
  assertStringIncludes(result, "data:image/png;base64,YWJj");
  assertStringIncludes(result, 'href="#"');
});

Deno.test("buildReaderDocumentHtml applies a restrictive document policy", () => {
  const result = buildReaderDocumentHtml("<p>Hello</p>", []);

  assertStringIncludes(result, "Content-Security-Policy");
  assertStringIncludes(result, "default-src 'none'");
  assertStringIncludes(result, "<p>Hello</p>");
});
