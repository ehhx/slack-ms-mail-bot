import { assertEquals, assertMatch, assertStringIncludes } from "@std/assert";
import { WEB_APP_CSS, WEB_APP_JS } from "./assets.ts";
import { renderAppPage, renderLoginPage } from "./ui.ts";

async function responseText(response: Response): Promise<string> {
  assertEquals(response.status, 200);
  assertStringIncludes(
    response.headers.get("content-type") ?? "",
    "text/html",
  );
  return await response.text();
}

Deno.test("renderLoginPage exposes a focused administrator login flow", async () => {
  const html = await responseText(
    renderLoginPage({ configured: true, error: "密码 <错误>" }),
  );

  assertStringIncludes(html, 'class="login-page"');
  assertStringIncludes(html, 'autocomplete="current-password"');
  assertStringIncludes(html, 'action="/app/login"');
  assertStringIncludes(html, "密码 &lt;错误&gt;");
  assertEquals(html.includes("密码 <错误>"), false);
});

Deno.test("renderLoginPage explains when the web console is disabled", async () => {
  const html = await responseText(renderLoginPage({ configured: false }));

  assertStringIncludes(html, "WEB_ADMIN_PASSWORD");
  assertEquals(html.includes('action="/app/login"'), false);
});

Deno.test("renderAppPage includes the responsive workspace landmarks", async () => {
  const html = await responseText(renderAppPage());

  for (
    const marker of [
      'id="accountTrigger"',
      'id="mailSearch"',
      'id="messageList"',
      'id="streamNoticeText"',
      'id="emptyAction"',
      'id="readerPane"',
      'data-mobile-view="list"',
      'aria-label="邮件文件夹"',
    ]
  ) {
    assertStringIncludes(html, marker);
  }
  assertMatch(html, /app\.js\?v=\d{8}-\d/);
});

Deno.test("web assets retain responsive, reduced-motion, and valid script coverage", () => {
  assertStringIncludes(WEB_APP_CSS, "@media (max-width: 760px)");
  assertStringIncludes(WEB_APP_CSS, "@media (prefers-reduced-motion: reduce)");
  assertStringIncludes(WEB_APP_JS, "elements.noticeRetry.addEventListener");
  assertStringIncludes(WEB_APP_JS, "elements.streamPane.inert");
  assertStringIncludes(WEB_APP_JS, 'setAttribute("aria-hidden"');
  assertStringIncludes(WEB_APP_JS, "new AbortController()");
  new Function(WEB_APP_JS);
});
