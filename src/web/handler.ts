import { getConfigAsync } from "../config.ts";
import { notificationBodyText } from "../mail/message.ts";
import {
  getMailboxMessageForWeb,
  listAllMailboxBundlesForWeb,
  listMailboxMessagesForWeb,
} from "../mail/service.ts";
import {
  buildClearWebAdminSessionCookie,
  buildWebAdminSessionCookie,
  isWebAdminAuthenticated,
  isWebConsoleEnabled,
  verifyWebAdminPassword,
} from "./auth.ts";
import {
  buildWebConsoleState,
  toWebMailboxSummary,
  toWebMessageDetail,
  toWebMessageSummary,
} from "./service.ts";
import { renderAppPage, renderLoginPage } from "./ui.ts";

function redirect(location: string, headers?: HeadersInit): Response {
  return new Response(null, {
    status: 302,
    headers: {
      location,
      ...(headers ?? {}),
    },
  });
}

function jsonResponse(data: unknown, status = 200): Response {
  return new Response(JSON.stringify(data), {
    status,
    headers: { "content-type": "application/json; charset=utf-8" },
  });
}

function apiUnauthorizedResponse(message: string, status = 401): Response {
  return jsonResponse({ error: message }, status);
}

function matchPath(pathname: string, pattern: RegExp): RegExpMatchArray | null {
  return pathname.match(pattern);
}

export async function handleWebRequest(request: Request): Promise<Response | null> {
  const url = new URL(request.url);
  if (!url.pathname.startsWith("/app") && !url.pathname.startsWith("/api/")) {
    return null;
  }

  if (url.pathname === "/app/") {
    return redirect("/app");
  }
  if (url.pathname === "/app/login/") {
    return redirect("/app/login");
  }

  const config = await getConfigAsync();

  if (url.pathname === "/app/login" && request.method === "GET") {
    if (isWebConsoleEnabled(config) && await isWebAdminAuthenticated(request, config)) {
      return redirect("/app");
    }
    return renderLoginPage({ configured: isWebConsoleEnabled(config) });
  }

  if (url.pathname === "/app/login" && request.method === "POST") {
    if (!isWebConsoleEnabled(config)) {
      return renderLoginPage({
        configured: false,
        error: "当前未配置 WEB_ADMIN_PASSWORD。",
      });
    }
    const form = await request.formData();
    const password = String(form.get("password") ?? "");
    const ok = await verifyWebAdminPassword(password, config);
    if (!ok) {
      return renderLoginPage({
        configured: true,
        error: "管理员密码错误。",
      });
    }
    return redirect("/app", {
      "set-cookie": await buildWebAdminSessionCookie(config),
    });
  }

  if (url.pathname === "/app/logout" && request.method === "POST") {
    return redirect("/app/login", {
      "set-cookie": buildClearWebAdminSessionCookie(config),
    });
  }

  if (!isWebConsoleEnabled(config)) {
    return url.pathname.startsWith("/api/")
      ? apiUnauthorizedResponse("WEB_ADMIN_PASSWORD 未配置，Web 控制台不可用。", 503)
      : renderLoginPage({ configured: false });
  }

  const authenticated = await isWebAdminAuthenticated(request, config);
  if (!authenticated) {
    return url.pathname.startsWith("/api/")
      ? apiUnauthorizedResponse("未登录或会话已过期。")
      : redirect("/app/login");
  }

  if (url.pathname === "/app" && request.method === "GET") {
    const state = await buildWebConsoleState({
      mailboxId: url.searchParams.get("mailbox"),
      folder: url.searchParams.get("folder"),
      messageId: url.searchParams.get("message"),
    });
    return renderAppPage(state);
  }

  if (url.pathname === "/api/mailboxes" && request.method === "GET") {
    const mailboxes = await listAllMailboxBundlesForWeb();
    return jsonResponse({
      mailboxes: mailboxes.map((mailbox) => toWebMailboxSummary(mailbox)),
      count: mailboxes.length,
    });
  }

  const messageListMatch = matchPath(url.pathname, /^\/api\/mailboxes\/([^/]+)\/messages$/);
  if (messageListMatch && request.method === "GET") {
    try {
      const mailboxId = decodeURIComponent(messageListMatch[1]);
      const page = await listMailboxMessagesForWeb({
        mailboxId,
        folderKind: url.searchParams.get("folder") ?? "inbox",
      });
      return jsonResponse({
        mailbox: toWebMailboxSummary(page.bundle),
        folder: page.folder,
        messages: page.messages.map((message) => toWebMessageSummary(message)),
        nextPageUrl: page.nextPageUrl,
      });
    } catch (error) {
      return jsonResponse(
        { error: error instanceof Error ? error.message : String(error) },
        400,
      );
    }
  }

  const detailMatch = matchPath(
    url.pathname,
    /^\/api\/mailboxes\/([^/]+)\/messages\/([^/]+)$/,
  );
  if (detailMatch && request.method === "GET") {
    try {
      const detail = await getMailboxMessageForWeb({
        mailboxId: decodeURIComponent(detailMatch[1]),
        messageId: decodeURIComponent(detailMatch[2]),
        folderKind: url.searchParams.get("folder") ?? "inbox",
      });
      return jsonResponse({
        mailbox: toWebMailboxSummary(detail.bundle),
        message: toWebMessageDetail({
          message: detail.message,
          bodyPlainText: notificationBodyText(detail.message),
          bodyHtml: detail.message.bodyContentType === "html" ? detail.message.bodyText : undefined,
        }),
      });
    } catch (error) {
      return jsonResponse(
        { error: error instanceof Error ? error.message : String(error) },
        400,
      );
    }
  }

  return new Response("Not found", { status: 404 });
}
