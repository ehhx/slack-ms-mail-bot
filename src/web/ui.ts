import { formatFolderLabel, monitoredFoldersText } from "../mail/message.ts";
import type { MailInlineImage, MailboxBundle } from "../mail/types.ts";
import type { WebConsoleState, WebMessageDetail } from "./service.ts";

function escapeHtml(input: string): string {
  return input
    .replace(/&/g, "&amp;")
    .replace(/</g, "&lt;")
    .replace(/>/g, "&gt;")
    .replace(/"/g, "&quot;")
    .replace(/'/g, "&#39;");
}

function fmtTime(iso: string | undefined): string {
  if (!iso) return "-";
  const parsed = new Date(iso);
  if (Number.isNaN(parsed.getTime())) return iso;
  return parsed.toLocaleString("zh-CN", { hour12: false });
}

function appHref(input: {
  mailboxId?: string;
  folder?: "inbox" | "junk";
  messageId?: string;
}): string {
  const params = new URLSearchParams();
  if (input.mailboxId) params.set("mailbox", input.mailboxId);
  if (input.folder) params.set("folder", input.folder);
  if (input.messageId) params.set("message", input.messageId);
  const query = params.toString();
  return query ? `/app?${query}` : "/app";
}

function providerLabel(bundle: MailboxBundle): string {
  return bundle.connection.providerType === "ms_oauth2api" ? "msOauth2api" : "Graph Native";
}

function normalizeContentId(input: string | undefined): string {
  return (input ?? "")
    .trim()
    .replace(/^cid:/i, "")
    .replace(/^<|>$/g, "")
    .toLowerCase();
}

function dataUrlForInlineImage(image: MailInlineImage): string {
  return `data:${image.contentType};base64,${image.dataBase64}`;
}

function rewriteCidImages(html: string, inlineImages: MailInlineImage[] | undefined): string {
  const contentIdMap = new Map<string, string>();
  for (const image of inlineImages ?? []) {
    const contentId = normalizeContentId(image.contentId);
    if (contentId) {
      contentIdMap.set(contentId, dataUrlForInlineImage(image));
    }
    contentIdMap.set(normalizeContentId(image.name), dataUrlForInlineImage(image));
  }

  return html.replace(
    /(<img\b[^>]*\bsrc\s*=\s*)(["'])(cid:[^"']+)\2/gi,
    (_full, prefix, quote, src) => {
      const cid = normalizeContentId(String(src));
      const resolved = contentIdMap.get(cid);
      if (!resolved) return `${prefix}${quote}${src}${quote}`;
      return `${prefix}${quote}${resolved}${quote}`;
    },
  );
}

function sanitizeEmailHtml(html: string, inlineImages: MailInlineImage[] | undefined): string {
  return rewriteCidImages(html, inlineImages)
    .replace(/<script[\s\S]*?<\/script>/gi, "")
    .replace(/<iframe[\s\S]*?<\/iframe>/gi, "")
    .replace(/<object[\s\S]*?<\/object>/gi, "")
    .replace(/<embed[\s\S]*?>/gi, "")
    .replace(/<form[\s\S]*?<\/form>/gi, "")
    .replace(/\son\w+\s*=\s*(".*?"|'.*?'|[^\s>]+)/gi, "")
    .replace(/\s(href|src)\s*=\s*(['"])\s*javascript:[^'"]*\2/gi, ' $1="#"')
    .replace(/<a\b/gi, '<a target="_blank" rel="noopener noreferrer"');
}

function renderMessageBody(detail: WebMessageDetail | null): string {
  if (!detail) {
    return `
      <div class="empty-state">
        <h3>未选择邮件</h3>
        <p>从中间列表选择一封邮件查看详情。</p>
      </div>
    `;
  }

  const attachments = detail.message.attachments ?? [];
  const htmlBody = detail.bodyHtml?.trim();
  const bodyBlock = htmlBody
    ? `
      <iframe
        class="mail-body-frame"
        sandbox="allow-popups allow-popups-to-escape-sandbox"
        srcdoc="${escapeHtml(sanitizeEmailHtml(htmlBody, detail.message.inlineImages))}"
      ></iframe>
    `
    : `<pre class="mail-body-text">${escapeHtml(detail.bodyPlainText || "(无可用正文)")}</pre>`;

  return `
    <div class="detail-header">
      <div>
        <div class="detail-subject">${escapeHtml(detail.message.subject || "(no subject)")}</div>
        <div class="detail-meta-line">
          <span>${escapeHtml(detail.message.fromName || detail.message.fromAddress || "Unknown sender")}</span>
          <span>${escapeHtml(detail.message.fromAddress || "")}</span>
        </div>
      </div>
      ${detail.message.webLink
        ? `<a class="button-link" href="${escapeHtml(detail.message.webLink)}" target="_blank" rel="noopener noreferrer">Open in Outlook</a>`
        : ""}
    </div>
    <div class="detail-grid">
      <div><span class="label">时间</span><span>${escapeHtml(fmtTime(detail.message.receivedDateTime))}</span></div>
      <div><span class="label">文件夹</span><span>${escapeHtml(formatFolderLabel(detail.message.folderKind, detail.message.folderName))}</span></div>
      <div><span class="label">附件</span><span>${attachments.length}</span></div>
      <div><span class="label">内联图片</span><span>${detail.message.inlineImages?.length ?? 0}</span></div>
    </div>
    ${attachments.length > 0
      ? `
      <div class="attachments">
        <div class="section-title">附件</div>
        <ul>
          ${attachments.map((attachment) =>
            `<li>${escapeHtml(attachment.name)}${attachment.contentType ? ` · ${escapeHtml(attachment.contentType)}` : ""}${attachment.size ? ` · ${Math.round(attachment.size / 1024)} KB` : ""}</li>`
          ).join("")}
        </ul>
      </div>
    `
      : ""}
    <div class="section-title">正文</div>
    ${bodyBlock}
  `;
}

function renderMailboxItem(
  mailbox: MailboxBundle,
  selectedMailboxId: string | undefined,
  selectedFolder: "inbox" | "junk",
): string {
  const active = mailbox.connection.mailboxId === selectedMailboxId;
  return `
    <a class="mailbox-item${active ? " is-active" : ""}" href="${appHref({
      mailboxId: mailbox.connection.mailboxId,
      folder: selectedFolder,
    })}">
      <div class="mailbox-title-row">
        <span class="mailbox-title">${escapeHtml(mailbox.connection.displayName || mailbox.connection.emailAddress)}</span>
        <span class="badge">${escapeHtml(providerLabel(mailbox))}</span>
      </div>
      <div class="mailbox-subtitle">${escapeHtml(mailbox.connection.emailAddress)}</div>
      <div class="mailbox-meta">
        <span>${escapeHtml(mailbox.connection.status)}</span>
        <span>${escapeHtml(monitoredFoldersText(mailbox))}</span>
      </div>
      <div class="mailbox-meta">
        <span>Slack</span>
        <span>${escapeHtml(mailbox.route?.slackChannelName || mailbox.route?.slackChannelId || "未配置")}</span>
      </div>
    </a>
  `;
}

function renderMessageItem(
  state: WebConsoleState,
  message: { messageId: string; subject: string; fromName?: string; fromAddress?: string; bodyPreview?: string; receivedDateTime?: string; hasAttachments?: boolean },
): string {
  const selectedMessageId = state.selectedMessage?.message.messageId;
  const active = message.messageId === selectedMessageId;
  return `
    <a class="message-item${active ? " is-active" : ""}" href="${appHref({
      mailboxId: state.selectedMailbox?.connection.mailboxId,
      folder: state.selectedFolder,
      messageId: message.messageId,
    })}">
      <div class="message-topline">
        <span class="message-subject">${escapeHtml(message.subject || "(no subject)")}</span>
        <span class="message-time">${escapeHtml(fmtTime(message.receivedDateTime))}</span>
      </div>
      <div class="message-sender">${escapeHtml(message.fromName || message.fromAddress || "Unknown sender")}</div>
      <div class="message-preview">${escapeHtml(message.bodyPreview || "(无预览)")}</div>
      <div class="message-flags">${message.hasAttachments ? "📎 含附件" : "&nbsp;"}</div>
    </a>
  `;
}

function renderEmptyMailboxes(): string {
  return `
    <div class="empty-state">
      <h3>还没有连接邮箱</h3>
      <p>先在 Slack 里运行 <code>/mail connect graph</code>，把 Outlook 账号接入后，这里会自动显示多账号列表。</p>
    </div>
  `;
}

function renderAppShell(title: string, body: string): Response {
  return new Response(
    `<!doctype html>
<html lang="zh-CN">
  <head>
    <meta charset="utf-8" />
    <meta name="viewport" content="width=device-width, initial-scale=1" />
    <title>${escapeHtml(title)}</title>
    <style>
      :root {
        color-scheme: dark;
        --bg: #0b1020;
        --bg-soft: #11182c;
        --bg-panel: #0f1728;
        --border: rgba(255,255,255,.08);
        --text: #e7edf7;
        --muted: #91a0bd;
        --accent: #6ea8ff;
        --accent-soft: rgba(110,168,255,.14);
        --danger: #ff8d8d;
      }
      * { box-sizing: border-box; }
      html, body { margin: 0; min-height: 100%; background: linear-gradient(180deg, #0b1020 0%, #0d1426 100%); color: var(--text); font-family: Inter, ui-sans-serif, system-ui, -apple-system, BlinkMacSystemFont, sans-serif; }
      a { color: inherit; text-decoration: none; }
      code { font-family: ui-monospace, SFMono-Regular, Menlo, monospace; }
      .page { min-height: 100vh; display: grid; grid-template-rows: 64px 1fr; }
      .topbar {
        display: flex; align-items: center; justify-content: space-between;
        padding: 0 20px; border-bottom: 1px solid var(--border); backdrop-filter: blur(10px);
        background: rgba(11,16,32,.72); position: sticky; top: 0; z-index: 10;
      }
      .brand { display: flex; flex-direction: column; gap: 2px; }
      .brand-title { font-size: 16px; font-weight: 700; letter-spacing: .02em; }
      .brand-subtitle { font-size: 12px; color: var(--muted); }
      .topbar-actions { display: flex; align-items: center; gap: 10px; }
      .shell {
        min-height: calc(100vh - 64px);
        display: grid;
        grid-template-columns: 280px 360px minmax(420px, 1fr);
      }
      .pane { min-height: calc(100vh - 64px); overflow: auto; }
      .pane + .pane { border-left: 1px solid var(--border); }
      .sidebar, .list-pane, .detail-pane { background: rgba(10,15,28,.72); }
      .sidebar { padding: 18px 14px 28px; }
      .list-pane { padding: 18px 0 28px; }
      .detail-pane { padding: 22px 28px 36px; }
      .section-heading { padding: 0 14px 12px; font-size: 12px; font-weight: 700; letter-spacing: .08em; text-transform: uppercase; color: var(--muted); }
      .mailbox-item, .message-item {
        display: block; border: 1px solid transparent; border-radius: 14px;
        background: transparent; transition: background .15s ease, border-color .15s ease;
      }
      .mailbox-item { padding: 14px; margin-bottom: 8px; }
      .message-item { padding: 14px 18px; border-radius: 0; border-top: 1px solid var(--border); }
      .mailbox-item:hover, .message-item:hover { background: rgba(255,255,255,.03); }
      .mailbox-item.is-active, .message-item.is-active { background: var(--accent-soft); border-color: rgba(110,168,255,.22); }
      .mailbox-title-row, .message-topline, .detail-header {
        display: flex; align-items: flex-start; justify-content: space-between; gap: 12px;
      }
      .mailbox-title, .message-subject, .detail-subject { font-weight: 700; }
      .mailbox-subtitle, .message-sender, .message-preview, .mailbox-meta, .detail-meta-line { color: var(--muted); }
      .mailbox-subtitle, .message-sender, .message-preview, .mailbox-meta { font-size: 13px; line-height: 1.45; }
      .mailbox-meta { display: flex; justify-content: space-between; gap: 12px; margin-top: 8px; }
      .badge {
        display: inline-flex; align-items: center; padding: 2px 8px; border-radius: 999px;
        border: 1px solid rgba(255,255,255,.08); color: var(--muted); font-size: 11px; white-space: nowrap;
      }
      .filters { display: flex; gap: 8px; padding: 0 14px 14px; }
      .filter-tab {
        display: inline-flex; align-items: center; padding: 8px 12px; border-radius: 999px;
        border: 1px solid var(--border); color: var(--muted); font-size: 13px;
      }
      .filter-tab.is-active { color: var(--text); border-color: rgba(110,168,255,.26); background: var(--accent-soft); }
      .alert {
        margin: 0 14px 14px; padding: 12px 14px; border-radius: 12px;
        border: 1px solid rgba(255,141,141,.18); background: rgba(255,141,141,.08); color: #ffd4d4; font-size: 13px;
      }
      .empty-state {
        min-height: 220px; display: grid; place-content: center; gap: 8px; text-align: center;
        color: var(--muted); padding: 24px;
      }
      .empty-state h3 { margin: 0; color: var(--text); font-size: 20px; }
      .button-link, .button-ghost {
        display: inline-flex; align-items: center; justify-content: center; padding: 9px 14px; border-radius: 10px;
        border: 1px solid rgba(110,168,255,.22); background: var(--accent-soft); color: var(--text); font-size: 13px;
      }
      .button-ghost { background: transparent; border-color: var(--border); color: var(--muted); }
      .detail-header { margin-bottom: 18px; }
      .detail-subject { font-size: 24px; line-height: 1.2; margin-bottom: 8px; }
      .detail-meta-line { display: flex; flex-wrap: wrap; gap: 10px; font-size: 14px; }
      .detail-grid {
        display: grid; grid-template-columns: repeat(2, minmax(0, 1fr)); gap: 12px 18px;
        padding: 16px 0 20px; border-top: 1px solid var(--border); border-bottom: 1px solid var(--border); margin-bottom: 20px;
      }
      .detail-grid > div { display: flex; flex-direction: column; gap: 6px; font-size: 14px; }
      .label, .section-title { color: var(--muted); font-size: 12px; text-transform: uppercase; letter-spacing: .08em; }
      .section-title { margin-bottom: 10px; }
      .attachments ul { list-style: none; margin: 0 0 20px; padding: 0; display: grid; gap: 8px; }
      .attachments li {
        padding: 10px 12px; border-radius: 12px; border: 1px solid var(--border); background: rgba(255,255,255,.02);
        font-size: 13px; color: var(--muted);
      }
      .mail-body-frame {
        width: 100%; min-height: 680px; border: 1px solid var(--border); border-radius: 16px; background: #fff;
      }
      .mail-body-text {
        white-space: pre-wrap; word-break: break-word; line-height: 1.65; margin: 0;
        padding: 20px; border-radius: 16px; border: 1px solid var(--border); background: rgba(255,255,255,.02);
        color: var(--text);
      }
      .login-wrap {
        min-height: 100vh; display: grid; place-items: center; padding: 24px;
      }
      .login-panel {
        width: min(420px, 100%); padding: 28px; border-radius: 20px; border: 1px solid var(--border);
        background: rgba(15,23,40,.86); box-shadow: 0 24px 80px rgba(0,0,0,.35);
      }
      .login-panel h1 { margin: 0 0 10px; font-size: 28px; }
      .login-panel p { margin: 0 0 18px; color: var(--muted); line-height: 1.6; }
      .login-panel form { display: grid; gap: 12px; }
      .input {
        width: 100%; padding: 12px 14px; border-radius: 12px; border: 1px solid var(--border);
        background: rgba(255,255,255,.03); color: var(--text); outline: none;
      }
      .footer-note { padding: 18px 14px 0; color: var(--muted); font-size: 12px; line-height: 1.6; }
      @media (max-width: 1120px) {
        .shell { grid-template-columns: 250px 320px minmax(320px, 1fr); }
      }
      @media (max-width: 920px) {
        .shell { grid-template-columns: 1fr; }
        .pane { min-height: auto; }
        .pane + .pane { border-left: 0; border-top: 1px solid var(--border); }
        .detail-pane { padding: 20px 18px 28px; }
        .mail-body-frame { min-height: 420px; }
      }
    </style>
  </head>
  <body>
    ${body}
  </body>
</html>`,
    { status: 200, headers: { "content-type": "text/html; charset=utf-8" } },
  );
}

export function renderLoginPage(input: {
  error?: string;
  configured: boolean;
}): Response {
  return renderAppShell(
    "Mail Console Login",
    `
      <div class="login-wrap">
        <div class="login-panel">
          <h1>Mail Console</h1>
          <p>只读多账号 Outlook 控制台。使用管理员密码登录后，可以查看各邮箱的 Inbox / Junk 最近邮件和详情。</p>
          ${
            input.configured
              ? `
                ${input.error ? `<div class="alert" style="margin:0 0 14px;">${escapeHtml(input.error)}</div>` : ""}
                <form method="POST" action="/app/login">
                  <input class="input" type="password" name="password" placeholder="输入管理员密码" autocomplete="current-password" required />
                  <button class="button-link" type="submit">登录</button>
                </form>
              `
              : `
                <div class="alert" style="margin:16px 0 0;">当前未配置 <code>WEB_ADMIN_PASSWORD</code>，Web 控制台尚未启用。</div>
              `
          }
        </div>
      </div>
    `,
  );
}

export function renderAppPage(state: WebConsoleState): Response {
  const selectedMailboxId = state.selectedMailbox?.connection.mailboxId;
  const selectedFolder = state.selectedFolder;
  const mailboxCount = state.mailboxes.length;
  return renderAppShell(
    "Mail Console",
    `
      <div class="page">
        <header class="topbar">
          <div class="brand">
            <div class="brand-title">Mail Console</div>
            <div class="brand-subtitle">多账号 Outlook 只读控制台 · ${mailboxCount} 个邮箱</div>
          </div>
          <div class="topbar-actions">
            <a class="button-ghost" href="/healthz" target="_blank" rel="noopener noreferrer">healthz</a>
            <form method="POST" action="/app/logout">
              <button class="button-ghost" type="submit">退出</button>
            </form>
          </div>
        </header>
        <div class="shell">
          <aside class="pane sidebar">
            <div class="section-heading">Mailboxes</div>
            ${
              state.mailboxes.length > 0
                ? state.mailboxes.map((mailbox) =>
                  renderMailboxItem(mailbox, selectedMailboxId, selectedFolder)
                ).join("")
                : renderEmptyMailboxes()
            }
            <div class="footer-note">
              管理动作仍然建议在 Slack 执行：<br />
              <code>/mail connect graph</code><br />
              <code>/mail route &lt;mailbox&gt; &lt;#channel&gt;</code><br />
              <code>/mail provider &lt;mailbox&gt; graph</code>
            </div>
          </aside>
          <section class="pane list-pane">
            <div class="section-heading">Messages</div>
            ${
              state.selectedMailbox
                ? `
                  <div class="filters">
                    <a class="filter-tab${selectedFolder === "inbox" ? " is-active" : ""}" href="${appHref({
                      mailboxId: state.selectedMailbox.connection.mailboxId,
                      folder: "inbox",
                    })}">Inbox</a>
                    <a class="filter-tab${selectedFolder === "junk" ? " is-active" : ""}" href="${appHref({
                      mailboxId: state.selectedMailbox.connection.mailboxId,
                      folder: "junk",
                    })}">Junk</a>
                  </div>
                `
                : ""
            }
            ${state.error ? `<div class="alert">${escapeHtml(state.error)}</div>` : ""}
            ${
              state.messages.length > 0
                ? state.messages.map((message) => renderMessageItem(state, message)).join("")
                : `
                  <div class="empty-state">
                    <h3>这个文件夹暂无可展示邮件</h3>
                    <p>当前只展示最近一页邮件。若邮箱刚接入，先等同步完成或在 Slack 中执行 <code>/mail sync &lt;mailbox&gt;</code>。</p>
                  </div>
                `
            }
            ${
              state.nextPageUrl
                ? `<div class="footer-note">当前先展示最近一页邮件，后续可继续扩展分页 / 搜索。</div>`
                : ""
            }
          </section>
          <main class="pane detail-pane">
            ${renderMessageBody(state.selectedMessage)}
          </main>
        </div>
      </div>
    `,
  );
}
