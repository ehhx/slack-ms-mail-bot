import { formatFolderLabel, monitoredFoldersText } from "../mail/message.ts";
import type { MailInlineImage, MailboxBundle } from "../mail/types.ts";
import type { WebConsoleState, WebMessageDetail } from "./service.ts";

function escapeHtml(input: string): string {
  return input
    .replace(/&/g, "&amp;")
    .replace(/</g, "&lt;")
    .replace(/>/g, "&gt;")
    .replace(/\"/g, "&quot;")
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
    const dataUrl = dataUrlForInlineImage(image);
    if (contentId) {
      contentIdMap.set(contentId, dataUrl);
    }
    contentIdMap.set(normalizeContentId(image.name), dataUrl);
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
    .replace(/<!doctype[^>]*>/gi, "")
    .replace(/<script[\s\S]*?<\/script>/gi, "")
    .replace(/<iframe[\s\S]*?<\/iframe>/gi, "")
    .replace(/<object[\s\S]*?<\/object>/gi, "")
    .replace(/<embed[\s\S]*?>/gi, "")
    .replace(/<form[\s\S]*?<\/form>/gi, "")
    .replace(/<base[\s\S]*?>/gi, "")
    .replace(/<meta[\s\S]*?>/gi, "")
    .replace(/<link[\s\S]*?>/gi, "")
    .replace(/<\/?(html|body|head)[^>]*>/gi, "")
    .replace(/\son\w+\s*=\s*(".*?"|'.*?'|[^\s>]+)/gi, "")
    .replace(/\s(href|src)\s*=\s*(['"])\s*javascript:[^'"]*\2/gi, ' $1="#"')
    .replace(/<a\b/gi, '<a target="_blank" rel="noopener noreferrer"');
}

function buildReaderSrcdoc(html: string, inlineImages: MailInlineImage[] | undefined): string {
  const sanitized = sanitizeEmailHtml(html, inlineImages);
  return `<!doctype html>
<html lang="zh-CN">
  <head>
    <meta charset="utf-8" />
    <meta name="viewport" content="width=device-width, initial-scale=1" />
    <style>
      :root {
        color-scheme: light;
        --text: #0f172a;
        --muted: #475569;
        --border: #dbe4f0;
        --accent: #2563eb;
        --bg: #ffffff;
      }
      * { box-sizing: border-box; }
      html, body { margin: 0; padding: 0; background: var(--bg); color: var(--text); font: 15px/1.65 Inter, ui-sans-serif, system-ui, -apple-system, BlinkMacSystemFont, sans-serif; }
      body { padding: 24px; }
      img { max-width: 100%; height: auto; }
      table { max-width: 100% !important; }
      pre, code { white-space: pre-wrap; word-break: break-word; }
      a { color: var(--accent); }
      blockquote { margin: 1.2rem 0; padding-left: 1rem; border-left: 3px solid var(--border); color: var(--muted); }
      .mail-document { min-height: calc(100vh - 48px); }
    </style>
  </head>
  <body>
    <main class="mail-document">${sanitized}</main>
  </body>
</html>`;
}

function renderReaderEmpty(state: WebConsoleState): string {
  if (!state.selectedMailbox) {
    return `
      <section class="empty-panel">
        <h2>还没有可读邮箱</h2>
        <p>先在 Slack 里运行 <code>/mail connect graph</code>，连接至少一个 Outlook 账号。</p>
      </section>
    `;
  }

  return `
    <section class="reader-placeholder">
      <div class="reader-placeholder-header">
        <div>
          <div class="eyebrow">Reading pane</div>
          <h2>${escapeHtml(state.selectedMailbox.connection.displayName || state.selectedMailbox.connection.emailAddress)}</h2>
          <p>${escapeHtml(state.selectedMailbox.connection.emailAddress)}</p>
        </div>
        <span class="pill">${escapeHtml(providerLabel(state.selectedMailbox))}</span>
      </div>
      <div class="reader-summary-grid">
        <div><span>当前文件夹</span><strong>${escapeHtml(formatFolderLabel(state.selectedFolder))}</strong></div>
        <div><span>消息数量</span><strong>${state.messages.length}</strong></div>
        <div><span>已监控文件夹</span><strong>${escapeHtml(monitoredFoldersText(state.selectedMailbox))}</strong></div>
        <div><span>Slack 路由</span><strong>${escapeHtml(state.selectedMailbox.route?.slackChannelName || state.selectedMailbox.route?.slackChannelId || "未配置")}</strong></div>
      </div>
      <div class="reader-note">
        <p>为减少等待时间，页面现在只在你真正选择邮件后才加载正文、附件和内联图片。</p>
      </div>
    </section>
  `;
}

function renderMessageBody(detail: WebMessageDetail | null, state: WebConsoleState): string {
  if (!detail) {
    return renderReaderEmpty(state);
  }

  const attachments = detail.message.attachments ?? [];
  const htmlBody = detail.bodyHtml?.trim();
  const bodyBlock = htmlBody
    ? `
      <iframe
        class="mail-body-frame"
        title="Mail content"
        loading="lazy"
        sandbox="allow-popups allow-popups-to-escape-sandbox"
        srcdoc="${escapeHtml(buildReaderSrcdoc(htmlBody, detail.message.inlineImages))}"
      ></iframe>
    `
    : `<pre class="mail-body-text">${escapeHtml(detail.bodyPlainText || "(无可用正文)")}</pre>`;

  return `
    <article class="reader-article">
      <header class="reader-article-header">
        <div class="eyebrow">${escapeHtml(formatFolderLabel(detail.message.folderKind, detail.message.folderName))}</div>
        <div class="reader-title-row">
          <h1>${escapeHtml(detail.message.subject || "(no subject)")}</h1>
          ${detail.message.webLink
            ? `<a class="action-link" href="${escapeHtml(detail.message.webLink)}" target="_blank" rel="noopener noreferrer">Open in Outlook</a>`
            : ""}
        </div>
        <div class="reader-subtitle">
          <span>${escapeHtml(detail.message.fromName || detail.message.fromAddress || "Unknown sender")}</span>
          <span>${escapeHtml(detail.message.fromAddress || "")}</span>
        </div>
      </header>

      <section class="reader-meta-grid">
        <div><span>接收时间</span><strong>${escapeHtml(fmtTime(detail.message.receivedDateTime))}</strong></div>
        <div><span>附件数量</span><strong>${attachments.length}</strong></div>
        <div><span>内联图片</span><strong>${detail.message.inlineImages?.length ?? 0}</strong></div>
        <div><span>正文类型</span><strong>${escapeHtml(detail.message.bodyContentType || "text")}</strong></div>
      </section>

      ${attachments.length > 0
        ? `
          <section class="attachment-section">
            <div class="section-title">Attachments</div>
            <ul class="attachment-list">
              ${attachments.map((attachment) =>
                `<li>${escapeHtml(attachment.name)}${attachment.contentType ? ` <span>· ${escapeHtml(attachment.contentType)}</span>` : ""}${attachment.size ? ` <span>· ${Math.max(1, Math.round(attachment.size / 1024))} KB</span>` : ""}</li>`
              ).join("")}
            </ul>
          </section>
        `
        : ""}

      <section class="body-section">
        <div class="section-title">Body</div>
        ${bodyBlock}
      </section>
    </article>
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
        <span class="pill">${escapeHtml(providerLabel(mailbox))}</span>
      </div>
      <div class="mailbox-subtitle">${escapeHtml(mailbox.connection.emailAddress)}</div>
      <div class="mailbox-meta-row">
        <span>${escapeHtml(mailbox.connection.status)}</span>
        <span>${escapeHtml(mailbox.route?.slackChannelName || mailbox.route?.slackChannelId || "未配置")}</span>
      </div>
    </a>
  `;
}

function renderMessageItem(state: WebConsoleState, message: {
  messageId: string;
  subject: string;
  fromName?: string;
  fromAddress?: string;
  bodyPreview?: string;
  receivedDateTime?: string;
  hasAttachments?: boolean;
}): string {
  const selectedMessageId = state.selectedMessage?.message.messageId;
  const active = message.messageId === selectedMessageId;
  return `
    <a class="message-item${active ? " is-active" : ""}" href="${appHref({
      mailboxId: state.selectedMailbox?.connection.mailboxId,
      folder: state.selectedFolder,
      messageId: message.messageId,
    })}">
      <div class="message-item-top">
        <span class="message-subject">${escapeHtml(message.subject || "(no subject)")}</span>
        <span class="message-time">${escapeHtml(fmtTime(message.receivedDateTime))}</span>
      </div>
      <div class="message-sender">${escapeHtml(message.fromName || message.fromAddress || "Unknown sender")}</div>
      <div class="message-preview">${escapeHtml(message.bodyPreview || "(无预览)")}</div>
      <div class="message-flags">${message.hasAttachments ? "附件" : ""}</div>
    </a>
  `;
}

function renderEmptyMailboxes(): string {
  return `
    <section class="empty-panel compact">
      <h3>暂无邮箱</h3>
      <p>在 Slack 执行 <code>/mail connect graph</code> 后，这里会自动出现邮箱列表。</p>
    </section>
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
        --bg: #0a0f1a;
        --surface: #0f1724;
        --surface-2: #121c2b;
        --surface-3: #172235;
        --line: rgba(148, 163, 184, 0.18);
        --line-strong: rgba(148, 163, 184, 0.26);
        --text: #edf3fb;
        --muted: #95a6be;
        --accent: #67a7ff;
        --accent-soft: rgba(103, 167, 255, 0.12);
        --accent-line: rgba(103, 167, 255, 0.28);
        --danger: #ff8f8f;
        --success: #75d0a2;
      }
      * { box-sizing: border-box; }
      html, body {
        margin: 0;
        min-height: 100%;
        background: radial-gradient(circle at top left, #15233a 0%, var(--bg) 34%), var(--bg);
        color: var(--text);
        font-family: Inter, ui-sans-serif, system-ui, -apple-system, BlinkMacSystemFont, sans-serif;
      }
      body { -webkit-font-smoothing: antialiased; text-rendering: optimizeLegibility; }
      a { color: inherit; text-decoration: none; }
      code {
        font-family: ui-monospace, SFMono-Regular, Menlo, Monaco, monospace;
        font-size: 0.92em;
      }
      .app-shell {
        min-height: 100vh;
        display: grid;
        grid-template-rows: 64px 1fr;
      }
      .topbar {
        display: flex;
        align-items: center;
        justify-content: space-between;
        gap: 16px;
        padding: 0 20px;
        border-bottom: 1px solid var(--line);
        background: rgba(10, 15, 26, 0.96);
      }
      .brand-stack { display: grid; gap: 2px; }
      .brand-title {
        font-size: 17px;
        font-weight: 700;
        letter-spacing: 0.01em;
      }
      .brand-subtitle {
        font-size: 12px;
        color: var(--muted);
      }
      .toolbar-meta { display: flex; align-items: center; gap: 8px; flex-wrap: wrap; }
      .pill {
        display: inline-flex;
        align-items: center;
        gap: 6px;
        padding: 5px 10px;
        border-radius: 999px;
        border: 1px solid var(--line);
        color: var(--muted);
        font-size: 12px;
        white-space: nowrap;
      }
      .workspace {
        min-height: calc(100vh - 64px);
        display: grid;
        grid-template-columns: 276px 390px minmax(420px, 1fr);
      }
      .pane {
        min-height: calc(100vh - 64px);
        overflow: auto;
        background: rgba(15, 23, 36, 0.9);
        contain: content;
        overscroll-behavior: contain;
      }
      .pane + .pane { border-left: 1px solid var(--line); }
      .rail { padding: 16px 14px 24px; }
      .stream { padding: 16px 0 24px; }
      .reader { padding: 22px 28px 30px; background: linear-gradient(180deg, rgba(18, 28, 43, 0.96) 0%, rgba(14, 22, 34, 0.96) 100%); }
      .section-heading {
        padding: 0 14px 12px;
        font-size: 11px;
        font-weight: 700;
        letter-spacing: 0.12em;
        text-transform: uppercase;
        color: var(--muted);
      }
      .mailbox-list,
      .message-list {
        display: grid;
        gap: 4px;
      }
      .mailbox-item,
      .message-item {
        position: relative;
        display: grid;
        gap: 6px;
        transition: background-color 120ms ease, border-color 120ms ease, transform 120ms ease;
        content-visibility: auto;
        contain-intrinsic-size: 120px;
      }
      .mailbox-item {
        padding: 12px 12px 12px 14px;
        border-radius: 14px;
        border: 1px solid transparent;
      }
      .message-item {
        padding: 13px 18px 14px 18px;
        border-top: 1px solid var(--line);
      }
      .mailbox-item::before,
      .message-item::before {
        content: "";
        position: absolute;
        left: 0;
        top: 10px;
        bottom: 10px;
        width: 2px;
        border-radius: 999px;
        background: transparent;
      }
      .mailbox-item:hover,
      .message-item:hover {
        background: rgba(255,255,255,0.02);
      }
      .mailbox-item.is-active,
      .message-item.is-active {
        background: var(--accent-soft);
        border-color: var(--accent-line);
      }
      .mailbox-item.is-active::before,
      .message-item.is-active::before {
        background: var(--accent);
      }
      .mailbox-title-row,
      .message-item-top,
      .reader-title-row,
      .reader-placeholder-header {
        display: flex;
        align-items: flex-start;
        justify-content: space-between;
        gap: 12px;
      }
      .mailbox-title,
      .message-subject,
      .reader-article h1,
      .reader-placeholder h2 {
        font-weight: 700;
      }
      .mailbox-title,
      .message-subject { line-height: 1.35; }
      .mailbox-subtitle,
      .mailbox-meta-row,
      .message-sender,
      .message-preview,
      .reader-subtitle,
      .reader-note,
      .empty-panel p,
      .login-panel p {
        color: var(--muted);
      }
      .mailbox-subtitle,
      .mailbox-meta-row,
      .message-sender,
      .message-preview,
      .message-time,
      .message-flags {
        font-size: 13px;
      }
      .mailbox-subtitle,
      .message-sender,
      .message-preview {
        line-height: 1.45;
      }
      .mailbox-meta-row {
        display: flex;
        justify-content: space-between;
        gap: 12px;
      }
      .stream-header {
        display: grid;
        gap: 12px;
        padding: 0 14px 14px;
        border-bottom: 1px solid var(--line);
      }
      .stream-title-row {
        display: flex;
        align-items: flex-end;
        justify-content: space-between;
        gap: 16px;
      }
      .stream-title {
        display: grid;
        gap: 4px;
      }
      .stream-title h2,
      .reader-placeholder h2,
      .empty-panel h3,
      .login-panel h1,
      .reader-article h1 {
        margin: 0;
      }
      .stream-title p,
      .reader-placeholder-header p {
        margin: 0;
        color: var(--muted);
        font-size: 13px;
      }
      .tab-row { display: flex; gap: 8px; flex-wrap: wrap; }
      .tab {
        display: inline-flex;
        align-items: center;
        justify-content: center;
        min-width: 82px;
        padding: 8px 12px;
        border-radius: 999px;
        border: 1px solid var(--line);
        color: var(--muted);
        font-size: 13px;
      }
      .tab.is-active {
        background: var(--accent-soft);
        border-color: var(--accent-line);
        color: var(--text);
      }
      .message-time { color: var(--muted); white-space: nowrap; }
      .message-preview {
        display: -webkit-box;
        -webkit-box-orient: vertical;
        -webkit-line-clamp: 2;
        overflow: hidden;
      }
      .message-flags {
        min-height: 16px;
        color: var(--success);
      }
      .reader-placeholder,
      .reader-article,
      .empty-panel {
        display: grid;
        gap: 18px;
      }
      .eyebrow {
        font-size: 11px;
        font-weight: 700;
        letter-spacing: 0.12em;
        text-transform: uppercase;
        color: var(--muted);
      }
      .reader-placeholder-header p,
      .reader-subtitle {
        display: flex;
        flex-wrap: wrap;
        gap: 10px;
        font-size: 14px;
      }
      .reader-summary-grid,
      .reader-meta-grid {
        display: grid;
        grid-template-columns: repeat(2, minmax(0, 1fr));
        gap: 12px 16px;
      }
      .reader-summary-grid > div,
      .reader-meta-grid > div {
        padding: 14px 16px;
        border-radius: 14px;
        border: 1px solid var(--line);
        background: rgba(255,255,255,0.02);
        display: grid;
        gap: 6px;
      }
      .reader-summary-grid span,
      .reader-meta-grid span {
        font-size: 12px;
        color: var(--muted);
        text-transform: uppercase;
        letter-spacing: 0.08em;
      }
      .reader-summary-grid strong,
      .reader-meta-grid strong {
        font-size: 14px;
        color: var(--text);
      }
      .reader-note {
        padding: 16px 18px;
        border-radius: 14px;
        border: 1px dashed var(--line-strong);
        background: rgba(255,255,255,0.015);
        line-height: 1.6;
      }
      .reader-article-header {
        display: grid;
        gap: 12px;
      }
      .reader-title-row h1 {
        font-size: 32px;
        line-height: 1.1;
        letter-spacing: -0.02em;
      }
      .reader-subtitle { font-size: 14px; }
      .action-link,
      .button-link,
      .button-ghost {
        display: inline-flex;
        align-items: center;
        justify-content: center;
        padding: 9px 14px;
        border-radius: 12px;
        border: 1px solid var(--line);
        font-size: 13px;
        white-space: nowrap;
      }
      .action-link,
      .button-link {
        background: var(--accent-soft);
        border-color: var(--accent-line);
      }
      .button-ghost { color: var(--muted); background: transparent; }
      .section-title {
        font-size: 11px;
        font-weight: 700;
        letter-spacing: 0.12em;
        text-transform: uppercase;
        color: var(--muted);
      }
      .attachment-section,
      .body-section { display: grid; gap: 12px; }
      .attachment-list {
        list-style: none;
        margin: 0;
        padding: 0;
        display: grid;
        gap: 8px;
      }
      .attachment-list li {
        padding: 12px 14px;
        border-radius: 12px;
        border: 1px solid var(--line);
        background: rgba(255,255,255,0.02);
        color: var(--text);
        font-size: 13px;
      }
      .attachment-list li span { color: var(--muted); }
      .mail-body-frame {
        width: 100%;
        min-height: 680px;
        border: 1px solid var(--line-strong);
        border-radius: 18px;
        background: #ffffff;
      }
      .mail-body-text {
        margin: 0;
        padding: 22px;
        white-space: pre-wrap;
        word-break: break-word;
        line-height: 1.7;
        color: var(--text);
        border-radius: 18px;
        border: 1px solid var(--line);
        background: rgba(255,255,255,0.02);
      }
      .login-page {
        min-height: 100vh;
        display: grid;
        place-items: center;
        padding: 24px;
      }
      .login-panel {
        width: min(420px, 100%);
        display: grid;
        gap: 16px;
        padding: 28px;
        border-radius: 22px;
        border: 1px solid var(--line);
        background: rgba(15, 23, 36, 0.92);
      }
      .login-panel form { display: grid; gap: 12px; }
      .input {
        width: 100%;
        padding: 12px 14px;
        border-radius: 12px;
        border: 1px solid var(--line);
        background: rgba(255,255,255,0.025);
        color: var(--text);
        outline: none;
      }
      .footer-note {
        padding: 18px 14px 0;
        color: var(--muted);
        font-size: 12px;
        line-height: 1.65;
      }
      .empty-panel {
        min-height: 220px;
        place-content: center;
        text-align: center;
        color: var(--muted);
        padding: 24px;
      }
      .empty-panel.compact { min-height: 180px; }
      .alert {
        margin: 0 14px 14px;
        padding: 12px 14px;
        border-radius: 12px;
        border: 1px solid rgba(255, 143, 143, 0.22);
        background: rgba(255, 143, 143, 0.08);
        color: #ffd0d0;
        font-size: 13px;
      }
      .reader .alert { margin: 0 0 18px; }
      @media (prefers-reduced-motion: reduce) {
        *, *::before, *::after { transition: none !important; animation: none !important; }
      }
      @media (max-width: 1180px) {
        .workspace { grid-template-columns: 252px 336px minmax(320px, 1fr); }
        .reader-title-row h1 { font-size: 28px; }
      }
      @media (max-width: 920px) {
        .workspace { grid-template-columns: 1fr; }
        .pane { min-height: auto; }
        .pane + .pane { border-left: 0; border-top: 1px solid var(--line); }
        .reader { padding: 18px; }
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
      <div class="login-page">
        <section class="login-panel">
          <div>
            <div class="eyebrow">Mail Console</div>
            <h1>登录只读控制台</h1>
          </div>
          <p>用管理员密码进入多账号 Outlook 阅读台。这里专注于查看邮箱状态、消息列表和正文，不承担管理写操作。</p>
          ${
            input.configured
              ? `
                ${input.error ? `<div class="alert" style="margin:0;">${escapeHtml(input.error)}</div>` : ""}
                <form method="POST" action="/app/login">
                  <input class="input" type="password" name="password" placeholder="输入管理员密码" autocomplete="current-password" required />
                  <button class="button-link" type="submit">进入 Mail Console</button>
                </form>
              `
              : `
                <div class="alert" style="margin:0;">当前未配置 <code>WEB_ADMIN_PASSWORD</code>，Web 控制台尚未启用。</div>
              `
          }
        </section>
      </div>
    `,
  );
}

export function renderAppPage(state: WebConsoleState): Response {
  const selectedMailboxId = state.selectedMailbox?.connection.mailboxId;
  const selectedFolder = state.selectedFolder;
  const selectedMailboxLabel = state.selectedMailbox?.connection.displayName || state.selectedMailbox?.connection.emailAddress || "No mailbox";

  return renderAppShell(
    "Mail Console",
    `
      <div class="app-shell">
        <header class="topbar">
          <div class="brand-stack">
            <div class="brand-title">Mail Console</div>
            <div class="brand-subtitle">多账号 Outlook 只读工作台</div>
          </div>
          <div class="toolbar-meta">
            <span class="pill">${state.mailboxes.length} mailboxes</span>
            <span class="pill">${escapeHtml(selectedMailboxLabel)}</span>
            <span class="pill">${escapeHtml(formatFolderLabel(selectedFolder))}</span>
            <form method="POST" action="/app/logout">
              <button class="button-ghost" type="submit">退出</button>
            </form>
          </div>
        </header>

        <div class="workspace">
          <aside class="pane rail">
            <div class="section-heading">Mailboxes</div>
            ${
              state.mailboxes.length > 0
                ? `<nav class="mailbox-list">${state.mailboxes.map((mailbox) => renderMailboxItem(mailbox, selectedMailboxId, selectedFolder)).join("")}</nav>`
                : renderEmptyMailboxes()
            }
            <div class="footer-note">
              管理动作仍建议在 Slack 完成：<br />
              <code>/mail connect graph</code><br />
              <code>/mail route &lt;mailbox&gt; &lt;#channel&gt;</code><br />
              <code>/mail provider &lt;mailbox&gt; graph</code>
            </div>
          </aside>

          <section class="pane stream">
            <div class="stream-header">
              <div class="stream-title-row">
                <div class="stream-title">
                  <h2>${escapeHtml(selectedMailboxLabel)}</h2>
                  <p>${state.selectedMailbox ? escapeHtml(state.selectedMailbox.connection.emailAddress) : "连接邮箱后可在这里查看邮件流。"}</p>
                </div>
                ${state.selectedMailbox ? `<span class="pill">${escapeHtml(providerLabel(state.selectedMailbox))}</span>` : ""}
              </div>
              ${state.selectedMailbox
                ? `
                  <div class="tab-row">
                    <a class="tab${selectedFolder === "inbox" ? " is-active" : ""}" href="${appHref({ mailboxId: state.selectedMailbox.connection.mailboxId, folder: "inbox" })}">Inbox</a>
                    <a class="tab${selectedFolder === "junk" ? " is-active" : ""}" href="${appHref({ mailboxId: state.selectedMailbox.connection.mailboxId, folder: "junk" })}">Junk</a>
                  </div>
                `
                : ""
              }
            </div>
            ${state.error ? `<div class="alert">${escapeHtml(state.error)}</div>` : ""}
            ${state.messages.length > 0
              ? `<div class="message-list">${state.messages.map((message) => renderMessageItem(state, message)).join("")}</div>`
              : `
                <section class="empty-panel compact">
                  <h3>这个文件夹里暂时没有可展示邮件</h3>
                  <p>如果邮箱刚接入，可以先等待同步，或在 Slack 中执行 <code>/mail sync &lt;mailbox&gt;</code>。</p>
                </section>
              `}
          </section>

          <main class="pane reader">
            ${renderMessageBody(state.selectedMessage, state)}
          </main>
        </div>
      </div>
    `,
  );
}
