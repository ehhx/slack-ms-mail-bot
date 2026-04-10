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
  pageCursor?: string | null;
  page?: number | null;
}): string {
  const params = new URLSearchParams();
  if (input.mailboxId) params.set("mailbox", input.mailboxId);
  if (input.folder) params.set("folder", input.folder);
  if (input.messageId) params.set("message", input.messageId);
  if (input.pageCursor) params.set("pageCursor", input.pageCursor);
  if (input.page && input.page > 1) params.set("page", String(input.page));
  const query = params.toString();
  return query ? `/app?${query}` : "/app";
}

function providerLabel(bundle: MailboxBundle): string {
  return bundle.connection.providerType === "ms_oauth2api" ? "msOauth2api" : "Graph 原生";
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
    const dataUrl = dataUrlForInlineImage(image);
    const contentId = normalizeContentId(image.contentId);
    if (contentId) contentIdMap.set(contentId, dataUrl);
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
        --line: #d7e1ee;
        --accent: #2563eb;
        --bg: #ffffff;
      }
      * { box-sizing: border-box; }
      html, body {
        margin: 0;
        padding: 0;
        background: var(--bg);
        color: var(--text);
        font: 15px/1.68 Inter, ui-sans-serif, system-ui, -apple-system, BlinkMacSystemFont, sans-serif;
      }
      body { padding: 30px; }
      img { max-width: 100%; height: auto; }
      table { max-width: 100% !important; }
      pre, code { white-space: pre-wrap; word-break: break-word; }
      a { color: var(--accent); }
      blockquote {
        margin: 1.2rem 0;
        padding-left: 1rem;
        border-left: 3px solid var(--line);
        color: var(--muted);
      }
      .mail-document { min-height: calc(100vh - 60px); }
    </style>
  </head>
  <body>
    <main class="mail-document">${sanitized}</main>
  </body>
</html>`;
}

function renderReaderIntro(state: WebConsoleState): string {
  if (!state.selectedMailbox) {
    return `
      <section class="reader-intro">
        <div class="reader-kicker">阅读区</div>
        <h1>先连接一个邮箱</h1>
        <p>先在 Slack 里运行 <code>/mail connect graph</code>，接入至少一个 Outlook 账号后，这里才会显示消息阅读区。</p>
      </section>
    `;
  }

  return `
    <section class="reader-intro">
      <div class="reader-kicker">阅读区</div>
      <h1>选择一封邮件开始阅读</h1>
      <p>消息流会先快速加载，正文、附件和内联图片只在你真正点开邮件时再读取，避免整个界面一起变慢。</p>
      <div class="reader-inline-meta">
        <div><span>邮箱</span><strong>${escapeHtml(state.selectedMailbox.connection.displayName || state.selectedMailbox.connection.emailAddress)}</strong></div>
        <div><span>文件夹</span><strong>${escapeHtml(formatFolderLabel(state.selectedFolder))}</strong></div>
        <div><span>Slack 路由</span><strong>${escapeHtml(state.selectedMailbox.route?.slackChannelName || state.selectedMailbox.route?.slackChannelId || "未配置")}</strong></div>
        <div><span>监控范围</span><strong>${escapeHtml(monitoredFoldersText(state.selectedMailbox))}</strong></div>
      </div>
    </section>
  `;
}

function renderMessageBody(detail: WebMessageDetail | null, state: WebConsoleState): string {
  if (!detail) {
    return renderReaderIntro(state);
  }

  const attachments = detail.message.attachments ?? [];
  const htmlBody = detail.bodyHtml?.trim();
  const bodyBlock = htmlBody
    ? `
      <iframe
        class="mail-body-frame"
        title="邮件正文"
        loading="lazy"
        sandbox="allow-popups allow-popups-to-escape-sandbox"
        srcdoc="${escapeHtml(buildReaderSrcdoc(htmlBody, detail.message.inlineImages))}"
      ></iframe>
    `
    : `<pre class="mail-body-text">${escapeHtml(detail.bodyPlainText || "(无可用正文)")}</pre>`;

  return `
    <article class="reader-document">
      <header class="reader-header">
        <div class="reader-kicker">${escapeHtml(formatFolderLabel(detail.message.folderKind, detail.message.folderName))}</div>
        <div class="reader-title-row">
          <h1>${escapeHtml(detail.message.subject || "(无主题)")}</h1>
          ${detail.message.webLink
            ? `<a class="reader-action" href="${escapeHtml(detail.message.webLink)}" target="_blank" rel="noopener noreferrer">在 Outlook 中打开</a>`
            : ""}
        </div>
        <div class="reader-byline">
          <span>${escapeHtml(detail.message.fromName || detail.message.fromAddress || "未知发件人")}</span>
          <span>${escapeHtml(detail.message.fromAddress || "")}</span>
        </div>
      </header>

      <section class="reader-statline">
        <div><span>接收时间</span><strong>${escapeHtml(fmtTime(detail.message.receivedDateTime))}</strong></div>
        <div><span>附件数量</span><strong>${attachments.length}</strong></div>
        <div><span>内联图片</span><strong>${detail.message.inlineImages?.length ?? 0}</strong></div>
        <div><span>正文类型</span><strong>${escapeHtml(detail.message.bodyContentType || "text")}</strong></div>
      </section>

      ${attachments.length > 0
        ? `
          <section class="reader-section">
            <div class="reader-section-title">附件</div>
            <ul class="attachment-list">
              ${attachments.map((attachment) =>
                `<li>
                  <strong>${escapeHtml(attachment.name)}</strong>
                  <span>${attachment.contentType ? escapeHtml(attachment.contentType) : "未知类型"}</span>
                  <span>${attachment.size ? `${Math.max(1, Math.round(attachment.size / 1024))} KB` : "-"}</span>
                </li>`
              ).join("")}
            </ul>
          </section>
        `
        : ""
      }

      <section class="reader-section reader-section-body">
        <div class="reader-section-title">正文</div>
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
    <a class="mailbox-item${active ? " is-active" : ""}" title="${escapeHtml(mailbox.connection.emailAddress)}" href="${appHref({
      mailboxId: mailbox.connection.mailboxId,
      folder: selectedFolder,
    })}">
      <div class="mailbox-title">${escapeHtml(mailbox.connection.displayName || mailbox.connection.emailAddress)}</div>
      <div class="mailbox-subtitle">${escapeHtml(mailbox.connection.emailAddress)}</div>
      <div class="mailbox-meta">
        <span>${escapeHtml(providerLabel(mailbox))}</span>
        <span>${escapeHtml(mailbox.route?.slackChannelName || mailbox.route?.slackChannelId || "未配置")}</span>
      </div>
    </a>
  `;
}

function renderMailboxSwitcher(state: WebConsoleState): string {
  if (state.mailboxes.length === 0 || !state.selectedMailbox) {
    return `
      <div class="mailbox-switcher-empty">
        <span>未连接邮箱</span>
      </div>
    `;
  }

  return `
    <details class="mailbox-switcher">
      <summary class="mailbox-switcher-trigger">
        <span class="mailbox-switcher-badge">邮箱</span>
        <span class="mailbox-switcher-copy">
          <strong>${escapeHtml(
            state.selectedMailbox.connection.displayName || state.selectedMailbox.connection.emailAddress,
          )}</strong>
          <span>${state.mailboxes.length} 个账号 · ${escapeHtml(state.selectedMailbox.connection.emailAddress)}</span>
        </span>
        <span class="mailbox-switcher-caret" aria-hidden="true">▾</span>
      </summary>
      <div class="mailbox-switcher-menu">
        ${state.mailboxes.map((mailbox) => `
          <a
            class="mailbox-switcher-option${mailbox.connection.mailboxId === state.selectedMailbox?.connection.mailboxId ? " is-active" : ""}"
            href="${appHref({
              mailboxId: mailbox.connection.mailboxId,
              folder: state.selectedFolder,
            })}"
          >
            <div class="mailbox-switcher-option-title">${escapeHtml(mailbox.connection.displayName || mailbox.connection.emailAddress)}</div>
            <div class="mailbox-switcher-option-subtitle">${escapeHtml(mailbox.connection.emailAddress)}</div>
            <div class="mailbox-switcher-option-meta">
              <span>${escapeHtml(providerLabel(mailbox))}</span>
              <span>${escapeHtml(mailbox.route?.slackChannelName || mailbox.route?.slackChannelId || "未配置")}</span>
            </div>
          </a>
        `).join("")}
      </div>
    </details>
  `;
}

function renderMessageItem(
  state: WebConsoleState,
  message: {
    messageId: string;
    subject: string;
    fromName?: string;
    fromAddress?: string;
    bodyPreview?: string;
    receivedDateTime?: string;
    hasAttachments?: boolean;
  },
): string {
  const selectedMessageId = state.selectedMessage?.message.messageId;
  const active = message.messageId === selectedMessageId;
  return `
    <a class="message-item${active ? " is-active" : ""}" href="${appHref({
      mailboxId: state.selectedMailbox?.connection.mailboxId,
      folder: state.selectedFolder,
      messageId: message.messageId,
      pageCursor: state.currentPageCursor,
      page: state.pageIndex > 1 ? state.pageIndex : undefined,
    })}">
      <div class="message-row-top">
        <span class="message-subject">${escapeHtml(message.subject || "(无主题)")}</span>
        <span class="message-time">${escapeHtml(fmtTime(message.receivedDateTime))}</span>
      </div>
      <div class="message-sender">${escapeHtml(message.fromName || message.fromAddress || "未知发件人")}</div>
      <div class="message-preview">${escapeHtml(message.bodyPreview || "(无预览)")}</div>
      ${message.hasAttachments ? `<div class="message-meta">含附件</div>` : ""}
    </a>
  `;
}

function renderMessagePagination(state: WebConsoleState): string {
  if (!state.selectedMailbox) return "";
  if (!state.nextPageCursor && !state.hasPreviousPage) return "";

  const latestHref = appHref({
    mailboxId: state.selectedMailbox.connection.mailboxId,
    folder: state.selectedFolder,
    messageId: state.selectedMessage?.message.messageId,
  });
  const olderHref = state.nextPageCursor
    ? appHref({
      mailboxId: state.selectedMailbox.connection.mailboxId,
      folder: state.selectedFolder,
      messageId: state.selectedMessage?.message.messageId,
      pageCursor: state.nextPageCursor,
      page: state.pageIndex + 1,
    })
    : null;

  return `
    <div class="stream-pagination">
      <div class="stream-pagination-copy">
        <span class="section-label">分页</span>
        <strong>${state.pageIndex === 1 ? "最新邮件" : `第 ${state.pageIndex} 页`}</strong>
      </div>
      <div class="stream-pagination-actions">
        ${state.hasPreviousPage ? `<a class="stream-page-link" href="${latestHref}">回到最新</a>` : ""}
        ${olderHref
          ? `<a class="stream-page-link is-primary" href="${olderHref}">更早邮件</a>`
          : `<span class="stream-page-link is-disabled">没有更早邮件了</span>`}
      </div>
    </div>
  `;
}

function renderEmptyMailboxes(): string {
  return `
    <section class="empty-note">
      <h3>暂无邮箱</h3>
      <p>先在 Slack 里执行 <code>/mail connect graph</code>，连接至少一个 Outlook 账号后，这里才会显示消息流和阅读区。</p>
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
        --bg: #050810;
        --shell: #0a111b;
        --shell-2: #0d1521;
        --shell-3: #101927;
        --line: rgba(148, 163, 184, 0.12);
        --line-strong: rgba(148, 163, 184, 0.24);
        --text: #edf4ff;
        --muted: #92a5c0;
        --accent: #7aaeff;
        --accent-soft: rgba(122, 174, 255, 0.12);
        --accent-strong: rgba(122, 174, 255, 0.22);
        --reader-bg: #eef3f8;
        --paper: #ffffff;
        --paper-soft: #fbfdff;
        --reader-line: #d9e3ef;
        --reader-text: #111827;
        --reader-muted: #5f7086;
        --shadow: 0 30px 100px rgba(15, 23, 42, 0.10), 0 10px 30px rgba(15, 23, 42, 0.06);
      }
      * { box-sizing: border-box; }
      html, body {
        margin: 0;
        min-height: 100%;
        background:
          radial-gradient(circle at top left, rgba(74, 116, 196, 0.14), transparent 24%),
          linear-gradient(180deg, #07101a 0%, var(--bg) 100%);
        color: var(--text);
        font-family: Inter, ui-sans-serif, system-ui, -apple-system, BlinkMacSystemFont, sans-serif;
      }
      body {
        -webkit-font-smoothing: antialiased;
        text-rendering: optimizeLegibility;
      }
      a { color: inherit; text-decoration: none; }
      code {
        font-family: ui-monospace, SFMono-Regular, Menlo, Monaco, monospace;
        font-size: 0.92em;
      }
      .app-shell {
        min-height: 100vh;
        display: grid;
        grid-template-rows: 68px 1fr;
      }
      .topbar {
        display: flex;
        align-items: center;
        justify-content: space-between;
        gap: 16px;
        padding: 0 22px;
        border-bottom: 1px solid var(--line);
        background: rgba(7, 11, 18, 0.96);
      }
      .topbar-left {
        display: flex;
        align-items: center;
        gap: 18px;
        min-width: 0;
      }
      .brand {
        display: flex;
        align-items: center;
        gap: 12px;
        flex: 0 0 auto;
      }
      .brand-mark {
        width: 11px;
        height: 11px;
        border-radius: 999px;
        background: linear-gradient(135deg, #bfd8ff 0%, var(--accent) 100%);
        box-shadow: 0 0 0 10px rgba(122, 174, 255, 0.08);
      }
      .brand-copy {
        display: grid;
        gap: 2px;
      }
      .brand-kicker {
        font-size: 11px;
        font-weight: 700;
        letter-spacing: 0.14em;
        text-transform: uppercase;
        color: var(--muted);
      }
      .brand-title {
        font-size: 17px;
        font-weight: 700;
        letter-spacing: -0.02em;
      }
      .topbar-meta {
        display: flex;
        align-items: center;
        gap: 16px;
        flex-wrap: wrap;
        justify-content: flex-end;
      }
      .meta-block {
        display: grid;
        gap: 1px;
      }
      .meta-block span {
        font-size: 11px;
        letter-spacing: 0.1em;
        text-transform: uppercase;
        color: var(--muted);
      }
      .meta-block strong {
        font-size: 13px;
        font-weight: 600;
        color: var(--text);
      }
      .ghost-button {
        display: inline-flex;
        align-items: center;
        justify-content: center;
        padding: 10px 15px;
        border-radius: 999px;
        border: 1px solid var(--line-strong);
        background: rgba(255, 255, 255, 0.02);
        color: var(--text);
        cursor: pointer;
      }
      .mailbox-switcher,
      .mailbox-switcher-empty {
        position: relative;
        flex: 0 1 360px;
        min-width: 0;
      }
      .mailbox-switcher-trigger,
      .mailbox-switcher-empty {
        display: inline-flex;
        align-items: center;
        gap: 12px;
        width: min(360px, 100%);
        min-height: 48px;
        padding: 8px 14px;
        border: 1px solid var(--line-strong);
        border-radius: 18px;
        background: rgba(255, 255, 255, 0.02);
      }
      .mailbox-switcher-empty {
        color: var(--muted);
      }
      .mailbox-switcher summary {
        list-style: none;
        cursor: pointer;
      }
      .mailbox-switcher summary::-webkit-details-marker {
        display: none;
      }
      .mailbox-switcher-badge {
        flex: 0 0 auto;
        display: inline-flex;
        align-items: center;
        justify-content: center;
        min-width: 40px;
        height: 28px;
        padding: 0 10px;
        border-radius: 999px;
        background: rgba(122, 174, 255, 0.12);
        color: var(--accent);
        font-size: 11px;
        font-weight: 700;
        letter-spacing: 0.08em;
        text-transform: uppercase;
      }
      .mailbox-switcher-copy {
        min-width: 0;
        display: grid;
        gap: 2px;
      }
      .mailbox-switcher-copy strong {
        font-size: 14px;
        font-weight: 650;
        color: var(--text);
        white-space: nowrap;
        overflow: hidden;
        text-overflow: ellipsis;
      }
      .mailbox-switcher-copy span {
        font-size: 12px;
        color: var(--muted);
        white-space: nowrap;
        overflow: hidden;
        text-overflow: ellipsis;
      }
      .mailbox-switcher-caret {
        margin-left: auto;
        color: var(--muted);
        font-size: 14px;
      }
      .mailbox-switcher[open] .mailbox-switcher-caret {
        transform: rotate(180deg);
      }
      .mailbox-switcher-menu {
        position: absolute;
        top: calc(100% + 10px);
        left: 0;
        z-index: 30;
        width: min(380px, calc(100vw - 32px));
        display: grid;
        gap: 6px;
        padding: 8px;
        border: 1px solid var(--line-strong);
        border-radius: 20px;
        background: rgba(8, 13, 22, 0.98);
        box-shadow: 0 18px 48px rgba(0, 0, 0, 0.34);
      }
      .mailbox-switcher-option {
        display: grid;
        gap: 4px;
        padding: 12px 14px;
        border-radius: 14px;
      }
      .mailbox-switcher-option:hover {
        background: rgba(255, 255, 255, 0.03);
      }
      .mailbox-switcher-option.is-active {
        background: linear-gradient(90deg, var(--accent-soft) 0%, rgba(122, 174, 255, 0.03) 100%);
      }
      .mailbox-switcher-option-title {
        font-size: 14px;
        font-weight: 650;
        color: var(--text);
      }
      .mailbox-switcher-option-subtitle,
      .mailbox-switcher-option-meta {
        font-size: 12px;
        color: var(--muted);
      }
      .mailbox-switcher-option-meta {
        display: flex;
        align-items: center;
        gap: 12px;
      }
      .workspace {
        min-height: calc(100vh - 68px);
        display: grid;
        grid-template-columns: minmax(320px, 376px) minmax(0, 1fr);
        grid-template-areas: "stream reader";
      }
      .pane {
        min-height: calc(100vh - 68px);
        overflow: auto;
        scrollbar-gutter: stable;
        contain: content;
        overscroll-behavior: contain;
        scrollbar-width: thin;
      }
      .pane + .pane { border-left: 1px solid var(--line); }
      .stream-pane {
        grid-area: stream;
        padding: 0 0 30px;
        background: var(--shell-2);
      }
      .reader-pane {
        grid-area: reader;
        background: linear-gradient(180deg, #eff4fa 0%, var(--reader-bg) 100%);
        color: var(--reader-text);
      }
      .section-label {
        font-size: 11px;
        font-weight: 700;
        letter-spacing: 0.14em;
        text-transform: uppercase;
        color: var(--muted);
      }
      .stream-head {
        display: grid;
        gap: 10px;
        position: sticky;
        top: 0;
        z-index: 4;
        padding: 22px 18px 16px;
      }
      .stream-head {
        background: linear-gradient(180deg, rgba(13, 21, 33, 0.99) 0%, rgba(13, 21, 33, 0.92) 72%, rgba(13, 21, 33, 0) 100%);
      }
      .stream-head h2,
      .empty-note h3,
      .login-panel h1,
      .reader-intro h1,
      .reader-title-row h1 {
        margin: 0;
      }
      .stream-head h2 {
        font-size: 21px;
        letter-spacing: -0.03em;
      }
      .stream-head p,
      .empty-note p,
      .login-panel p,
      .reader-intro p {
        margin: 0;
        color: var(--muted);
        line-height: 1.6;
      }
      .message-list {
        display: grid;
        gap: 4px;
        content-visibility: auto;
        contain-intrinsic-size: 720px;
      }
      .message-item {
        position: relative;
        display: grid;
        gap: 7px;
        transition: background-color 140ms ease, color 140ms ease, transform 140ms ease;
        content-visibility: auto;
        contain-intrinsic-size: 96px;
      }
      .message-item {
        padding: 18px 20px 16px 24px;
        border-bottom: 1px solid var(--line);
      }
      .message-item::before {
        content: "";
        position: absolute;
        left: 0;
        top: 16px;
        bottom: 16px;
        width: 3px;
        border-radius: 999px;
        background: transparent;
      }
      .message-item:hover {
        background: rgba(255, 255, 255, 0.03);
        transform: translateX(1px);
      }
      .message-item.is-active {
        background: linear-gradient(90deg, var(--accent-soft) 0%, rgba(122, 174, 255, 0.03) 100%);
      }
      .message-item.is-active::before { background: var(--accent); }
      .message-subject {
        font-size: 15px;
        font-weight: 650;
        line-height: 1.35;
      }
      .message-sender,
      .message-preview,
      .message-time,
      .message-meta {
        font-size: 13px;
        color: var(--muted);
      }
      .message-sender,
      .message-preview { line-height: 1.5; }
      .message-row-top {
        display: flex;
        align-items: flex-start;
        justify-content: space-between;
        gap: 12px;
      }
      .message-time { white-space: nowrap; }
      .message-preview {
        display: -webkit-box;
        -webkit-box-orient: vertical;
        -webkit-line-clamp: 2;
        overflow: hidden;
      }
      .message-meta {
        width: fit-content;
        padding: 4px 8px;
        border-radius: 999px;
        background: rgba(255, 255, 255, 0.04);
        font-size: 11px;
        letter-spacing: 0.08em;
        text-transform: uppercase;
      }
      .folder-tabs {
        display: flex;
        gap: 22px;
        align-items: center;
      }
      .folder-tab {
        position: relative;
        padding: 6px 0 10px;
        font-size: 13px;
        color: var(--muted);
      }
      .folder-tab.is-active {
        color: var(--text);
      }
      .folder-tab.is-active::after {
        content: "";
        position: absolute;
        left: 0;
        right: 0;
        bottom: 0;
        height: 2px;
        border-radius: 999px;
        background: var(--accent);
      }
      .stream-headline {
        display: grid;
        gap: 6px;
      }
      .stream-meta {
        font-size: 13px;
        color: var(--muted);
      }
      .stream-alert {
        margin: 0 18px 14px;
        padding: 12px 14px;
        border-radius: 14px;
        border: 1px solid rgba(220, 38, 38, 0.24);
        background: rgba(220, 38, 38, 0.08);
        color: #fecaca;
        font-size: 13px;
      }
      .stream-pagination {
        display: flex;
        align-items: center;
        justify-content: space-between;
        gap: 16px;
        margin: 16px 18px 0;
        padding-top: 18px;
        border-top: 1px solid var(--line);
      }
      .stream-pagination-copy {
        display: grid;
        gap: 4px;
      }
      .stream-pagination-copy strong {
        font-size: 13px;
        font-weight: 600;
        color: var(--text);
      }
      .stream-pagination-actions {
        display: flex;
        gap: 10px;
        align-items: center;
        flex-wrap: wrap;
      }
      .stream-page-link {
        display: inline-flex;
        align-items: center;
        justify-content: center;
        min-height: 36px;
        padding: 0 14px;
        border-radius: 999px;
        border: 1px solid var(--line-strong);
        color: var(--text);
        font-size: 13px;
      }
      .stream-page-link.is-primary {
        background: rgba(255, 255, 255, 0.04);
      }
      .stream-page-link.is-disabled {
        color: var(--muted);
        border-style: dashed;
      }
      .reader-wrap {
        min-height: calc(100vh - 68px);
        padding: 36px 42px 56px;
        display: flex;
        justify-content: center;
        align-items: flex-start;
      }
      .reader-intro,
      .reader-document {
        display: grid;
        gap: 28px;
        width: min(1080px, 100%);
        padding: 46px 54px 60px;
        background: linear-gradient(180deg, var(--paper) 0%, var(--paper-soft) 100%);
        border-radius: 32px;
        box-shadow: 0 16px 48px rgba(15, 23, 42, 0.08);
        position: relative;
        overflow: hidden;
      }
      .reader-intro::before,
      .reader-document::before {
        content: "";
        position: absolute;
        inset: 0 0 auto 0;
        height: 112px;
        background: linear-gradient(180deg, rgba(122, 174, 255, 0.10) 0%, rgba(122, 174, 255, 0) 100%);
        pointer-events: none;
      }
      .reader-intro > *,
      .reader-document > * {
        position: relative;
        z-index: 1;
      }
      .reader-kicker {
        font-size: 11px;
        font-weight: 700;
        letter-spacing: 0.14em;
        text-transform: uppercase;
        color: #466a96;
      }
      .reader-intro h1,
      .reader-title-row h1 {
        font-size: clamp(38px, 4.2vw, 62px);
        line-height: 1.04;
        letter-spacing: -0.035em;
        color: var(--reader-text);
        max-width: 14ch;
      }
      .reader-intro p,
      .reader-byline {
        font-size: 15px;
        color: var(--reader-muted);
        line-height: 1.7;
        max-width: 70ch;
      }
      .reader-inline-meta,
      .reader-statline {
        display: grid;
        grid-template-columns: repeat(4, minmax(0, 1fr));
        border-top: 1px solid var(--reader-line);
        border-bottom: 1px solid var(--reader-line);
        background: rgba(250, 252, 255, 0.8);
      }
      .reader-inline-meta > div,
      .reader-statline > div {
        display: grid;
        gap: 6px;
        padding: 16px 0;
      }
      .reader-inline-meta > div + div,
      .reader-statline > div + div {
        padding-left: 18px;
        margin-left: 18px;
        border-left: 1px solid var(--reader-line);
      }
      .reader-inline-meta span,
      .reader-statline span {
        font-size: 11px;
        letter-spacing: 0.1em;
        text-transform: uppercase;
        color: var(--reader-muted);
      }
      .reader-inline-meta strong,
      .reader-statline strong {
        font-size: 14px;
        font-weight: 600;
        color: var(--reader-text);
      }
      .reader-header {
        display: grid;
        gap: 18px;
      }
      .reader-title-row {
        display: flex;
        align-items: flex-start;
        justify-content: space-between;
        gap: 24px;
      }
      .reader-action {
        display: inline-flex;
        align-items: center;
        justify-content: center;
        padding: 11px 16px;
        border-radius: 999px;
        border: 1px solid #0f172a;
        background: #0f172a;
        color: #f8fbff;
        font-size: 13px;
        white-space: nowrap;
      }
      .reader-byline {
        display: flex;
        flex-wrap: wrap;
        gap: 12px;
      }
      .reader-section {
        display: grid;
        gap: 12px;
      }
      .reader-section-title {
        font-size: 11px;
        font-weight: 700;
        letter-spacing: 0.14em;
        text-transform: uppercase;
        color: var(--reader-muted);
      }
      .attachment-list {
        list-style: none;
        margin: 0;
        padding: 0;
        border-top: 1px solid var(--reader-line);
      }
      .attachment-list li {
        display: grid;
        grid-template-columns: minmax(0, 1fr) auto auto;
        gap: 18px;
        align-items: center;
        padding: 14px 0;
        border-bottom: 1px solid var(--reader-line);
        color: var(--reader-text);
        font-size: 14px;
      }
      .attachment-list li span {
        color: var(--reader-muted);
        font-size: 13px;
      }
      .mail-body-frame {
        width: 100%;
        min-height: 840px;
        border: 0;
        border-radius: 22px;
        background: #ffffff;
        box-shadow: inset 0 0 0 1px var(--reader-line);
      }
      .mail-body-text {
        margin: 0;
        padding: 30px 32px;
        white-space: pre-wrap;
        word-break: break-word;
        line-height: 1.8;
        color: var(--reader-text);
        background: #ffffff;
        border: 0;
        border-radius: 22px;
        box-shadow: inset 0 0 0 1px var(--reader-line);
      }
      .empty-note {
        display: grid;
        gap: 8px;
        padding: 20px 18px 0;
      }
      .login-page {
        min-height: 100vh;
        display: grid;
        place-items: center;
        padding: 28px;
        background: radial-gradient(circle at 50% 0%, rgba(122, 174, 255, 0.10), transparent 28%);
      }
      .login-panel {
        width: min(520px, 100%);
        display: grid;
        gap: 18px;
        padding: 42px;
        border-radius: 30px;
        border: 1px solid var(--line-strong);
        background: linear-gradient(180deg, rgba(13, 20, 32, 0.92) 0%, rgba(10, 16, 26, 0.98) 100%);
        box-shadow: 0 18px 52px rgba(0, 0, 0, 0.22);
      }
      .login-panel form {
        display: grid;
        gap: 12px;
      }
      .login-input {
        width: 100%;
        padding: 14px 16px;
        border-radius: 16px;
        border: 1px solid var(--line-strong);
        background: rgba(255, 255, 255, 0.03);
        color: var(--text);
        outline: none;
      }
      .login-button {
        display: inline-flex;
        align-items: center;
        justify-content: center;
        padding: 14px 18px;
        border-radius: 999px;
        background: var(--accent);
        color: #08111e;
        border: 0;
        font-weight: 600;
        cursor: pointer;
      }
      .login-alert {
        padding: 12px 14px;
        border: 1px solid rgba(220, 38, 38, 0.24);
        background: rgba(220, 38, 38, 0.08);
        color: #fecaca;
        font-size: 13px;
      }
      .login-muted {
        color: var(--muted);
        line-height: 1.7;
      }
      @media (prefers-reduced-motion: reduce) {
        *, *::before, *::after { transition: none !important; animation: none !important; }
      }
      @media (max-width: 1440px) {
        .workspace { grid-template-columns: minmax(300px, 344px) minmax(0, 1fr); }
        .reader-wrap { padding: 28px 28px 42px; }
        .reader-intro,
        .reader-document { padding: 38px 40px 46px; }
      }
      @media (max-width: 960px) {
        .app-shell { grid-template-rows: auto 1fr; }
        .topbar {
          padding: 14px 16px;
          align-items: flex-start;
        }
        .topbar-left {
          width: 100%;
          flex-direction: column;
          align-items: stretch;
          gap: 12px;
        }
        .mailbox-switcher,
        .mailbox-switcher-empty {
          flex-basis: auto;
        }
        .mailbox-switcher-trigger,
        .mailbox-switcher-empty {
          width: 100%;
        }
        .mailbox-switcher-menu {
          width: 100%;
        }
        .workspace {
          grid-template-columns: 1fr;
          grid-template-areas:
            "stream"
            "reader";
        }
        .pane { min-height: auto; }
        .pane + .pane { border-left: 0; border-top: 1px solid var(--line); }
        .reader-wrap { min-height: auto; }
        .reader-inline-meta,
        .reader-statline { grid-template-columns: repeat(2, minmax(0, 1fr)); }
        .reader-inline-meta > div:nth-child(3),
        .reader-statline > div:nth-child(3) { padding-left: 0; margin-left: 0; border-left: 0; }
        .mail-body-frame { min-height: 520px; }
        .stream-pagination {
          align-items: flex-start;
          flex-direction: column;
        }
      }
      @media (max-width: 720px) {
        .topbar-meta {
          display: grid;
          grid-template-columns: repeat(2, minmax(0, 1fr));
          width: 100%;
        }
        .mailbox-switcher-trigger {
          align-items: flex-start;
        }
        .mailbox-switcher-badge {
          min-width: 34px;
          height: 26px;
          padding: 0 8px;
        }
        .mailbox-switcher-copy strong,
        .mailbox-switcher-copy span {
          white-space: normal;
        }
        .ghost-button { width: fit-content; }
        .reader-wrap { padding: 18px 16px 28px; }
        .reader-intro,
        .reader-document {
          padding: 28px 22px 34px;
          border-radius: 24px;
          gap: 22px;
        }
        .reader-title-row {
          flex-direction: column;
          gap: 16px;
        }
        .reader-inline-meta,
        .reader-statline { grid-template-columns: 1fr; }
        .reader-inline-meta > div + div,
        .reader-statline > div + div {
          padding-left: 0;
          margin-left: 0;
          border-left: 0;
        }
        .mail-body-frame { min-height: 440px; border-radius: 18px; }
        .mail-body-text { border-radius: 18px; }
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
          <div class="section-label">邮件工作台</div>
          <h1>登录只读邮件工作台</h1>
          <p class="login-muted">这里专注于多账号 Outlook 阅读。连接、路由和同步管理仍建议在 Slack 中完成。</p>
          ${input.configured
            ? `
              ${input.error ? `<div class="login-alert">${escapeHtml(input.error)}</div>` : ""}
              <form method="POST" action="/app/login">
                <input class="login-input" type="password" name="password" placeholder="输入管理员密码" autocomplete="current-password" required />
                <button class="login-button" type="submit">进入邮件工作台</button>
              </form>
            `
            : `<div class="login-alert">当前未配置 <code>WEB_ADMIN_PASSWORD</code>，Web 控制台尚未启用。</div>`}
        </section>
      </div>
    `,
  );
}

export function renderAppPage(state: WebConsoleState): Response {
  const selectedMailboxId = state.selectedMailbox?.connection.mailboxId;
  const selectedFolder = state.selectedFolder;
  const mailboxLabel = state.selectedMailbox?.connection.displayName || state.selectedMailbox?.connection.emailAddress || "No mailbox";

  return renderAppShell(
    "Mail Console",
    `
      <div class="app-shell">
        <header class="topbar">
          <div class="topbar-left">
            <div class="brand">
              <div class="brand-mark"></div>
              <div class="brand-copy">
                <div class="brand-kicker">邮件工作台</div>
                <div class="brand-title">多账号 Outlook 阅读台</div>
              </div>
            </div>
            ${renderMailboxSwitcher(state)}
          </div>
          <div class="topbar-meta">
            <div class="meta-block">
              <span>当前邮箱</span>
              <strong>${escapeHtml(mailboxLabel)}</strong>
            </div>
            <div class="meta-block">
              <span>当前文件夹</span>
              <strong>${escapeHtml(formatFolderLabel(selectedFolder))}</strong>
            </div>
            <div class="meta-block">
              <span>已载入</span>
              <strong>${state.messages.length} 封邮件${state.pageIndex > 1 ? ` · 第 ${state.pageIndex} 页` : ""}</strong>
            </div>
            <form method="POST" action="/app/logout">
              <button class="ghost-button" type="submit">退出</button>
            </form>
          </div>
        </header>

        <div class="workspace">
          <section class="pane stream-pane">
            <div class="stream-head">
              <div class="section-label">消息流</div>
              <div class="stream-headline">
                <h2>${escapeHtml(formatFolderLabel(selectedFolder))}</h2>
                <div class="stream-meta">${state.selectedMailbox ? `${escapeHtml(mailboxLabel)} · ${escapeHtml(state.selectedMailbox.connection.emailAddress)}` : "连接邮箱后可在这里查看消息流。"}</div>
              </div>
              ${state.selectedMailbox
                ? `
                  <div class="folder-tabs">
                    <a class="folder-tab${selectedFolder === "inbox" ? " is-active" : ""}" href="${appHref({ mailboxId: state.selectedMailbox.connection.mailboxId, folder: "inbox" })}">收件箱</a>
                    <a class="folder-tab${selectedFolder === "junk" ? " is-active" : ""}" href="${appHref({ mailboxId: state.selectedMailbox.connection.mailboxId, folder: "junk" })}">垃圾邮件</a>
                  </div>
                `
                : ""}
            </div>
            ${state.error ? `<div class="stream-alert">${escapeHtml(state.error)}</div>` : ""}
            ${!state.selectedMailbox
              ? renderEmptyMailboxes()
              : state.messages.length > 0
              ? `
                <div class="message-list">${state.messages.map((message) => renderMessageItem(state, message)).join("")}</div>
                ${renderMessagePagination(state)}
              `
              : `
                <section class="empty-note">
                  <h3>这个文件夹里没有可展示邮件</h3>
                  <p>如果邮箱刚接入，可以先等待同步，或者在 Slack 中执行 <code>/mail sync &lt;mailbox&gt;</code>。</p>
                </section>
              `}
          </section>

          <main class="pane reader-pane">
            <div class="reader-wrap">
              ${renderMessageBody(state.selectedMessage, state)}
            </div>
          </main>
        </div>
      </div>
    `,
  );
}
