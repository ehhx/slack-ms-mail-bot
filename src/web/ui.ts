import { formatFolderLabel, monitoredFoldersText } from "../mail/message.ts";
import type { MailboxBundle, MailInlineImage } from "../mail/types.ts";
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

function fmtListTime(iso: string | undefined): string {
  if (!iso) return "-";
  const parsed = new Date(iso);
  if (Number.isNaN(parsed.getTime())) return iso;
  return parsed.toLocaleString("zh-CN", {
    month: "numeric",
    day: "numeric",
    hour: "2-digit",
    minute: "2-digit",
    hour12: false,
  });
}

function compactText(input: string | undefined): string {
  return (input ?? "").replace(/\s+/g, " ").trim();
}

/**
 * 这里只提取“像验证码”的短码：
 * 1. 优先匹配验证码/OTP 等语义关键词附近的代码；
 * 2. 再回退到关键词上下文中的纯数字短码；
 * 3. 不在无关键词场景下盲目抓数字，避免把日期或时间识别成验证码。
 */
function detectVerificationCode(input: string | undefined): string | null {
  const text = compactText(input);
  if (!text) return null;

  const directPatterns = [
    /(?:验证码|校验码|动态码|动态密码|一次性密码|登录码|安全码|提取码|确认码)\D{0,12}([A-Z0-9-]{4,10})/iu,
    /(?:verification code|security code|one[-\s]?time (?:password|code)|login code|auth(?:entication)? code|otp)\D{0,20}([A-Z0-9-]{4,10})/iu,
    /(?:code is|password is|otp is|use code)\D{0,12}([A-Z0-9-]{4,10})/iu,
  ];
  for (const pattern of directPatterns) {
    const match = text.match(pattern);
    if (match?.[1]) return match[1].toUpperCase();
  }

  const hasKeyword =
    /(验证码|校验码|动态码|动态密码|一次性密码|登录码|安全码|提取码|确认码|verification code|security code|one[-\s]?time (?:password|code)|login code|auth(?:entication)? code|otp)/iu
      .test(text);
  if (!hasKeyword) return null;

  const fallback = text.match(/\b(\d{4,8})\b/);
  return fallback?.[1] ?? null;
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

function readerFragmentHref(input: {
  mailboxId?: string;
  folder?: "inbox" | "junk";
  messageId?: string;
}): string {
  const params = new URLSearchParams();
  if (input.mailboxId) params.set("mailbox", input.mailboxId);
  if (input.folder) params.set("folder", input.folder);
  if (input.messageId) params.set("message", input.messageId);
  return `/app/reader-fragment?${params.toString()}`;
}

function providerLabel(bundle: MailboxBundle): string {
  return bundle.connection.providerType === "ms_oauth2api"
    ? "msOauth2api"
    : "Graph 原生";
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

function rewriteCidImages(
  html: string,
  inlineImages: MailInlineImage[] | undefined,
): string {
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

function sanitizeEmailHtml(
  html: string,
  inlineImages: MailInlineImage[] | undefined,
): string {
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

function buildReaderSrcdoc(
  html: string,
  inlineImages: MailInlineImage[] | undefined,
): string {
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
        <div><span>邮箱</span><strong>${
    escapeHtml(
      state.selectedMailbox.connection.displayName ||
        state.selectedMailbox.connection.emailAddress,
    )
  }</strong></div>
        <div><span>文件夹</span><strong>${
    escapeHtml(formatFolderLabel(state.selectedFolder))
  }</strong></div>
        <div><span>Slack 路由</span><strong>${
    escapeHtml(
      state.selectedMailbox.route?.slackChannelName ||
        state.selectedMailbox.route?.slackChannelId || "未配置",
    )
  }</strong></div>
        <div><span>监控范围</span><strong>${
    escapeHtml(monitoredFoldersText(state.selectedMailbox))
  }</strong></div>
      </div>
    </section>
  `;
}

function renderReaderDetail(detail: WebMessageDetail): string {
  const attachments = detail.message.attachments ?? [];
  const verificationCode = detectVerificationCode(
    `${detail.message.subject ?? ""}\n${detail.bodyPlainText ?? ""}`,
  );
  const htmlBody = detail.bodyHtml?.trim();
  const bodyBlock = htmlBody
    ? `
      <iframe
        class="mail-body-frame"
        title="邮件正文"
        loading="lazy"
        sandbox="allow-popups allow-popups-to-escape-sandbox"
        srcdoc="${
      escapeHtml(buildReaderSrcdoc(htmlBody, detail.message.inlineImages))
    }"
      ></iframe>
    `
    : `<pre class="mail-body-text">${
      escapeHtml(detail.bodyPlainText || "(无可用正文)")
    }</pre>`;

  return `
    <article class="reader-document">
      <header class="reader-header">
        <div class="reader-kicker">${
    escapeHtml(
      formatFolderLabel(detail.message.folderKind, detail.message.folderName),
    )
  }</div>
        <div class="reader-title-row">
          <h1>${escapeHtml(detail.message.subject || "(无主题)")}</h1>
          ${
    detail.message.webLink
      ? `<a class="reader-action" href="${
        escapeHtml(detail.message.webLink)
      }" target="_blank" rel="noopener noreferrer">在 Outlook 中打开</a>`
      : ""
  }
        </div>
        <div class="reader-byline">
          <strong>${
    escapeHtml(
      detail.message.fromName || detail.message.fromAddress || "未知发件人",
    )
  }</strong>
          ${
    detail.message.fromAddress
      ? `<span>${escapeHtml(detail.message.fromAddress)}</span>`
      : ""
  }
          <span>接收于 ${
    escapeHtml(fmtTime(detail.message.receivedDateTime))
  }</span>
        </div>
      </header>

      ${
    verificationCode
      ? `
          <section class="reader-code-banner">
            <div class="reader-code-row">
              <div class="reader-code-main">
                <div class="reader-code-label">验证码</div>
                <div class="reader-code-value">${
        escapeHtml(verificationCode)
      }</div>
              </div>
              <button
                class="reader-code-copy"
                type="button"
                data-copy-code="${escapeHtml(verificationCode)}"
                data-copy-default="复制验证码"
              >
                复制验证码
              </button>
            </div>
            <div class="reader-code-hint">已从主题或正文中提取，下面仍保留完整正文方便继续核对上下文。</div>
          </section>
        `
      : ""
  }

      ${bodyBlock}

      ${
    attachments.length > 0
      ? `
          <section class="reader-section">
            <div class="reader-section-title">附件</div>
            <ul class="attachment-list">
              ${
        attachments.map((attachment) =>
          `<li>
                  <strong>${escapeHtml(attachment.name)}</strong>
                  <span>${
            attachment.contentType
              ? escapeHtml(attachment.contentType)
              : "未知类型"
          }</span>
                  <span>${
            attachment.size
              ? `${Math.max(1, Math.round(attachment.size / 1024))} KB`
              : "-"
          }</span>
                </li>`
        ).join("")
      }
            </ul>
          </section>
        `
      : ""
  }
    </article>
  `;
}

function renderMessageBody(
  detail: WebMessageDetail | null,
  state: WebConsoleState,
): string {
  if (!detail) {
    return renderReaderIntro(state);
  }
  return renderReaderDetail(detail);
}

function renderMailboxItem(
  mailbox: MailboxBundle,
  selectedMailboxId: string | undefined,
  selectedFolder: "inbox" | "junk",
): string {
  const active = mailbox.connection.mailboxId === selectedMailboxId;
  return `
    <a class="mailbox-item${active ? " is-active" : ""}" title="${
    escapeHtml(mailbox.connection.emailAddress)
  }" href="${
    appHref({
      mailboxId: mailbox.connection.mailboxId,
      folder: selectedFolder,
    })
  }">
      <div class="mailbox-title">${
    escapeHtml(
      mailbox.connection.displayName || mailbox.connection.emailAddress,
    )
  }</div>
      <div class="mailbox-subtitle">${
    escapeHtml(mailbox.connection.emailAddress)
  }</div>
      <div class="mailbox-meta">
        <span>${escapeHtml(providerLabel(mailbox))}</span>
        <span>${
    escapeHtml(
      mailbox.route?.slackChannelName || mailbox.route?.slackChannelId ||
        "未配置",
    )
  }</span>
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
          <strong>${
    escapeHtml(
      state.selectedMailbox.connection.displayName ||
        state.selectedMailbox.connection.emailAddress,
    )
  }</strong>
          <span>${state.mailboxes.length} 个账号 · ${
    escapeHtml(state.selectedMailbox.connection.emailAddress)
  }</span>
        </span>
        <span class="mailbox-switcher-caret" aria-hidden="true">▾</span>
      </summary>
      <div class="mailbox-switcher-menu">
        ${
    state.mailboxes.map((mailbox) => `
          <a
            class="mailbox-switcher-option${
      mailbox.connection.mailboxId ===
          state.selectedMailbox?.connection.mailboxId
        ? " is-active"
        : ""
    }"
            href="${
      appHref({
        mailboxId: mailbox.connection.mailboxId,
        folder: state.selectedFolder,
      })
    }"
          >
            <div class="mailbox-switcher-option-title">${
      escapeHtml(
        mailbox.connection.displayName || mailbox.connection.emailAddress,
      )
    }</div>
            <div class="mailbox-switcher-option-subtitle">${
      escapeHtml(mailbox.connection.emailAddress)
    }</div>
            <div class="mailbox-switcher-option-meta">
              <span>${escapeHtml(providerLabel(mailbox))}</span>
              <span>${
      escapeHtml(
        mailbox.route?.slackChannelName || mailbox.route?.slackChannelId ||
          "未配置",
      )
    }</span>
            </div>
          </a>
        `).join("")
  }
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
  const verificationCode = detectVerificationCode(
    `${message.subject ?? ""}\n${message.bodyPreview ?? ""}`,
  );
  const searchText = compactText(
    [
      message.subject,
      message.fromName,
      message.fromAddress,
      message.bodyPreview,
      verificationCode ?? "",
    ].filter(Boolean).join(" "),
  );
  const mailboxId = state.selectedMailbox?.connection.mailboxId;
  const folder = state.selectedFolder;
  return `
    <a
      class="message-item${active ? " is-active" : ""}"
      href="${
    appHref({
      mailboxId,
      folder,
      messageId: message.messageId,
      pageCursor: state.currentPageCursor,
      page: state.pageIndex > 1 ? state.pageIndex : undefined,
    })
  }"
      data-message-id="${escapeHtml(message.messageId)}"
      data-mailbox-id="${escapeHtml(mailboxId ?? "")}"
      data-folder="${escapeHtml(folder)}"
      data-search-text="${escapeHtml(searchText)}"
      data-has-code="${verificationCode ? "true" : "false"}"
      data-has-attachments="${message.hasAttachments ? "true" : "false"}"
      data-fragment-url="${
    escapeHtml(
      readerFragmentHref({
        mailboxId,
        folder,
        messageId: message.messageId,
      }),
    )
  }"
      ${active ? 'aria-current="true"' : ""}
    >
      <div class="message-row-top">
        <span class="message-subject">${
    escapeHtml(message.subject || "(无主题)")
  }</span>
        <span class="message-time">${
    escapeHtml(fmtListTime(message.receivedDateTime))
  }</span>
      </div>
      <div class="message-row-bottom">
        <div class="message-sender">${
    escapeHtml(message.fromName || message.fromAddress || "未知发件人")
  }</div>
        <div class="message-tags">
          ${
    verificationCode
      ? `<span class="message-chip is-code">${
        escapeHtml(verificationCode)
      }</span>`
      : ""
  }
          ${
    message.hasAttachments ? `<span class="message-chip">附件</span>` : ""
  }
        </div>
      </div>
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
        <strong>${
    state.pageIndex === 1 ? "最新邮件" : `第 ${state.pageIndex} 页`
  }</strong>
      </div>
      <div class="stream-pagination-actions">
        ${
    state.hasPreviousPage
      ? `<a class="stream-page-link" href="${latestHref}">回到最新</a>`
      : ""
  }
        ${
    olderHref
      ? `<a class="stream-page-link is-primary" href="${olderHref}">更早邮件</a>`
      : `<span class="stream-page-link is-disabled">没有更早邮件了</span>`
  }
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

function renderReaderInteractionScript(): string {
  return `
    <script>
      (() => {
        const messageList = document.querySelector('[data-message-list]');
        const readerPane = document.querySelector('[data-reader-pane]');
        const readerWrap = document.querySelector('[data-reader-wrap]');
        if (!messageList || !readerPane || !readerWrap || typeof window.fetch !== 'function' || !window.history || typeof window.history.pushState !== 'function') {
          return;
        }

        const searchInput = document.querySelector('[data-message-search]');
        const filterButtons = Array.from(document.querySelectorAll('[data-message-filter]'));
        const streamCount = document.querySelector('[data-stream-count]');
        const filterEmpty = document.querySelector('[data-filter-empty]');
        const cache = new Map();
        const inflight = new Map();
        let activeController = null;
        let activeFragmentUrl = '';
        let activeFilter = 'all';

        const listItems = () => Array.from(messageList.querySelectorAll('.message-item'));
        const visibleItems = () => listItems().filter((item) => !item.hidden);

        function setLoading(loading) {
          readerPane.classList.toggle('is-loading', loading);
          readerPane.setAttribute('aria-busy', loading ? 'true' : 'false');
        }

        function setActive(item) {
          for (const node of listItems()) {
            const isActive = node === item;
            node.classList.toggle('is-active', isActive);
            if (isActive) node.setAttribute('aria-current', 'true');
            else node.removeAttribute('aria-current');
          }
        }

        function itemByHref(href) {
          return listItems().find((item) => item.href === href) || null;
        }

        function normalizeText(value) {
          return String(value || '').toLowerCase().replace(/\\s+/g, ' ').trim();
        }

        function updateCount() {
          if (!streamCount) return;
          const total = listItems().length;
          const visible = visibleItems().length;
          streamCount.textContent = visible === total ? total + ' 封' : visible + ' / ' + total + ' 封';
        }

        function updateFilterButtons() {
          for (const button of filterButtons) {
            const active = button.dataset.messageFilter === activeFilter;
            button.classList.toggle('is-active', active);
            button.setAttribute('aria-pressed', active ? 'true' : 'false');
          }
        }

        function applyListFilters() {
          const query = normalizeText(searchInput ? searchInput.value : '');
          let visibleCount = 0;

          for (const item of listItems()) {
            const haystack = normalizeText(item.dataset.searchText);
            const matchesQuery = !query || haystack.includes(query);
            const matchesFilter =
              activeFilter === 'all' ||
              (activeFilter === 'code' && item.dataset.hasCode === 'true') ||
              (activeFilter === 'attachments' && item.dataset.hasAttachments === 'true');
            const matched = matchesQuery && matchesFilter;

            item.hidden = !matched;
            item.classList.toggle('is-hidden', !matched);
            if (matched) visibleCount += 1;
          }

          if (filterEmpty) {
            filterEmpty.hidden = visibleCount > 0;
          }

          updateFilterButtons();
          updateCount();
        }

        function currentVisibleItem() {
          return messageList.querySelector('.message-item.is-active:not([hidden])') || visibleItems()[0] || null;
        }

        function trimCache() {
          while (cache.size > 18) {
            const oldestKey = cache.keys().next().value;
            if (!oldestKey) break;
            cache.delete(oldestKey);
          }
        }

        function fetchFragment(url, signal) {
          if (cache.has(url)) return Promise.resolve(cache.get(url));
          if (inflight.has(url)) return inflight.get(url);

          const request = fetch(url, {
            credentials: 'same-origin',
            headers: { 'x-requested-with': 'mail-console' },
            signal,
          }).then((response) => {
            if (!response.ok) throw new Error('HTTP ' + response.status);
            return response.text();
          }).then((html) => {
            cache.set(url, html);
            trimCache();
            inflight.delete(url);
            return html;
          }).catch((error) => {
            inflight.delete(url);
            throw error;
          });

          inflight.set(url, request);
          return request;
        }

        function applyReader(html) {
          readerWrap.innerHTML = html;
          if (typeof readerPane.scrollTo === 'function') {
            readerPane.scrollTo({ top: 0, behavior: 'auto' });
          } else {
            readerPane.scrollTop = 0;
          }
        }

        function schedulePrefetch(item) {
          if (!item) return;
          const url = item.dataset.fragmentUrl;
          if (!url || cache.has(url) || inflight.has(url)) return;

          const run = () => fetchFragment(url).catch(() => {});
          if ('requestIdleCallback' in window) {
            window.requestIdleCallback(run, { timeout: 900 });
          } else {
            window.setTimeout(run, 120);
          }
        }

        function scheduleNeighborPrefetch(item) {
          if (!item) return;
          const candidates = visibleItems();
          const index = candidates.indexOf(item);
          if (index === -1) return;
          for (const candidate of [candidates[index - 1], candidates[index + 1]]) {
            if (candidate) schedulePrefetch(candidate);
          }
        }

        async function activateItem(item, pushHistory) {
          if (!item || item.hidden) return;
          const fragmentUrl = item.dataset.fragmentUrl;
          if (!fragmentUrl) {
            window.location.href = item.href;
            return;
          }
          if (item.classList.contains('is-active') && !activeFragmentUrl) return;
          if (activeFragmentUrl === fragmentUrl) return;

          if (activeController) activeController.abort();
          activeController = new AbortController();
          activeFragmentUrl = fragmentUrl;

          setActive(item);
          setLoading(true);

          try {
            const html = await fetchFragment(fragmentUrl, activeController.signal);
            if (activeFragmentUrl !== fragmentUrl) return;
            applyReader(html);
            if (pushHistory) {
              window.history.pushState({ href: item.href, fragmentUrl }, '', item.href);
            } else {
              window.history.replaceState({ href: item.href, fragmentUrl }, '', item.href);
            }
            scheduleNeighborPrefetch(item);
          } catch (error) {
            if (error && error.name === 'AbortError') return;
            window.location.href = item.href;
          } finally {
            if (activeFragmentUrl === fragmentUrl) {
              activeFragmentUrl = '';
              setLoading(false);
            }
          }
        }

        function shouldIgnoreShortcut(target) {
          if (!(target instanceof Element)) return false;
          if (target.closest('input, textarea, select, button, [contenteditable="true"]')) return true;
          const linkedTarget = target.closest('a[href]');
          return Boolean(linkedTarget && !linkedTarget.classList.contains('message-item'));
        }

        function copyText(text) {
          if (!text) return Promise.reject(new Error('Missing text'));
          if (navigator.clipboard && window.isSecureContext) {
            return navigator.clipboard.writeText(text);
          }
          return new Promise((resolve, reject) => {
            const textarea = document.createElement('textarea');
            textarea.value = text;
            textarea.setAttribute('readonly', '');
            textarea.style.position = 'fixed';
            textarea.style.top = '-9999px';
            textarea.style.opacity = '0';
            document.body.appendChild(textarea);
            textarea.select();
            try {
              document.execCommand('copy');
              resolve();
            } catch (error) {
              reject(error);
            } finally {
              document.body.removeChild(textarea);
            }
          });
        }

        function updateCopyButtonState(button, label, copied) {
          if (!button) return;
          const fallbackLabel = button.dataset.copyDefault || '复制验证码';
          if (button.__copyStateTimer) window.clearTimeout(button.__copyStateTimer);
          button.textContent = label;
          button.classList.toggle('is-copied', Boolean(copied));
          button.__copyStateTimer = window.setTimeout(() => {
            button.textContent = fallbackLabel;
            button.classList.remove('is-copied');
          }, copied ? 1400 : 1800);
        }

        async function handleCopyCode(button) {
          const code = button && button.dataset ? button.dataset.copyCode : '';
          if (!code) return;
          try {
            await copyText(code);
            updateCopyButtonState(button, '已复制', true);
          } catch (_error) {
            updateCopyButtonState(button, '复制失败', false);
          }
        }

        const initialItem = messageList.querySelector('.message-item.is-active');
        if (initialItem && initialItem.dataset.fragmentUrl) {
          window.history.replaceState(
            { href: initialItem.href, fragmentUrl: initialItem.dataset.fragmentUrl },
            '',
            initialItem.href,
          );
          schedulePrefetch(initialItem);
          scheduleNeighborPrefetch(initialItem);
        }

        applyListFilters();

        if (searchInput) {
          searchInput.addEventListener('input', () => {
            applyListFilters();
          });
          searchInput.addEventListener('keydown', (event) => {
            if (event.key === 'Escape' && searchInput.value) {
              event.preventDefault();
              searchInput.value = '';
              applyListFilters();
            }
          });
        }

        for (const button of filterButtons) {
          button.addEventListener('click', () => {
            activeFilter = button.dataset.messageFilter || 'all';
            applyListFilters();
          });
        }

        messageList.addEventListener('click', (event) => {
          if (event.defaultPrevented) return;
          if (!(event.target instanceof Element)) return;
          const item = event.target.closest('.message-item');
          if (!item) return;
          if (event.button !== 0 || event.metaKey || event.ctrlKey || event.shiftKey || event.altKey) return;
          event.preventDefault();
          activateItem(item, true);
        }, true);

        messageList.addEventListener('pointerenter', (event) => {
          if (!(event.target instanceof Element)) return;
          const item = event.target.closest('.message-item');
          if (!item) return;
          schedulePrefetch(item);
        }, true);

        messageList.addEventListener('focusin', (event) => {
          if (!(event.target instanceof Element)) return;
          const item = event.target.closest('.message-item');
          if (!item) return;
          schedulePrefetch(item);
        });

        document.addEventListener('click', (event) => {
          if (!(event.target instanceof Element)) return;
          const button = event.target.closest('[data-copy-code]');
          if (!button) return;
          event.preventDefault();
          handleCopyCode(button);
        });

        window.addEventListener('keydown', (event) => {
          if (event.defaultPrevented || event.metaKey || event.ctrlKey || event.altKey) return;
          if (shouldIgnoreShortcut(event.target)) return;

          const key = event.key;
          const lowerKey = key.toLowerCase();

          if (key === '/') {
            if (!searchInput) return;
            event.preventDefault();
            searchInput.focus();
            searchInput.select();
            return;
          }

          if (lowerKey === 'y') {
            const copyButton = readerWrap.querySelector('[data-copy-code]');
            if (!copyButton) return;
            event.preventDefault();
            handleCopyCode(copyButton);
            return;
          }

          if (lowerKey === 'j') {
            const items = visibleItems();
            if (!items.length) return;
            event.preventDefault();
            const current = currentVisibleItem();
            const index = current ? items.indexOf(current) : -1;
            const next = items[Math.min(items.length - 1, Math.max(0, index + 1))];
            if (next) {
              next.focus({ preventScroll: true });
              activateItem(next, true);
            }
            return;
          }

          if (lowerKey === 'k') {
            const items = visibleItems();
            if (!items.length) return;
            event.preventDefault();
            const current = currentVisibleItem();
            const index = current ? items.indexOf(current) : items.length;
            const previous = items[Math.max(0, index - 1)];
            if (previous) {
              previous.focus({ preventScroll: true });
              activateItem(previous, true);
            }
            return;
          }

          if (lowerKey === 'o' || key === 'Enter') {
            const current = currentVisibleItem();
            if (!current) return;
            event.preventDefault();
            current.focus({ preventScroll: true });
            activateItem(current, true);
          }
        });

        window.addEventListener('popstate', () => {
          const item = itemByHref(window.location.href);
          if (!item) {
            window.location.reload();
            return;
          }
          activateItem(item, false);
        });
      })();
    </script>
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
        height: 100%;
        background:
          radial-gradient(circle at top left, rgba(74, 116, 196, 0.14), transparent 24%),
          linear-gradient(180deg, #07101a 0%, var(--bg) 100%);
        color: var(--text);
        font-family: Inter, ui-sans-serif, system-ui, -apple-system, BlinkMacSystemFont, sans-serif;
        overflow: hidden;
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
        height: 100vh;
        min-height: 100vh;
        display: grid;
        grid-template-rows: 68px 1fr;
        overflow: hidden;
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
        height: calc(100vh - 68px);
        min-height: 0;
        display: grid;
        grid-template-columns: minmax(244px, 292px) minmax(0, 1fr);
        grid-template-areas: "stream reader";
        overflow: hidden;
      }
      .pane {
        min-height: 0;
        height: 100%;
        overflow-y: auto;
        overflow-x: hidden;
        scrollbar-gutter: stable both-edges;
        contain: content;
        overscroll-behavior: contain;
        scrollbar-width: thin;
        scrollbar-color: rgba(122, 174, 255, 0.42) transparent;
      }
      .pane::-webkit-scrollbar {
        width: 10px;
        height: 10px;
      }
      .pane::-webkit-scrollbar-track {
        background: transparent;
      }
      .pane::-webkit-scrollbar-thumb {
        border-radius: 999px;
        background: rgba(122, 174, 255, 0.28);
        border: 2px solid transparent;
        background-clip: padding-box;
      }
      .pane + .pane { border-left: 1px solid var(--line); }
      .stream-pane {
        grid-area: stream;
        padding: 0 0 12px;
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
        display: block;
        position: sticky;
        top: 0;
        z-index: 4;
        padding: 10px 12px 8px;
      }
      .stream-head {
        background: linear-gradient(180deg, rgba(13, 21, 33, 0.99) 0%, rgba(13, 21, 33, 0.92) 72%, rgba(13, 21, 33, 0) 100%);
      }
      .empty-note h3,
      .login-panel h1,
      .reader-intro h1,
      .reader-title-row h1 {
        margin: 0;
      }
      .empty-note p,
      .login-panel p,
      .reader-intro p {
        margin: 0;
        color: var(--muted);
        line-height: 1.6;
      }
      .stream-toolbar {
        display: flex;
        align-items: center;
        justify-content: space-between;
        gap: 10px;
      }
      .stream-tools {
        display: grid;
        gap: 8px;
        margin-top: 10px;
      }
      .stream-search {
        display: flex;
      }
      .stream-search-input {
        width: 100%;
        min-height: 38px;
        padding: 0 14px;
        border: 1px solid var(--line);
        border-radius: 14px;
        background: rgba(255, 255, 255, 0.03);
        color: var(--text);
        outline: none;
      }
      .stream-search-input::placeholder {
        color: var(--muted);
      }
      .stream-search-input:focus {
        border-color: rgba(122, 174, 255, 0.44);
        box-shadow: 0 0 0 3px rgba(122, 174, 255, 0.14);
      }
      .stream-filter-row {
        display: flex;
        align-items: center;
        justify-content: space-between;
        gap: 10px;
        flex-wrap: wrap;
      }
      .stream-filter-group {
        display: inline-flex;
        align-items: center;
        gap: 6px;
        flex-wrap: wrap;
      }
      .stream-filter {
        display: inline-flex;
        align-items: center;
        justify-content: center;
        min-height: 30px;
        padding: 0 10px;
        border: 1px solid var(--line);
        border-radius: 999px;
        background: rgba(255, 255, 255, 0.02);
        color: var(--muted);
        font-size: 12px;
        cursor: pointer;
      }
      .stream-filter.is-active {
        border-color: rgba(122, 174, 255, 0.34);
        background: rgba(122, 174, 255, 0.12);
        color: var(--text);
      }
      .stream-keyhint {
        font-size: 11px;
        color: var(--muted);
        white-space: nowrap;
      }
      .message-list {
        display: grid;
        gap: 0;
        content-visibility: auto;
        contain-intrinsic-size: 560px;
      }
      .message-item[hidden],
      .message-item.is-hidden {
        display: none !important;
      }
      .message-item {
        position: relative;
        display: grid;
        gap: 6px;
        transition: background-color 120ms ease, color 120ms ease;
        content-visibility: auto;
        contain-intrinsic-size: 64px;
      }
      .message-item {
        padding: 10px 12px 10px 14px;
        border-bottom: 1px solid var(--line);
      }
      .message-item::before {
        content: "";
        position: absolute;
        left: 0;
        top: 10px;
        bottom: 10px;
        width: 3px;
        border-radius: 999px;
        background: transparent;
      }
      .message-item:hover {
        background: rgba(255, 255, 255, 0.03);
      }
      .message-item.is-active {
        background: linear-gradient(90deg, var(--accent-soft) 0%, rgba(122, 174, 255, 0.03) 100%);
      }
      .message-item.is-active::before { background: var(--accent); }
      .message-subject {
        display: -webkit-box;
        -webkit-box-orient: vertical;
        -webkit-line-clamp: 2;
        overflow: hidden;
        font-size: 14px;
        font-weight: 650;
        line-height: 1.4;
      }
      .message-sender,
      .message-time {
        font-size: 12px;
        color: var(--muted);
      }
      .message-sender {
        min-width: 0;
        line-height: 1.4;
        white-space: nowrap;
        overflow: hidden;
        text-overflow: ellipsis;
      }
      .message-row-top {
        display: grid;
        grid-template-columns: minmax(0, 1fr) auto;
        align-items: start;
        gap: 8px;
      }
      .message-row-bottom {
        display: flex;
        align-items: center;
        justify-content: space-between;
        gap: 8px;
      }
      .message-time {
        flex: 0 0 auto;
        white-space: nowrap;
        font-variant-numeric: tabular-nums;
      }
      .message-tags {
        display: flex;
        align-items: center;
        justify-content: flex-end;
        gap: 4px;
        flex-wrap: wrap;
      }
      .message-chip {
        width: fit-content;
        padding: 3px 6px;
        border-radius: 999px;
        border: 1px solid rgba(148, 163, 184, 0.16);
        background: rgba(255, 255, 255, 0.04);
        font-size: 10px;
        letter-spacing: 0.06em;
        text-transform: uppercase;
        color: var(--muted);
      }
      .message-chip.is-code {
        border-color: rgba(122, 174, 255, 0.32);
        background: rgba(122, 174, 255, 0.1);
        color: var(--accent);
        font-family: ui-monospace, SFMono-Regular, Menlo, Monaco, monospace;
        font-weight: 700;
      }
      .folder-tabs {
        display: inline-flex;
        align-items: center;
        gap: 4px;
        padding: 4px;
        border: 1px solid var(--line);
        border-radius: 999px;
        background: rgba(255, 255, 255, 0.02);
      }
      .stream-count {
        flex: 0 0 auto;
        font-size: 12px;
        color: var(--muted);
        white-space: nowrap;
      }
      .stream-count strong {
        color: var(--text);
      }
      .folder-tabs {
        align-items: center;
      }
      .folder-tab {
        position: relative;
        padding: 7px 12px;
        border-radius: 999px;
        font-size: 12px;
        color: var(--muted);
      }
      .folder-tab.is-active {
        color: var(--text);
        background: rgba(122, 174, 255, 0.12);
      }
      .folder-tab.is-active::after {
        display: none;
      }
      .stream-alert {
        margin: 0 12px 10px;
        padding: 10px 12px;
        border-radius: 14px;
        border: 1px solid rgba(220, 38, 38, 0.24);
        background: rgba(220, 38, 38, 0.08);
        color: #fecaca;
        font-size: 13px;
      }
      .stream-filter-empty {
        margin: 8px 12px 0;
        padding: 12px 14px;
        border: 1px dashed var(--line);
        border-radius: 14px;
        color: var(--muted);
        font-size: 12px;
        line-height: 1.6;
      }
      .stream-pagination {
        display: flex;
        align-items: center;
        justify-content: space-between;
        gap: 16px;
        margin: 10px 12px 0;
        padding-top: 14px;
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
        min-height: 100%;
        padding: 24px 28px 32px;
        display: flex;
        justify-content: center;
        align-items: flex-start;
        transition: opacity 120ms ease;
      }
      .reader-pane.is-loading {
        cursor: progress;
      }
      .reader-pane.is-loading .reader-wrap {
        opacity: 0.72;
      }
      .reader-intro,
      .reader-document {
        display: grid;
        gap: 20px;
        width: min(1240px, 100%);
        padding: 30px 34px 36px;
        background: linear-gradient(180deg, var(--paper) 0%, var(--paper-soft) 100%);
        border-radius: 24px;
        box-shadow: 0 12px 34px rgba(15, 23, 42, 0.08);
        position: relative;
        overflow: hidden;
      }
      .reader-intro::before,
      .reader-document::before {
        content: "";
        position: absolute;
        inset: 0 0 auto 0;
        height: 88px;
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
        font-size: clamp(30px, 3.4vw, 42px);
        line-height: 1.08;
        letter-spacing: -0.035em;
        color: var(--reader-text);
        max-width: 18ch;
      }
      .reader-intro p,
      .reader-byline {
        font-size: 15px;
        color: var(--reader-muted);
        line-height: 1.6;
        max-width: 80ch;
      }
      .reader-inline-meta {
        display: grid;
        grid-template-columns: repeat(4, minmax(0, 1fr));
        border-top: 1px solid var(--reader-line);
        border-bottom: 1px solid var(--reader-line);
        background: rgba(250, 252, 255, 0.8);
      }
      .reader-inline-meta > div {
        display: grid;
        gap: 6px;
        padding: 14px 0;
      }
      .reader-inline-meta > div + div {
        padding-left: 18px;
        margin-left: 18px;
        border-left: 1px solid var(--reader-line);
      }
      .reader-inline-meta span,
      .reader-meta-chip span {
        font-size: 11px;
        letter-spacing: 0.1em;
        text-transform: uppercase;
        color: var(--reader-muted);
      }
      .reader-inline-meta strong,
      .reader-meta-chip strong {
        font-size: 14px;
        font-weight: 600;
        color: var(--reader-text);
      }
      .reader-header {
        display: grid;
        gap: 14px;
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
        gap: 10px 14px;
      }
      .reader-byline strong {
        color: var(--reader-text);
      }
      .reader-code-banner {
        display: grid;
        gap: 8px;
        padding: 18px 20px;
        border-radius: 20px;
        border: 1px solid rgba(122, 174, 255, 0.24);
        background: linear-gradient(135deg, rgba(122, 174, 255, 0.16) 0%, rgba(122, 174, 255, 0.04) 100%);
      }
      .reader-code-row {
        display: flex;
        align-items: flex-start;
        justify-content: space-between;
        gap: 16px;
      }
      .reader-code-main {
        min-width: 0;
        display: grid;
        gap: 8px;
      }
      .reader-code-label {
        font-size: 11px;
        font-weight: 700;
        letter-spacing: 0.14em;
        text-transform: uppercase;
        color: #466a96;
      }
      .reader-code-value {
        font-size: clamp(28px, 4.6vw, 44px);
        font-weight: 800;
        line-height: 1;
        letter-spacing: 0.16em;
        color: #0f172a;
        font-family: ui-monospace, SFMono-Regular, Menlo, Monaco, monospace;
      }
      .reader-code-copy {
        flex: 0 0 auto;
        display: inline-flex;
        align-items: center;
        justify-content: center;
        min-height: 40px;
        padding: 0 16px;
        border: 1px solid rgba(15, 23, 42, 0.14);
        border-radius: 999px;
        background: rgba(255, 255, 255, 0.74);
        color: #0f172a;
        font-size: 13px;
        font-weight: 700;
        cursor: pointer;
      }
      .reader-code-copy.is-copied {
        border-color: rgba(37, 99, 235, 0.26);
        background: rgba(37, 99, 235, 0.12);
        color: #0f172a;
      }
      .reader-code-hint {
        font-size: 13px;
        line-height: 1.6;
        color: var(--reader-muted);
      }
      .reader-meta-bar {
        display: flex;
        flex-wrap: wrap;
        gap: 12px;
        padding-bottom: 4px;
      }
      .reader-meta-chip {
        display: grid;
        gap: 6px;
        min-width: 0;
        padding: 10px 14px;
        border-radius: 999px;
        border: 1px solid var(--reader-line);
        background: rgba(255, 255, 255, 0.74);
      }
      .reader-section {
        display: block;
      }
      .reader-section + .reader-section {
        margin-top: 18px;
      }
      .reader-section-body {
        margin-top: 2px;
      }
      .reader-section-body .reader-section-title {
        margin-bottom: 12px;
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
        min-height: clamp(720px, calc(100vh - 260px), 1200px);
        border: 0;
        border-radius: 22px;
        background: #ffffff;
        box-shadow: inset 0 0 0 1px var(--reader-line);
      }
      .mail-body-text {
        margin: 0;
        padding: 24px 26px;
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
        .workspace { grid-template-columns: minmax(232px, 272px) minmax(0, 1fr); }
        .reader-wrap { padding: 22px 24px 28px; }
        .reader-intro,
        .reader-document {
          width: min(1120px, 100%);
          padding: 28px 30px 34px;
        }
        .reader-intro h1,
        .reader-title-row h1 {
          font-size: clamp(28px, 3.2vw, 40px);
        }
      }
      @media (max-width: 960px) {
        html, body {
          height: auto;
          overflow: auto;
        }
        .app-shell {
          height: auto;
          overflow: visible;
          grid-template-rows: auto 1fr;
        }
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
          height: auto;
          overflow: visible;
          grid-template-columns: 1fr;
          grid-template-areas:
            "stream"
            "reader";
        }
        .pane {
          height: auto;
          min-height: auto;
          overflow: visible;
        }
        .pane + .pane { border-left: 0; border-top: 1px solid var(--line); }
        .stream-head {
          padding: 10px 14px 8px;
        }
        .stream-toolbar {
          flex-wrap: wrap;
        }
        .stream-filter-row {
          align-items: flex-start;
        }
        .stream-keyhint {
          white-space: normal;
        }
        .stream-count {
          font-size: 11px;
        }
        .reader-wrap {
          min-height: auto;
          padding: 20px 16px 28px;
        }
        .reader-intro,
        .reader-document {
          padding: 30px 22px 34px;
          border-radius: 24px;
        }
        .reader-inline-meta { grid-template-columns: repeat(2, minmax(0, 1fr)); }
        .reader-inline-meta > div + div {
          padding-left: 0;
          margin-left: 0;
          border-left: 0;
        }
        .reader-inline-meta > div:nth-child(odd) {
          padding-right: 14px;
          border-right: 1px solid var(--reader-line);
        }
        .reader-inline-meta > div:nth-child(n + 3) {
          padding-left: 0;
          border-top: 1px solid var(--reader-line);
        }
        .reader-meta-bar {
          display: grid;
          grid-template-columns: repeat(2, minmax(0, 1fr));
          gap: 0;
          border-top: 1px solid var(--reader-line);
          border-bottom: 1px solid var(--reader-line);
          background: rgba(250, 252, 255, 0.8);
        }
        .reader-meta-chip {
          padding: 14px 0;
          border: 0;
          border-radius: 0;
          background: transparent;
        }
        .reader-meta-chip:nth-child(odd) {
          padding-right: 14px;
          border-right: 1px solid var(--reader-line);
        }
        .reader-meta-chip:nth-child(n + 3) {
          border-top: 1px solid var(--reader-line);
        }
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
        .stream-toolbar {
          align-items: stretch;
          flex-direction: column;
        }
        .stream-filter-row,
        .reader-code-row {
          align-items: stretch;
          flex-direction: column;
        }
        .stream-filter-group {
          width: 100%;
        }
        .folder-tabs {
          width: fit-content;
        }
        .stream-count {
          align-self: flex-start;
        }
        .reader-wrap { padding: 18px 16px 28px; }
        .reader-intro,
        .reader-document {
          padding: 28px 18px 32px;
          border-radius: 24px;
          gap: 20px;
        }
        .message-row-bottom {
          align-items: flex-start;
          flex-direction: column;
        }
        .message-tags {
          justify-content: flex-start;
        }
        .reader-title-row {
          flex-direction: column;
          gap: 16px;
        }
        .reader-inline-meta,
        .reader-meta-bar { grid-template-columns: 1fr; }
        .reader-inline-meta > div,
        .reader-meta-chip {
          padding-right: 0 !important;
          border-right: 0 !important;
        }
        .reader-inline-meta > div + div,
        .reader-meta-chip + .reader-meta-chip {
          padding-left: 0;
          margin-left: 0;
          border-left: 0;
        }
        .reader-inline-meta > div:nth-child(n + 2),
        .reader-meta-chip:nth-child(n + 2) {
          border-top: 1px solid var(--reader-line);
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
          ${
      input.configured
        ? `
              ${
          input.error
            ? `<div class="login-alert">${escapeHtml(input.error)}</div>`
            : ""
        }
              <form method="POST" action="/app/login">
                <input class="login-input" type="password" name="password" placeholder="输入管理员密码" autocomplete="current-password" required />
                <button class="login-button" type="submit">进入邮件工作台</button>
              </form>
            `
        : `<div class="login-alert">当前未配置 <code>WEB_ADMIN_PASSWORD</code>，Web 控制台尚未启用。</div>`
    }
        </section>
      </div>
    `,
  );
}

export function renderReaderContentFragment(
  detail: WebMessageDetail,
): Response {
  return new Response(renderReaderDetail(detail), {
    headers: { "content-type": "text/html; charset=utf-8" },
  });
}

export function renderAppPage(state: WebConsoleState): Response {
  const selectedFolder = state.selectedFolder;
  const mailboxLabel = state.selectedMailbox?.connection.displayName ||
    state.selectedMailbox?.connection.emailAddress || "No mailbox";
  const hasMessages = state.messages.length > 0;

  return renderAppShell(
    "Mail Console",
    `
      <div class="app-shell" data-app-shell="mail-console">
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
              <strong>${state.messages.length} 封邮件${
      state.pageIndex > 1 ? ` · 第 ${state.pageIndex} 页` : ""
    }</strong>
            </div>
            <form method="POST" action="/app/logout">
              <button class="ghost-button" type="submit">退出</button>
            </form>
          </div>
        </header>

        <div class="workspace">
          <section class="pane stream-pane">
            <div class="stream-head">
              ${
      state.selectedMailbox
        ? `
                  <div class="stream-toolbar">
                    <div class="folder-tabs">
                      <a class="folder-tab${
          selectedFolder === "inbox" ? " is-active" : ""
        }" href="${
          appHref({
            mailboxId: state.selectedMailbox.connection.mailboxId,
            folder: "inbox",
          })
        }">收件箱</a>
                      <a class="folder-tab${
          selectedFolder === "junk" ? " is-active" : ""
        }" href="${
          appHref({
            mailboxId: state.selectedMailbox.connection.mailboxId,
            folder: "junk",
          })
        }">垃圾邮件</a>
                    </div>
                    <div class="stream-count" data-stream-count>${state.messages.length} 封</div>
                  </div>
                  ${
          hasMessages
            ? `
                    <div class="stream-tools">
                      <label class="stream-search">
                        <input
                          class="stream-search-input"
                          type="search"
                          placeholder="搜索当前页主题、发件人或验证码"
                          aria-label="搜索当前页邮件"
                          data-message-search
                        />
                      </label>
                      <div class="stream-filter-row">
                        <div class="stream-filter-group" role="toolbar" aria-label="邮件筛选">
                          <button class="stream-filter is-active" type="button" data-message-filter="all" aria-pressed="true">全部</button>
                          <button class="stream-filter" type="button" data-message-filter="code" aria-pressed="false">验证码</button>
                          <button class="stream-filter" type="button" data-message-filter="attachments" aria-pressed="false">附件</button>
                        </div>
                        <div class="stream-keyhint">/ 搜索 · J/K 切换 · Y 复制验证码</div>
                      </div>
                    </div>
                  `
            : ""
        }
                `
        : `<div class="section-label">消息流</div>`
    }
            </div>
            ${
      state.error
        ? `<div class="stream-alert">${escapeHtml(state.error)}</div>`
        : ""
    }
            ${
      !state.selectedMailbox
        ? renderEmptyMailboxes()
        : state.messages.length > 0
        ? `
                <div class="message-list" data-message-list>${
          state.messages.map((message) => renderMessageItem(state, message))
            .join("")
        }</div>
                <div class="stream-filter-empty" data-filter-empty hidden>
                  当前页没有匹配邮件，可以清空搜索或切到更早分页继续找。
                </div>
                ${renderMessagePagination(state)}
              `
        : `
                <section class="empty-note">
                  <h3>这个文件夹里没有可展示邮件</h3>
                  <p>如果邮箱刚接入，可以先等待同步，或者在 Slack 中执行 <code>/mail sync &lt;mailbox&gt;</code>。</p>
                </section>
              `
    }
          </section>

          <main class="pane reader-pane" data-reader-pane aria-live="polite" aria-busy="false">
            <div class="reader-wrap" data-reader-wrap>
              ${renderMessageBody(state.selectedMessage, state)}
            </div>
          </main>
        </div>
      </div>
      ${renderReaderInteractionScript()}
    `,
  );
}
