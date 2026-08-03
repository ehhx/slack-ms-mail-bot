const WEB_ASSET_VERSION = "20260803-5";

function escapeHtml(input: string): string {
  return input
    .replace(/&/g, "&amp;")
    .replace(/</g, "&lt;")
    .replace(/>/g, "&gt;")
    .replace(/\"/g, "&quot;")
    .replace(/'/g, "&#39;");
}

type IconName =
  | "arrow-left"
  | "chevron-down"
  | "inbox"
  | "log-out"
  | "mail"
  | "refresh"
  | "search"
  | "shield-alert";

// The paths are the matching Lucide icons embedded locally to keep the SPA dependency-free.
function icon(name: IconName): string {
  const paths: Record<IconName, string> = {
    "arrow-left": '<path d="m12 19-7-7 7-7"/><path d="M19 12H5"/>',
    "chevron-down": '<path d="m6 9 6 6 6-6"/>',
    inbox:
      '<polyline points="22 12 16 12 14 15 10 15 8 12 2 12"/><path d="M5.45 5.11 2 12v6a2 2 0 0 0 2 2h16a2 2 0 0 0 2-2v-6l-3.45-6.89A2 2 0 0 0 16.76 4H7.24a2 2 0 0 0-1.79 1.11z"/>',
    "log-out":
      '<path d="M9 21H5a2 2 0 0 1-2-2V5a2 2 0 0 1 2-2h4"/><polyline points="16 17 21 12 16 7"/><line x1="21" x2="9" y1="12" y2="12"/>',
    mail:
      '<rect width="20" height="16" x="2" y="4" rx="2"/><path d="m22 7-8.97 5.7a1.94 1.94 0 0 1-2.06 0L2 7"/>',
    refresh:
      '<path d="M21 12a9 9 0 0 1-15.8 5.9L3 15"/><path d="M3 21v-6h6"/><path d="M3 12A9 9 0 0 1 18.8 6.1L21 9"/><path d="M21 3v6h-6"/>',
    search: '<circle cx="11" cy="11" r="8"/><path d="m21 21-4.3-4.3"/>',
    "shield-alert":
      '<path d="M20 13c0 5-3.5 7.5-7.66 8.95a1 1 0 0 1-.67-.01C7.5 20.5 4 18 4 13V6a1 1 0 0 1 1-1c2 0 4.5-1.2 6.24-2.72a1.17 1.17 0 0 1 1.52 0C14.51 3.81 17 5 19 5a1 1 0 0 1 1 1z"/><path d="M12 8v4"/><path d="M12 16h.01"/>',
  };
  return `<svg class="icon" viewBox="0 0 24 24" fill="none" stroke="currentColor" stroke-width="1.8" stroke-linecap="round" stroke-linejoin="round" aria-hidden="true">${
    paths[name]
  }</svg>`;
}

function htmlDocument(input: {
  title: string;
  body: string;
  appScript?: boolean;
  bodyClass?: string;
}): Response {
  const script = input.appScript
    ? `<script type="module" src="/app/assets/app.js?v=${WEB_ASSET_VERSION}"></script>`
    : "";
  return new Response(
    `<!doctype html>
<html lang="zh-CN">
  <head>
    <meta charset="utf-8" />
    <meta name="viewport" content="width=device-width, initial-scale=1, viewport-fit=cover" />
    <meta name="color-scheme" content="light" />
    <meta name="theme-color" content="#f6f7f9" />
    <title>${escapeHtml(input.title)}</title>
    <link rel="icon" href="data:image/svg+xml,%3Csvg xmlns='http://www.w3.org/2000/svg' viewBox='0 0 32 32'%3E%3Crect width='32' height='32' rx='7' fill='%23146b9f'/%3E%3Cpath d='M8 10h16v12H8z' fill='none' stroke='white' stroke-width='2'/%3E%3Cpath d='m8 11 8 6 8-6' fill='none' stroke='white' stroke-width='2'/%3E%3C/svg%3E" />
    <link rel="stylesheet" href="/app/assets/app.css?v=${WEB_ASSET_VERSION}" />
    ${script}
  </head>
  <body class="${escapeHtml(input.bodyClass ?? "")}">
    ${input.body}
  </body>
</html>`,
    {
      status: 200,
      headers: {
        "content-type": "text/html; charset=utf-8",
        "cache-control": "private, no-store",
      },
    },
  );
}

export function renderLoginPage(input: {
  error?: string;
  configured: boolean;
}): Response {
  return htmlDocument({
    title: "登录 - Outlook 邮件工作台",
    bodyClass: "login-body",
    body: `
      <main class="login-page">
        <section class="login-intro" aria-labelledby="product-title">
          <div class="login-brand">
            <span class="brand-mark brand-mark-large" aria-hidden="true">${
      icon("mail")
    }</span>
            <div>
              <span class="eyebrow">MAIL WORKSPACE</span>
              <strong>Mail</strong>
            </div>
          </div>
          <div class="login-intro-copy">
            <p class="login-kicker">管理员控制台</p>
            <h1 id="product-title">Outlook<br />邮件工作台</h1>
            <p>集中查看已连接邮箱，并快速处理重要邮件。</p>
          </div>
          <p class="login-intro-foot">Microsoft Graph · Lark</p>
        </section>

        <section class="login-form-pane" aria-labelledby="login-title">
          <div class="login-form-wrap">
            <div class="login-form-heading">
              <span class="eyebrow">SECURE ACCESS</span>
              <h2 id="login-title">登录工作台</h2>
              <p>使用管理员密码继续。</p>
            </div>
            ${
      input.configured
        ? `
            ${
          input.error
            ? `<div class="login-alert" role="alert">${
              escapeHtml(input.error)
            }</div>`
            : ""
        }
            <form class="login-form" method="POST" action="/app/login">
              <input class="sr-only" type="text" name="username" value="admin" autocomplete="username" tabindex="-1" aria-hidden="true" />
              <label for="password">管理员密码</label>
              <input id="password" type="password" name="password" placeholder="输入密码" autocomplete="current-password" required autofocus />
              <button type="submit">
                <span>进入工作台</span>
                <span aria-hidden="true">→</span>
              </button>
            </form>
          `
        : `<div class="login-alert" role="alert">当前未配置 <code>WEB_ADMIN_PASSWORD</code>，Web 工作台尚未启用。</div>`
    }
            <div class="login-security-note">
              <span class="status-dot" aria-hidden="true"></span>
              <span>会话通过安全的 HttpOnly Cookie 保存</span>
            </div>
          </div>
        </section>
      </main>
    `,
  });
}

export function renderAppPage(): Response {
  return htmlDocument({
    title: "Outlook 邮件工作台",
    appScript: true,
    bodyClass: "app-body",
    body: `
      <div class="app-shell" id="mailApp" data-mobile-view="list">
        <header class="topbar">
          <div class="brand" aria-label="Outlook 邮件工作台">
            <span class="brand-mark" aria-hidden="true">${icon("mail")}</span>
            <div class="brand-copy">
              <strong>Mail</strong>
              <span>Workspace</span>
            </div>
          </div>

          <div class="account-switcher">
            <button class="account-trigger" id="accountTrigger" type="button" aria-haspopup="listbox" aria-expanded="false">
              <span class="account-avatar" id="accountAvatar">M</span>
              <span class="account-copy">
                <strong id="accountName">正在载入邮箱</strong>
                <span id="accountAddress">请稍候</span>
              </span>
              <span class="chevron" aria-hidden="true">${
      icon("chevron-down")
    }</span>
            </button>
            <div class="account-menu" id="accountMenu" role="listbox" aria-label="切换邮箱" hidden></div>
          </div>

          <label class="global-search" for="mailSearch">
            ${icon("search")}
            <input id="mailSearch" type="search" placeholder="搜索发件人、主题或验证码" autocomplete="off" />
          </label>

          <div class="topbar-actions">
            <button class="icon-button" id="refreshButton" type="button" title="刷新邮件" aria-label="刷新邮件">${
      icon("refresh")
    }</button>
            <form method="POST" action="/app/logout">
              <button class="icon-button" type="submit" title="退出登录" aria-label="退出登录">${
      icon("log-out")
    }</button>
            </form>
          </div>
        </header>

        <div class="workspace">
          <nav class="folder-rail" aria-label="邮件文件夹">
            <div class="folder-nav">
              <button class="folder-button is-active" type="button" data-folder="inbox" aria-label="收件箱" title="收件箱">
                ${icon("inbox")}<small>收件箱</small>
              </button>
              <button class="folder-button" type="button" data-folder="junk" aria-label="垃圾邮件" title="垃圾邮件">
                ${icon("shield-alert")}<small>垃圾邮件</small>
              </button>
            </div>
            <div class="rail-status" title="邮箱连接状态">
              <span class="connection-state" id="connectionState"></span>
              <span>连接</span>
            </div>
          </nav>

          <section class="stream-pane" aria-label="邮件列表">
            <header class="stream-header">
              <div>
                <p class="eyebrow" id="folderEyebrow">INBOX</p>
                <h1 id="folderTitle">收件箱</h1>
              </div>
              <div class="stream-summary">
                <strong class="message-count" id="messageCount">--</strong>
                <span id="mailboxStatusLabel">正在同步状态</span>
              </div>
            </header>
            <div class="filter-bar" role="toolbar" aria-label="邮件筛选">
              <button class="filter-button is-active" type="button" data-filter="all" aria-pressed="true">全部</button>
              <button class="filter-button" type="button" data-filter="code" aria-pressed="false">验证码</button>
              <button class="filter-button" type="button" data-filter="attachments" aria-pressed="false">有附件</button>
            </div>
            <div class="stream-notice" id="streamNotice" role="alert" hidden>
              <span class="notice-copy" id="streamNoticeText"></span>
              <button id="noticeRetry" type="button">重试</button>
            </div>
            <div class="message-stage">
              <div class="message-list" id="messageList" role="listbox" aria-label="邮件"></div>
              <div class="stream-empty" id="streamEmpty" hidden>
                <span class="empty-mark" aria-hidden="true">${
      icon("mail")
    }</span>
                <strong>这里还没有邮件</strong>
                <span>切换文件夹或刷新后再试。</span>
                <button id="emptyAction" type="button">刷新邮件</button>
              </div>
            </div>
            <footer class="pagination" id="pagination" hidden>
              <button type="button" id="previousPage" aria-label="上一页">←</button>
              <span id="pageLabel">第 1 页</span>
              <button type="button" id="nextPage" aria-label="下一页">→</button>
            </footer>
          </section>

          <main class="reader-pane" id="readerPane" aria-live="polite" aria-busy="false">
            <div class="reader-mobile-bar">
              <button class="reader-back" id="readerBack" type="button">${
      icon("arrow-left")
    }<span>邮件列表</span></button>
              <strong id="readerMobileTitle">邮件正文</strong>
            </div>
            <div class="reader-content" id="readerContent">
              <section class="reader-placeholder">
                <span class="placeholder-mark" aria-hidden="true">${
      icon("mail")
    }</span>
                <h2>选择一封邮件</h2>
                <p>邮件正文会在这里打开。</p>
              </section>
            </div>
          </main>
        </div>

        <div class="toast" id="toast" role="status" aria-live="polite" hidden></div>
      </div>
    `,
  });
}
