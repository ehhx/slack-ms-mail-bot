export const WEB_APP_JS = String.raw`
const app = document.getElementById("mailApp");

if (app) {
  const elements = {
    accountTrigger: document.getElementById("accountTrigger"),
    accountMenu: document.getElementById("accountMenu"),
    accountAvatar: document.getElementById("accountAvatar"),
    accountName: document.getElementById("accountName"),
    accountAddress: document.getElementById("accountAddress"),
    search: document.getElementById("mailSearch"),
    refresh: document.getElementById("refreshButton"),
    folderTitle: document.getElementById("folderTitle"),
    folderEyebrow: document.getElementById("folderEyebrow"),
    folderRail: document.querySelector(".folder-rail"),
    streamPane: document.querySelector(".stream-pane"),
    messageCount: document.getElementById("messageCount"),
    mailboxStatusLabel: document.getElementById("mailboxStatusLabel"),
    messageList: document.getElementById("messageList"),
    streamEmpty: document.getElementById("streamEmpty"),
    streamNotice: document.getElementById("streamNotice"),
    streamNoticeText: document.getElementById("streamNoticeText"),
    noticeRetry: document.getElementById("noticeRetry"),
    emptyAction: document.getElementById("emptyAction"),
    pagination: document.getElementById("pagination"),
    previousPage: document.getElementById("previousPage"),
    nextPage: document.getElementById("nextPage"),
    pageLabel: document.getElementById("pageLabel"),
    readerPane: document.getElementById("readerPane"),
    readerContent: document.getElementById("readerContent"),
    readerBack: document.getElementById("readerBack"),
    readerMobileTitle: document.getElementById("readerMobileTitle"),
    connectionState: document.getElementById("connectionState"),
    toast: document.getElementById("toast"),
  };

  const urlState = new URL(window.location.href);
  const state = {
    mailboxes: [],
    mailboxId: urlState.searchParams.get("mailbox") || "",
    folder: urlState.searchParams.get("folder") === "junk" ? "junk" : "inbox",
    messages: [],
    selectedMessageId: urlState.searchParams.get("message") || "",
    selectedDetail: null,
    currentCursor: urlState.searchParams.get("pageCursor") || "",
    nextCursor: "",
    cursorHistory: [],
    page: 1,
    filter: "all",
    query: "",
    listRequest: null,
    detailRequest: null,
  };

  const listCache = new Map();
  const detailCache = new Map();
  let toastTimer = 0;
  let filterFrame = 0;
  let prefetchTimer = 0;

  const LIST_CACHE_TTL = 30_000;
  const DETAIL_CACHE_TTL = 5 * 60_000;
  const MAX_LIST_CACHE = 12;
  const MAX_DETAIL_CACHE = 20;

  const ICON_PATHS = {
    copy: '<rect width="14" height="14" x="8" y="8" rx="2" ry="2"/><path d="M4 16c-1.1 0-2-.9-2-2V4c0-1.1.9-2 2-2h10c1.1 0 2 .9 2 2"/>',
    external: '<path d="M15 3h6v6"/><path d="M10 14 21 3"/><path d="M18 13v6a2 2 0 0 1-2 2H5a2 2 0 0 1-2-2V8a2 2 0 0 1 2-2h6"/>',
    mail: '<rect width="20" height="16" x="2" y="4" rx="2"/><path d="m22 7-8.97 5.7a1.94 1.94 0 0 1-2.06 0L2 7"/>',
    paperclip: '<path d="m16 6-8.41 8.41a2 2 0 0 0 2.83 2.83L18 9.66a4 4 0 0 0-5.66-5.66l-8.48 8.49a6 6 0 0 0 8.48 8.48L20 13"/>',
  };

  function icon(name) {
    return '<svg class="icon" viewBox="0 0 24 24" fill="none" stroke="currentColor" stroke-width="1.8" stroke-linecap="round" stroke-linejoin="round" aria-hidden="true">' + ICON_PATHS[name] + '</svg>';
  }

  function escapeHtml(value) {
    return String(value ?? "")
      .replace(/&/g, "&amp;")
      .replace(/</g, "&lt;")
      .replace(/>/g, "&gt;")
      .replace(/"/g, "&quot;")
      .replace(/'/g, "&#39;");
  }

  function compact(value) {
    return String(value ?? "").replace(/\s+/g, " ").trim();
  }

  function initial(value) {
    const text = compact(value);
    return text ? Array.from(text)[0].toUpperCase() : "M";
  }

  function formatTime(value, compactMode = false) {
    if (!value) return "-";
    const date = new Date(value);
    if (Number.isNaN(date.getTime())) return String(value);
    const now = new Date();
    const sameDay = date.toDateString() === now.toDateString();
    if (compactMode && sameDay) {
      return new Intl.DateTimeFormat("zh-CN", {
        hour: "2-digit",
        minute: "2-digit",
        hour12: false,
      }).format(date);
    }
    return new Intl.DateTimeFormat("zh-CN", compactMode
      ? { month: "numeric", day: "numeric" }
      : { year: "numeric", month: "2-digit", day: "2-digit", hour: "2-digit", minute: "2-digit", hour12: false }
    ).format(date);
  }

  function formatBytes(value) {
    const size = Number(value || 0);
    if (!Number.isFinite(size) || size <= 0) return "";
    if (size < 1024) return size + " B";
    if (size < 1024 * 1024) return Math.round(size / 1024) + " KB";
    return (size / 1024 / 1024).toFixed(1) + " MB";
  }

  function mailboxStatusText(mailbox) {
    if (!mailbox) return "等待连接";
    if (mailbox.status !== "active") return "连接需要处理";
    const lastSyncAt = mailbox.syncState && mailbox.syncState.lastSyncAt;
    return lastSyncAt ? "同步于 " + formatTime(lastSyncAt, true) : "连接正常";
  }

  function trimCache(cache, maxSize) {
    while (cache.size > maxSize) {
      const oldest = cache.keys().next().value;
      if (oldest === undefined) break;
      cache.delete(oldest);
    }
  }

  function getCached(cache, key, ttl) {
    const item = cache.get(key);
    if (!item) return null;
    if (Date.now() - item.createdAt > ttl) {
      cache.delete(key);
      return null;
    }
    cache.delete(key);
    cache.set(key, item);
    return item.value;
  }

  function setCached(cache, key, value, maxSize) {
    cache.delete(key);
    cache.set(key, { createdAt: Date.now(), value });
    trimCache(cache, maxSize);
  }

  async function fetchJson(url, init = {}) {
    const response = await fetch(url, {
      credentials: "same-origin",
      headers: { accept: "application/json", ...(init.headers || {}) },
      ...init,
    });
    if (response.status === 401) {
      window.location.assign("/app/login");
      throw new Error("登录已过期");
    }
    const data = await response.json().catch(() => ({}));
    if (!response.ok) {
      throw new Error(data.error || "请求失败 (HTTP " + response.status + ")");
    }
    return data;
  }

  function currentMailbox() {
    return state.mailboxes.find((item) => item.mailboxId === state.mailboxId) || null;
  }

  function folderLabel() {
    return state.folder === "junk" ? "垃圾邮件" : "收件箱";
  }

  function setMobileView(view) {
    const mobile = window.matchMedia("(max-width: 760px)").matches;
    const readerVisible = mobile && view === "reader";
    app.dataset.mobileView = readerVisible ? "reader" : "list";
    elements.folderRail.inert = readerVisible;
    elements.streamPane.inert = readerVisible;
    elements.readerPane.inert = mobile && !readerVisible;
    elements.folderRail.setAttribute("aria-hidden", String(readerVisible));
    elements.streamPane.setAttribute("aria-hidden", String(readerVisible));
    elements.readerPane.setAttribute("aria-hidden", String(mobile && !readerVisible));
  }

  function showToast(message) {
    window.clearTimeout(toastTimer);
    elements.toast.textContent = message;
    elements.toast.hidden = false;
    toastTimer = window.setTimeout(() => {
      elements.toast.hidden = true;
    }, 2600);
  }

  function showNotice(message, canRetry = true) {
    elements.streamNoticeText.textContent = message || "";
    elements.streamNotice.hidden = !message;
    elements.noticeRetry.hidden = !message || !canRetry;
  }

  function updateUrl(mode = "replace") {
    const url = new URL(window.location.href);
    if (state.mailboxId) url.searchParams.set("mailbox", state.mailboxId);
    else url.searchParams.delete("mailbox");
    url.searchParams.set("folder", state.folder);
    if (state.selectedMessageId) url.searchParams.set("message", state.selectedMessageId);
    else url.searchParams.delete("message");
    if (state.currentCursor) url.searchParams.set("pageCursor", state.currentCursor);
    else url.searchParams.delete("pageCursor");
    window.history[mode === "push" ? "pushState" : "replaceState"]({}, "", url);
  }

  function renderAccountSwitcher() {
    const mailbox = currentMailbox();
    if (!mailbox) {
      elements.accountAvatar.textContent = "M";
      elements.accountName.textContent = state.mailboxes.length ? "选择邮箱" : "暂无邮箱";
      elements.accountAddress.textContent = state.mailboxes.length ? "" : "在 Lark 中连接 Outlook";
      elements.mailboxStatusLabel.textContent = mailboxStatusText(null);
      elements.connectionState.className = "connection-state";
      elements.connectionState.title = "暂无已连接邮箱";
      elements.accountMenu.innerHTML = state.mailboxes.length
        ? ""
        : '<div class="account-empty"><strong>暂无已连接邮箱</strong><span>在 Lark 群发送 mail connect graph</span></div>';
      return;
    }

    elements.accountAvatar.textContent = initial(mailbox.displayName || mailbox.emailAddress);
    elements.accountName.textContent = mailbox.displayName || mailbox.emailAddress;
    elements.accountAddress.textContent = mailbox.emailAddress;
    elements.mailboxStatusLabel.textContent = mailboxStatusText(mailbox);
    elements.connectionState.className = "connection-state " + (mailbox.status === "active" ? "is-online" : "is-error");
    elements.connectionState.title = mailbox.status === "active" ? "邮箱连接正常" : "邮箱需要处理";

    elements.accountMenu.innerHTML = state.mailboxes.map((item) => {
      const active = item.mailboxId === state.mailboxId;
      const route = item.route
        ? (item.route.platform === "lark" ? (item.route.chatName || "已绑定 Lark 群") : "等待迁移到 Lark")
        : "未绑定群";
      return '<button class="account-option' + (active ? " is-active" : "") + '" type="button" role="option" aria-selected="' + active + '" data-mailbox-id="' + escapeHtml(item.mailboxId) + '">' +
        '<span class="account-avatar">' + escapeHtml(initial(item.displayName || item.emailAddress)) + '</span>' +
        '<span class="account-option-copy"><strong>' + escapeHtml(item.displayName || item.emailAddress) + '</strong><span>' + escapeHtml(item.emailAddress) + '</span><span class="account-route">' + escapeHtml(route) + '</span></span>' +
        '<span class="account-state ' + (item.status === "active" ? "" : "is-error") + '"></span>' +
        '</button>';
    }).join("");
  }

  function renderFolderState() {
    elements.folderTitle.textContent = folderLabel();
    elements.folderEyebrow.textContent = state.folder === "junk" ? "JUNK" : "INBOX";
    document.querySelectorAll("[data-folder]").forEach((button) => {
      const active = button.dataset.folder === state.folder;
      button.classList.toggle("is-active", active);
      button.setAttribute("aria-current", active ? "page" : "false");
    });
  }

  function renderListSkeleton() {
    elements.messageList.innerHTML = '<div class="skeleton-list" aria-label="正在载入邮件">' + Array.from({ length: 7 }, () =>
      '<div class="skeleton-row"><span class="skeleton-avatar"></span><span class="skeleton-lines"><span class="skeleton-line"></span><span class="skeleton-line"></span></span></div>'
    ).join("") + '</div>';
    elements.messageCount.textContent = "--";
    elements.streamEmpty.hidden = true;
    elements.pagination.hidden = true;
  }

  function matchesCurrentFilter(message) {
    if (state.filter === "code" && !message.verificationCode) return false;
    if (state.filter === "attachments" && !message.hasAttachments) return false;
    if (!state.query) return true;
    const haystack = compact([
      message.subject,
      message.fromName,
      message.fromAddress,
      message.bodyPreview,
      message.verificationCode,
    ].join(" ")).toLocaleLowerCase();
    return haystack.includes(state.query);
  }

  function applyFilters() {
    let visible = 0;
    elements.messageList.querySelectorAll(".message-item").forEach((row) => {
      const message = state.messages.find((item) => item.messageId === row.dataset.messageId);
      const shown = Boolean(message && matchesCurrentFilter(message));
      row.hidden = !shown;
      if (shown) visible += 1;
    });
    elements.messageCount.textContent = visible === state.messages.length
      ? state.messages.length + " 封"
      : visible + " / " + state.messages.length + " 封";
    elements.streamEmpty.hidden = visible > 0;
    if (visible === 0) {
      const filtered = state.messages.length > 0;
      elements.streamEmpty.querySelector("strong").textContent = filtered ? "没有匹配邮件" : "这里还没有邮件";
      elements.streamEmpty.querySelector("span:not(.empty-mark)").textContent = filtered ? "调整搜索或筛选条件。" : "切换文件夹或刷新后再试。";
      elements.emptyAction.textContent = filtered ? "清除筛选" : "刷新邮件";
    }
  }

  function resetFilters() {
    state.filter = "all";
    state.query = "";
    elements.search.value = "";
    document.querySelectorAll("[data-filter]").forEach((item) => {
      const active = item.dataset.filter === "all";
      item.classList.toggle("is-active", active);
      item.setAttribute("aria-pressed", String(active));
    });
    applyFilters();
  }

  function renderMessages() {
    elements.messageList.innerHTML = state.messages.map((message) => {
      const sender = compact(message.fromName || message.fromAddress || "未知发件人");
      const active = message.messageId === state.selectedMessageId;
      const code = compact(message.verificationCode);
      return '<button class="message-item' + (active ? " is-active" : "") + '" type="button" role="option" aria-selected="' + active + '" data-message-id="' + escapeHtml(message.messageId) + '">' +
        '<span class="message-avatar">' + escapeHtml(initial(sender)) + '</span>' +
        '<strong class="message-sender">' + escapeHtml(sender) + '</strong>' +
        '<time class="message-time" datetime="' + escapeHtml(message.receivedDateTime || "") + '">' + escapeHtml(formatTime(message.receivedDateTime, true)) + '</time>' +
        '<span class="message-subject">' + escapeHtml(message.subject || "(无主题)") + '</span>' +
        '<span class="message-preview">' + escapeHtml(compact(message.bodyPreview) || message.fromAddress || "暂无摘要") + '</span>' +
        '<span class="message-meta">' +
          (code ? '<span class="message-code">验证码 ' + escapeHtml(code) + '</span>' : '') +
          (message.hasAttachments ? '<span class="attachment-dot">附件 ' + Number(message.attachmentCount || 0) + '</span>' : '') +
        '</span>' +
      '</button>';
    }).join("");

    elements.messageList.scrollTop = 0;
    elements.pagination.hidden = !(state.nextCursor || state.cursorHistory.length || state.currentCursor);
    elements.previousPage.disabled = state.cursorHistory.length === 0;
    elements.nextPage.disabled = !state.nextCursor;
    elements.pageLabel.textContent = "第 " + state.page + " 页";
    applyFilters();
  }

  function setActiveMessage(messageId, ensureVisible = false) {
    elements.messageList.querySelectorAll(".message-item").forEach((row) => {
      const active = row.dataset.messageId === messageId;
      row.classList.toggle("is-active", active);
      row.setAttribute("aria-selected", String(active));
      if (active && ensureVisible) row.scrollIntoView({ block: "nearest" });
    });
  }

  function renderReaderPlaceholder(title = "选择一封邮件", message = "邮件正文会在这里打开。") {
    elements.readerPane.setAttribute("aria-busy", "false");
    elements.readerContent.innerHTML = '<section class="reader-placeholder"><span class="placeholder-mark" aria-hidden="true">' + icon("mail") + '</span><h2>' + escapeHtml(title) + '</h2><p>' + escapeHtml(message) + '</p></section>';
    elements.readerMobileTitle.textContent = "邮件正文";
  }

  function renderReaderSkeleton() {
    elements.readerPane.setAttribute("aria-busy", "true");
    elements.readerContent.innerHTML = '<div class="reader-skeleton" aria-label="正在载入邮件正文">' + Array.from({ length: 10 }, () => '<span class="skeleton-line"></span>').join("") + '</div>';
  }

  function renderReader(detail) {
    state.selectedDetail = detail;
    elements.readerPane.setAttribute("aria-busy", "false");
    elements.readerMobileTitle.textContent = detail.subject || "邮件正文";

    const attachments = Array.isArray(detail.attachments)
      ? detail.attachments.filter((item) => !item.isInline)
      : [];
    const header = '<header class="reader-header">' +
      '<p class="eyebrow">' + escapeHtml(detail.folderKind === "junk" ? "JUNK" : "MESSAGE") + '</p>' +
      '<div class="reader-header-row"><div><h2 id="readerTitle" tabindex="-1">' + escapeHtml(detail.subject || "(无主题)") + '</h2>' +
      '<div class="reader-sender"><strong>' + escapeHtml(detail.fromName || detail.fromAddress || "未知发件人") + '</strong>' +
      (detail.fromAddress && detail.fromAddress !== detail.fromName ? '<span>' + escapeHtml(detail.fromAddress) + '</span>' : '') +
      '<span>' + escapeHtml(formatTime(detail.receivedDateTime)) + '</span></div></div>' +
      (detail.webLink ? '<a class="reader-open" href="' + escapeHtml(detail.webLink) + '" target="_blank" rel="noopener noreferrer">' + icon("external") + '<span>在 Outlook 中打开</span></a>' : '') +
      '</div></header>';

    const code = compact(detail.verificationCode);
    const codeBanner = code
      ? '<section class="code-banner"><div class="code-copy"><span>验证码</span><strong>' + escapeHtml(code) + '</strong></div><button class="copy-code" type="button" data-copy-code="' + escapeHtml(code) + '">' + icon("copy") + '<span>复制</span></button></section>'
      : '';

    const attachmentBlock = attachments.length
      ? '<section class="attachment-section"><h3>附件 · ' + attachments.length + '</h3><div class="attachment-list">' + attachments.map((item) =>
        '<span class="attachment-item">' + icon("paperclip") + '<span>' + escapeHtml(item.name || "未命名附件") + (formatBytes(item.size) ? " · " + escapeHtml(formatBytes(item.size)) : "") + '</span></span>'
      ).join("") + '</div></section>'
      : '';

    elements.readerContent.innerHTML = '<article class="reader-document">' + header + codeBanner + attachmentBlock + '<section class="reader-body" id="mailBody"></section></article>';
    const body = document.getElementById("mailBody");
    if (detail.readerHtml) {
      const frame = document.createElement("iframe");
      frame.className = "mail-body-frame";
      frame.title = "邮件正文";
      frame.loading = "eager";
      frame.setAttribute("sandbox", "allow-popups allow-popups-to-escape-sandbox");
      frame.srcdoc = detail.readerHtml;
      body.appendChild(frame);
    } else {
      const text = document.createElement("pre");
      text.className = "mail-body-text";
      text.textContent = detail.bodyPlainText || "(无可用正文)";
      body.appendChild(text);
    }
    elements.readerPane.scrollTop = 0;
  }

  function detailKey(messageId, mailboxId = state.mailboxId, folder = state.folder) {
    return mailboxId + ":" + folder + ":" + messageId;
  }

  async function fetchDetail(messageId, signal, context = {}) {
    const mailboxId = context.mailboxId || state.mailboxId;
    const folder = context.folder || state.folder;
    const key = detailKey(messageId, mailboxId, folder);
    const cached = getCached(detailCache, key, DETAIL_CACHE_TTL);
    if (cached) return cached;
    const url = "/api/mailboxes/" + encodeURIComponent(mailboxId) + "/messages/" + encodeURIComponent(messageId) + "?folder=" + encodeURIComponent(folder);
    const data = await fetchJson(url, { signal });
    setCached(detailCache, key, data.message, MAX_DETAIL_CACHE);
    return data.message;
  }

  function scheduleNeighborPrefetch(messageId) {
    window.clearTimeout(prefetchTimer);
    if (navigator.connection && navigator.connection.saveData) return;
    const index = state.messages.findIndex((item) => item.messageId === messageId);
    const next = state.messages[index + 1];
    const mailboxId = state.mailboxId;
    const folder = state.folder;
    if (!next || getCached(detailCache, detailKey(next.messageId, mailboxId, folder), DETAIL_CACHE_TTL)) return;
    const run = () => fetchDetail(next.messageId, undefined, { mailboxId, folder }).catch(() => {});
    prefetchTimer = window.setTimeout(() => {
      if ("requestIdleCallback" in window) window.requestIdleCallback(run, { timeout: 1800 });
      else run();
    }, 650);
  }

  async function openMessage(messageId, options = {}) {
    if (!messageId) return;
    state.selectedMessageId = messageId;
    setActiveMessage(messageId, options.ensureVisible === true);
    if (options.pushHistory !== false) updateUrl("push");
    else updateUrl("replace");
    if (options.showMobile !== false && window.matchMedia("(max-width: 760px)").matches) {
      setMobileView("reader");
    }

    const key = detailKey(messageId);
    const cached = getCached(detailCache, key, DETAIL_CACHE_TTL);
    if (cached) {
      renderReader(cached);
      if (app.dataset.mobileView === "reader") document.getElementById("readerTitle")?.focus({ preventScroll: true });
      scheduleNeighborPrefetch(messageId);
      return;
    }

    if (state.detailRequest) state.detailRequest.abort();
    const controller = new AbortController();
    state.detailRequest = controller;
    renderReaderSkeleton();
    try {
      const detail = await fetchDetail(messageId, controller.signal);
      if (state.selectedMessageId !== messageId) return;
      renderReader(detail);
      if (app.dataset.mobileView === "reader") document.getElementById("readerTitle")?.focus({ preventScroll: true });
      scheduleNeighborPrefetch(messageId);
    } catch (error) {
      if (error.name === "AbortError") return;
      if (state.selectedMessageId !== messageId) return;
      renderReaderPlaceholder("无法打开邮件", error.message || String(error));
      showToast("邮件正文加载失败");
    } finally {
      if (state.detailRequest === controller) state.detailRequest = null;
    }
  }

  function listKey() {
    return state.mailboxId + ":" + state.folder + ":" + (state.currentCursor || "first");
  }

  async function loadMessages(options = {}) {
    const mailbox = currentMailbox();
    renderFolderState();
    showNotice("");
    if (!mailbox) {
      state.messages = [];
      renderMessages();
      renderReaderPlaceholder("还没有邮箱", "在 Lark 群发送 mail connect graph 连接 Outlook 账号。");
      return;
    }
    if (mailbox.providerType !== "graph_native") {
      state.messages = [];
      renderMessages();
      showNotice("当前邮箱使用 msOauth2api，Web 工作台暂只支持 Graph Native。", false);
      renderReaderPlaceholder("当前邮箱无法在网页中读取", "在 Lark 中将该邮箱切换为 Graph Native 后再试。");
      return;
    }

    const key = listKey();
    const cached = options.force ? null : getCached(listCache, key, LIST_CACHE_TTL);
    if (cached) {
      applyListPayload(cached, options);
      return;
    }

    if (state.listRequest) state.listRequest.abort();
    const controller = new AbortController();
    state.listRequest = controller;
    renderListSkeleton();
    elements.refresh.classList.add("is-spinning");
    elements.refresh.disabled = true;

    const params = new URLSearchParams({ folder: state.folder });
    if (state.currentCursor) params.set("pageCursor", state.currentCursor);
    if (options.force) params.set("fresh", "1");
    const url = "/api/mailboxes/" + encodeURIComponent(state.mailboxId) + "/messages?" + params;

    try {
      const payload = await fetchJson(url, { signal: controller.signal });
      setCached(listCache, key, payload, MAX_LIST_CACHE);
      applyListPayload(payload, options);
    } catch (error) {
      if (error.name === "AbortError") return;
      state.messages = [];
      renderMessages();
      showNotice(error.message || String(error));
      showToast("邮件列表加载失败");
    } finally {
      if (state.listRequest === controller) {
        state.listRequest = null;
        elements.refresh.classList.remove("is-spinning");
        elements.refresh.disabled = false;
      }
    }
  }

  function applyListPayload(payload, options = {}) {
    state.messages = Array.isArray(payload.messages) ? payload.messages : [];
    state.nextCursor = payload.nextPageCursor || "";
    if (payload.mailbox) {
      const index = state.mailboxes.findIndex((item) => item.mailboxId === payload.mailbox.mailboxId);
      if (index >= 0) state.mailboxes[index] = payload.mailbox;
      renderAccountSwitcher();
    }
    renderMessages();
    showNotice("");

    const requested = state.selectedMessageId;
    const selectedExists = requested && state.messages.some((item) => item.messageId === requested);
    const nextId = selectedExists ? requested : (state.messages[0]?.messageId || "");
    if (nextId) {
      openMessage(nextId, {
        pushHistory: false,
        showMobile: options.showMobileSelection === true,
      });
    } else {
      state.selectedMessageId = "";
      updateUrl("replace");
      renderReaderPlaceholder();
    }
  }

  async function selectMailbox(mailboxId, options = {}) {
    if (!mailboxId) return;
    if (state.listRequest) state.listRequest.abort();
    if (state.detailRequest) state.detailRequest.abort();
    state.mailboxId = mailboxId;
    state.selectedMessageId = "";
    state.currentCursor = "";
    state.nextCursor = "";
    state.cursorHistory = [];
    state.page = 1;
    state.selectedDetail = null;
    setMobileView("list");
    renderAccountSwitcher();
    elements.accountMenu.hidden = true;
    elements.accountTrigger.setAttribute("aria-expanded", "false");
    updateUrl(options.pushHistory === false ? "replace" : "push");
    await loadMessages();
  }

  async function loadMailboxes() {
    try {
      const payload = await fetchJson("/api/mailboxes");
      state.mailboxes = Array.isArray(payload.mailboxes) ? payload.mailboxes : [];
      const requestedExists = state.mailboxes.some((item) => item.mailboxId === state.mailboxId);
      if (!requestedExists) {
        state.mailboxId = (state.mailboxes.find((item) => item.providerType === "graph_native") || state.mailboxes[0] || {}).mailboxId || "";
      }
      renderAccountSwitcher();
      renderFolderState();
      updateUrl("replace");
      await loadMessages();
    } catch (error) {
      renderAccountSwitcher();
      showNotice(error.message || String(error));
      renderReaderPlaceholder("工作台载入失败", error.message || String(error));
    }
  }

  elements.accountTrigger.addEventListener("click", () => {
    const open = elements.accountMenu.hidden;
    elements.accountMenu.hidden = !open;
    elements.accountTrigger.setAttribute("aria-expanded", String(open));
  });

  elements.accountTrigger.addEventListener("keydown", (event) => {
    if (event.key !== "ArrowDown") return;
    event.preventDefault();
    elements.accountMenu.hidden = false;
    elements.accountTrigger.setAttribute("aria-expanded", "true");
    elements.accountMenu.querySelector(".account-option")?.focus();
  });

  elements.accountMenu.addEventListener("click", (event) => {
    const option = event.target.closest("[data-mailbox-id]");
    if (option) selectMailbox(option.dataset.mailboxId);
  });

  elements.accountMenu.addEventListener("keydown", (event) => {
    const options = Array.from(elements.accountMenu.querySelectorAll(".account-option"));
    const index = options.indexOf(document.activeElement);
    if ((event.key === "ArrowDown" || event.key === "ArrowUp") && options.length) {
      event.preventDefault();
      const direction = event.key === "ArrowDown" ? 1 : -1;
      const nextIndex = (index + direction + options.length) % options.length;
      options[nextIndex].focus();
    }
    if (event.key === "Escape") {
      elements.accountMenu.hidden = true;
      elements.accountTrigger.setAttribute("aria-expanded", "false");
      elements.accountTrigger.focus();
    }
  });

  document.addEventListener("click", (event) => {
    if (!event.target.closest(".account-switcher")) {
      elements.accountMenu.hidden = true;
      elements.accountTrigger.setAttribute("aria-expanded", "false");
    }
  });

  document.querySelectorAll("[data-folder]").forEach((button) => {
    button.addEventListener("click", () => {
      if (state.folder === button.dataset.folder) return;
      state.folder = button.dataset.folder;
      state.selectedMessageId = "";
      state.currentCursor = "";
      state.nextCursor = "";
      state.cursorHistory = [];
      state.page = 1;
      setMobileView("list");
      updateUrl("push");
      loadMessages();
    });
  });

  document.querySelectorAll("[data-filter]").forEach((button) => {
    button.addEventListener("click", () => {
      state.filter = button.dataset.filter;
      document.querySelectorAll("[data-filter]").forEach((item) => {
        const active = item.dataset.filter === state.filter;
        item.classList.toggle("is-active", active);
        item.setAttribute("aria-pressed", String(active));
      });
      applyFilters();
    });
  });

  elements.search.addEventListener("input", () => {
    state.query = compact(elements.search.value).toLocaleLowerCase();
    window.cancelAnimationFrame(filterFrame);
    filterFrame = window.requestAnimationFrame(applyFilters);
  });

  elements.messageList.addEventListener("click", (event) => {
    const row = event.target.closest("[data-message-id]");
    if (row) openMessage(row.dataset.messageId);
  });

  elements.readerContent.addEventListener("click", async (event) => {
    const button = event.target.closest("[data-copy-code]");
    if (!button) return;
    try {
      await navigator.clipboard.writeText(button.dataset.copyCode);
      const label = button.querySelector("span");
      if (label) label.textContent = "已复制";
      showToast("验证码已复制");
      window.setTimeout(() => {
        if (label) label.textContent = "复制";
      }, 1200);
    } catch {
      showToast("无法复制，请手动选择验证码");
    }
  });

  elements.noticeRetry.addEventListener("click", () => {
    listCache.delete(listKey());
    loadMessages({ force: true });
  });

  elements.emptyAction.addEventListener("click", () => {
    if (state.messages.length) {
      resetFilters();
      elements.search.focus();
      return;
    }
    listCache.delete(listKey());
    loadMessages({ force: true });
  });

  elements.refresh.addEventListener("click", () => {
    listCache.delete(listKey());
    loadMessages({ force: true });
  });

  elements.previousPage.addEventListener("click", () => {
    if (!state.cursorHistory.length) return;
    state.currentCursor = state.cursorHistory.pop() || "";
    state.page = Math.max(1, state.page - 1);
    state.selectedMessageId = "";
    updateUrl("push");
    loadMessages();
  });

  elements.nextPage.addEventListener("click", () => {
    if (!state.nextCursor) return;
    state.cursorHistory.push(state.currentCursor);
    state.currentCursor = state.nextCursor;
    state.page += 1;
    state.selectedMessageId = "";
    updateUrl("push");
    loadMessages();
  });

  elements.readerBack.addEventListener("click", () => {
    setMobileView("list");
    elements.messageList.querySelector('[data-message-id="' + CSS.escape(state.selectedMessageId) + '"]')?.focus();
  });

  window.addEventListener("keydown", (event) => {
    const target = event.target;
    const typing = target instanceof HTMLInputElement || target instanceof HTMLTextAreaElement;
    if ((event.key === "/" || ((event.metaKey || event.ctrlKey) && event.key.toLowerCase() === "k")) && !typing) {
      event.preventDefault();
      elements.search.focus();
      return;
    }
    if (event.key === "Escape") {
      elements.search.blur();
      elements.accountMenu.hidden = true;
      elements.accountTrigger.setAttribute("aria-expanded", "false");
      setMobileView("list");
      return;
    }
    if (typing) return;

    const visible = state.messages.filter(matchesCurrentFilter);
    const index = visible.findIndex((item) => item.messageId === state.selectedMessageId);
    if ((event.key === "j" || event.key === "k" || event.key === "ArrowDown" || event.key === "ArrowUp") && visible.length) {
      event.preventDefault();
      const direction = event.key === "j" || event.key === "ArrowDown" ? 1 : -1;
      const nextIndex = Math.min(visible.length - 1, Math.max(0, index + direction));
      openMessage(visible[nextIndex].messageId, { ensureVisible: true });
    }
    if (event.key.toLowerCase() === "y" && state.selectedDetail?.verificationCode) {
      navigator.clipboard.writeText(state.selectedDetail.verificationCode)
        .then(() => showToast("验证码已复制"))
        .catch(() => showToast("无法复制验证码"));
    }
  });

  window.addEventListener("popstate", () => {
    const url = new URL(window.location.href);
    const mailboxId = url.searchParams.get("mailbox") || state.mailboxId;
    const folder = url.searchParams.get("folder") === "junk" ? "junk" : "inbox";
    const messageId = url.searchParams.get("message") || "";
    const cursor = url.searchParams.get("pageCursor") || "";
    const listChanged = mailboxId !== state.mailboxId || folder !== state.folder || cursor !== state.currentCursor;
    state.mailboxId = mailboxId;
    state.folder = folder;
    state.currentCursor = cursor;
    state.selectedMessageId = messageId;
    setMobileView(messageId ? "reader" : "list");
    renderAccountSwitcher();
    renderFolderState();
    if (listChanged) loadMessages();
    else if (messageId) openMessage(messageId, { pushHistory: false });
  });

  renderListSkeleton();
  renderFolderState();
  setMobileView("list");
  window.addEventListener("resize", () => setMobileView(app.dataset.mobileView));
  loadMailboxes();
}
`;
