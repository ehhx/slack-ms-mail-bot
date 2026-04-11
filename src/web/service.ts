import { monitoredFoldersText, notificationBodyText } from "../mail/message.ts";
import {
  InvalidWebMailPageCursorError,
  listAllMailboxBundlesForWeb,
  loadMailboxWebView,
} from "../mail/service.ts";
import type {
  MailboxBundle,
  MailFolderKind,
  MailMessageSummary,
} from "../mail/types.ts";

export interface WebMessageDetail {
  message: MailMessageSummary;
  bodyPlainText: string;
  bodyHtml?: string;
}

export interface WebConsoleState {
  mailboxes: MailboxBundle[];
  selectedMailbox: MailboxBundle | null;
  selectedFolder: MailFolderKind;
  messages: MailMessageSummary[];
  selectedMessage: WebMessageDetail | null;
  pageIndex: number;
  hasPreviousPage: boolean;
  currentPageCursor?: string;
  error?: string;
  nextPageCursor?: string;
}

function normalizeFolderKind(input: string | null | undefined): MailFolderKind {
  return input === "junk" ? "junk" : "inbox";
}

function normalizePageIndex(
  input: string | null | undefined,
  hasCursor: boolean,
): number {
  if (!hasCursor) return 1;
  const parsed = Number.parseInt(String(input ?? ""), 10);
  if (!Number.isFinite(parsed) || parsed < 2) return 2;
  return parsed;
}

export function buildMessageDetail(
  message: MailMessageSummary,
): WebMessageDetail {
  return {
    message,
    bodyPlainText: notificationBodyText(message),
    bodyHtml: message.bodyContentType === "html" ? message.bodyText : undefined,
  };
}

function replaceMailbox(
  mailboxes: MailboxBundle[],
  nextMailbox: MailboxBundle | null,
): MailboxBundle[] {
  if (!nextMailbox) return mailboxes;
  return mailboxes.map((mailbox) =>
    mailbox.connection.mailboxId === nextMailbox.connection.mailboxId
      ? nextMailbox
      : mailbox
  );
}

export async function buildWebConsoleState(input: {
  mailboxId?: string | null;
  folder?: string | null;
  messageId?: string | null;
  pageCursor?: string | null;
  page?: string | null;
  fetchImpl?: typeof fetch;
}): Promise<WebConsoleState> {
  const selectedFolder = normalizeFolderKind(input.folder);
  let pageIndex = normalizePageIndex(input.page, Boolean(input.pageCursor));
  let mailboxes = await listAllMailboxBundlesForWeb();
  if (mailboxes.length === 0) {
    return {
      mailboxes,
      selectedMailbox: null,
      selectedFolder,
      messages: [],
      selectedMessage: null,
      pageIndex: 1,
      hasPreviousPage: false,
      currentPageCursor: undefined,
    };
  }

  let selectedMailbox =
    mailboxes.find((mailbox) =>
      mailbox.connection.mailboxId === input.mailboxId
    ) ??
      mailboxes.find((mailbox) =>
        mailbox.connection.providerType === "graph_native"
      ) ?? mailboxes[0];
  let error: string | undefined;

  if (
    input.mailboxId && selectedMailbox.connection.mailboxId !== input.mailboxId
  ) {
    error = "指定的邮箱不存在，已自动切换到第一个可用邮箱。";
  }

  if (selectedMailbox.connection.providerType !== "graph_native") {
    return {
      mailboxes,
      selectedMailbox,
      selectedFolder,
      messages: [],
      selectedMessage: null,
      pageIndex: 1,
      hasPreviousPage: false,
      currentPageCursor: undefined,
      error: error ??
        "当前选中的邮箱使用 msOauth2api，Web 控制台暂只支持 Graph Native。",
    };
  }

  let page: Awaited<ReturnType<typeof loadMailboxWebView>>;
  try {
    page = await loadMailboxWebView({
      mailboxId: selectedMailbox.connection.mailboxId,
      folderKind: selectedFolder,
      messageId: input.messageId,
      pageCursor: input.pageCursor,
      fetchImpl: input.fetchImpl,
    });
  } catch (listError) {
    if (listError instanceof InvalidWebMailPageCursorError) {
      pageIndex = 1;
      try {
        page = await loadMailboxWebView({
          mailboxId: selectedMailbox.connection.mailboxId,
          folderKind: selectedFolder,
          messageId: input.messageId,
          fetchImpl: input.fetchImpl,
        });
        error = error ?? "分页游标已失效，已自动回到最新邮件。";
      } catch (fallbackError) {
        const message = fallbackError instanceof Error
          ? fallbackError.message
          : String(fallbackError);
        return {
          mailboxes,
          selectedMailbox,
          selectedFolder,
          messages: [],
          selectedMessage: null,
          pageIndex,
          hasPreviousPage: false,
          currentPageCursor: undefined,
          error: error ?? `读取邮件列表失败：${message}`,
        };
      }
    } else {
      const message = listError instanceof Error
        ? listError.message
        : String(listError);
      return {
        mailboxes,
        selectedMailbox,
        selectedFolder,
        messages: [],
        selectedMessage: null,
        pageIndex,
        hasPreviousPage: pageIndex > 1,
        currentPageCursor: input.pageCursor ?? undefined,
        error: error ?? `读取邮件列表失败：${message}`,
      };
    }
  }
  selectedMailbox = page.bundle;
  mailboxes = replaceMailbox(mailboxes, selectedMailbox);

  const selectedMessage = page.selectedMessage
    ? buildMessageDetail(page.selectedMessage)
    : null;

  return {
    mailboxes,
    selectedMailbox,
    selectedFolder,
    messages: page.messages,
    selectedMessage,
    pageIndex,
    hasPreviousPage: pageIndex > 1,
    currentPageCursor: pageIndex > 1
      ? (input.pageCursor ?? undefined)
      : undefined,
    error,
    nextPageCursor: page.nextPageCursor,
  };
}

export function toWebMailboxSummary(
  bundle: MailboxBundle,
): Record<string, unknown> {
  return {
    mailboxId: bundle.connection.mailboxId,
    teamId: bundle.connection.teamId,
    emailAddress: bundle.connection.emailAddress,
    displayName: bundle.connection.displayName,
    providerType: bundle.connection.providerType,
    status: bundle.connection.status,
    lastError: bundle.connection.lastError,
    route: bundle.route
      ? {
        slackChannelId: bundle.route.slackChannelId,
        slackChannelName: bundle.route.slackChannelName,
        updatedAt: bundle.route.updatedAt,
      }
      : null,
    syncState: bundle.syncState
      ? {
        lastSyncAt: bundle.syncState.lastSyncAt,
        lastNotificationAt: bundle.syncState.lastNotificationAt,
        lastMessageReceivedAt: bundle.syncState.lastMessageReceivedAt,
        lastError: bundle.syncState.lastError,
      }
      : null,
    lease: bundle.lease
      ? {
        status: bundle.lease.status,
        expiresAt: bundle.lease.expiresAt,
        lastError: bundle.lease.lastError,
      }
      : null,
    monitoredFolders: monitoredFoldersText(bundle),
  };
}

export function toWebMessageSummary(
  message: MailMessageSummary,
): Record<string, unknown> {
  return {
    messageId: message.messageId,
    internetMessageId: message.internetMessageId,
    subject: message.subject,
    fromName: message.fromName,
    fromAddress: message.fromAddress,
    bodyPreview: message.bodyPreview,
    receivedDateTime: message.receivedDateTime,
    webLink: message.webLink,
    hasAttachments: message.hasAttachments,
    attachmentCount: message.attachments?.length ?? 0,
    folderKind: message.folderKind,
    folderName: message.folderName,
  };
}

export function toWebMessageDetail(
  detail: WebMessageDetail,
): Record<string, unknown> {
  return {
    ...toWebMessageSummary(detail.message),
    bodyContentType: detail.message.bodyContentType,
    bodyPlainText: detail.bodyPlainText,
    bodyHtml: detail.bodyHtml,
    attachments: detail.message.attachments ?? [],
    inlineImageCount: detail.message.inlineImages?.length ?? 0,
  };
}
