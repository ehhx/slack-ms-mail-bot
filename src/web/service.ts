import { monitoredFoldersText, notificationBodyText } from "../mail/message.ts";
import {
  getMailboxMessageForWeb,
  listAllMailboxBundlesForWeb,
  listMailboxMessagesForWeb,
} from "../mail/service.ts";
import type { MailFolderKind, MailMessageSummary, MailboxBundle } from "../mail/types.ts";

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
  error?: string;
  nextPageUrl?: string;
}

function normalizeFolderKind(input: string | null | undefined): MailFolderKind {
  return input === "junk" ? "junk" : "inbox";
}

function buildMessageDetail(message: MailMessageSummary): WebMessageDetail {
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
    mailbox.connection.mailboxId === nextMailbox.connection.mailboxId ? nextMailbox : mailbox
  );
}

export async function buildWebConsoleState(input: {
  mailboxId?: string | null;
  folder?: string | null;
  messageId?: string | null;
  fetchImpl?: typeof fetch;
}): Promise<WebConsoleState> {
  const selectedFolder = normalizeFolderKind(input.folder);
  let mailboxes = await listAllMailboxBundlesForWeb();
  if (mailboxes.length === 0) {
    return {
      mailboxes,
      selectedMailbox: null,
      selectedFolder,
      messages: [],
      selectedMessage: null,
    };
  }

  let selectedMailbox = mailboxes.find((mailbox) =>
    mailbox.connection.mailboxId === input.mailboxId
  ) ?? mailboxes.find((mailbox) => mailbox.connection.providerType === "graph_native") ?? mailboxes[0];
  let error: string | undefined;

  if (input.mailboxId && selectedMailbox.connection.mailboxId !== input.mailboxId) {
    error = "指定的邮箱不存在，已自动切换到第一个可用邮箱。";
  }

  if (selectedMailbox.connection.providerType !== "graph_native") {
    return {
      mailboxes,
      selectedMailbox,
      selectedFolder,
      messages: [],
      selectedMessage: null,
      error: error ?? "当前选中的邮箱使用 msOauth2api，Web 控制台暂只支持 Graph Native。",
    };
  }

  let page: Awaited<ReturnType<typeof listMailboxMessagesForWeb>>;
  try {
    page = await listMailboxMessagesForWeb({
      mailboxId: selectedMailbox.connection.mailboxId,
      folderKind: selectedFolder,
      fetchImpl: input.fetchImpl,
    });
  } catch (listError) {
    const message = listError instanceof Error ? listError.message : String(listError);
    return {
      mailboxes,
      selectedMailbox,
      selectedFolder,
      messages: [],
      selectedMessage: null,
      error: error ?? `读取邮件列表失败：${message}`,
    };
  }
  selectedMailbox = page.bundle;
  mailboxes = replaceMailbox(mailboxes, selectedMailbox);

  const selectedMessageId = input.messageId ?? page.messages[0]?.messageId ?? null;
  let selectedMessage: WebMessageDetail | null = null;

  if (selectedMessageId) {
    try {
      const detail = await getMailboxMessageForWeb({
        mailboxId: selectedMailbox.connection.mailboxId,
        messageId: selectedMessageId,
        folderKind: selectedFolder,
        fetchImpl: input.fetchImpl,
      });
      selectedMailbox = detail.bundle;
      mailboxes = replaceMailbox(mailboxes, selectedMailbox);
      selectedMessage = buildMessageDetail(detail.message);
    } catch (detailError) {
      const message = detailError instanceof Error ? detailError.message : String(detailError);
      error = error ?? `读取邮件详情失败：${message}`;
    }
  }

  return {
    mailboxes,
    selectedMailbox,
    selectedFolder,
    messages: page.messages,
    selectedMessage,
    error,
    nextPageUrl: page.nextPageUrl,
  };
}

export function toWebMailboxSummary(bundle: MailboxBundle): Record<string, unknown> {
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

export function toWebMessageSummary(message: MailMessageSummary): Record<string, unknown> {
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

export function toWebMessageDetail(detail: WebMessageDetail): Record<string, unknown> {
  return {
    ...toWebMessageSummary(detail.message),
    bodyContentType: detail.message.bodyContentType,
    bodyPlainText: detail.bodyPlainText,
    bodyHtml: detail.bodyHtml,
    attachments: detail.message.attachments ?? [],
    inlineImageCount: detail.message.inlineImages?.length ?? 0,
  };
}
