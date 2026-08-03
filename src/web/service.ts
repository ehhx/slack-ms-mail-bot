import {
  detectVerificationCode,
  monitoredFoldersText,
  notificationBodyText,
} from "../mail/message.ts";
import type { MailboxBundle, MailMessageSummary } from "../mail/types.ts";
import { buildReaderDocumentHtml } from "./email_html.ts";

export interface WebMessageDetail {
  message: MailMessageSummary;
  bodyPlainText: string;
  bodyHtml?: string;
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
        platform: bundle.route.platform,
        chatId: bundle.route.chatId,
        chatName: bundle.route.chatName,
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
    verificationCode: detectVerificationCode({
      subject: message.subject,
      body: message.bodyPreview,
    }),
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
    readerHtml: detail.bodyHtml
      ? buildReaderDocumentHtml(
        detail.bodyHtml,
        detail.message.inlineImages,
      )
      : undefined,
    attachments: detail.message.attachments ?? [],
    inlineImageCount: detail.message.inlineImages?.length ?? 0,
    verificationCode: detectVerificationCode({
      subject: detail.message.subject,
      body: detail.bodyPlainText,
    }),
  };
}
