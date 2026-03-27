import type {
  MailAttachmentSummary,
  MailFolderKind,
  MailMessageSummary,
  MailboxBundle,
} from "./types.ts";

export function formatMailboxRef(mailboxId: string): string {
  return mailboxId.length > 8 ? mailboxId.slice(0, 8) : mailboxId;
}

export function buildDedupeKey(mailboxId: string, message: MailMessageSummary): string {
  return message.internetMessageId?.trim() || `${mailboxId}:${message.messageId}`;
}

export function formatFolderLabel(
  folderKind: MailFolderKind | undefined,
  fallback?: string,
): string {
  if (fallback?.trim()) return fallback.trim();
  if (folderKind === "junk") return "Junk";
  if (folderKind === "inbox") return "Inbox";
  return "Unknown";
}

export function monitoredFoldersText(bundle: MailboxBundle): string {
  if (bundle.connection.providerType === "ms_oauth2api") {
    const folderNames = new Set<string>();
    for (const state of Object.values(bundle.syncState?.folderStates ?? {})) {
      if (state?.folderName) folderNames.add(state.folderName);
    }
    return folderNames.size > 0 ? Array.from(folderNames).join(" + ") : "Inbox + Junk";
  }
  return "Inbox + Junk";
}

export function toPreviewText(input: string | undefined, maxChars: number): string {
  const raw = (input ?? "")
    .replace(/\r/g, "")
    .split("\n")
    .map((line) => line.replace(/[ \t]+/g, " ").trim())
    .join("\n")
    .replace(/\n{3,}/g, "\n\n")
    .trim();
  if (!raw) return "(No preview available)";
  if (raw.length <= maxChars) return raw;
  return `${raw.slice(0, Math.max(0, maxChars - 1)).trimEnd()}…`;
}

function decodeHtmlEntities(input: string): string {
  return input
    .replace(/&nbsp;/gi, " ")
    .replace(/&amp;/gi, "&")
    .replace(/&lt;/gi, "<")
    .replace(/&gt;/gi, ">")
    .replace(/&quot;/gi, "\"")
    .replace(/&#39;/gi, "'")
    .replace(/&#x([0-9a-f]+);/gi, (_, hex) => String.fromCodePoint(Number.parseInt(hex, 16)))
    .replace(/&#(\d+);/g, (_, dec) => String.fromCodePoint(Number.parseInt(dec, 10)));
}

export function htmlToPlainText(input: string | undefined): string {
  if (!input) return "";
  return decodeHtmlEntities(
    input
      .replace(/<style[\s\S]*?<\/style>/gi, " ")
      .replace(/<script[\s\S]*?<\/script>/gi, " ")
      .replace(/<head[\s\S]*?<\/head>/gi, " ")
      .replace(/<br\s*\/?>/gi, "\n")
      .replace(/<\/(p|div|section|article|tr|table|h[1-6])>/gi, "\n")
      .replace(/<li[^>]*>/gi, "• ")
      .replace(/<\/li>/gi, "\n")
      .replace(/<[^>]+>/g, " "),
  )
    .replace(/\r/g, "")
    .replace(/[ \t]+\n/g, "\n")
    .replace(/\n{3,}/g, "\n\n")
    .replace(/[ \t]{2,}/g, " ")
    .trim();
}

export function notificationBodyText(message: MailMessageSummary): string {
  const body = message.bodyText?.trim() || message.bodyPreview?.trim() || "";
  if (!body) return "";
  return message.bodyContentType === "html" ? htmlToPlainText(body) : body;
}

function formatBytes(input: number | undefined): string | null {
  if (!Number.isFinite(input) || !input || input <= 0) return null;
  if (input < 1024) return `${input} B`;
  if (input < 1024 * 1024) return `${(input / 1024).toFixed(1).replace(/\.0$/, "")} KB`;
  return `${(input / (1024 * 1024)).toFixed(1).replace(/\.0$/, "")} MB`;
}

export function attachmentSummaryText(
  attachments: MailAttachmentSummary[] | undefined,
  maxItems = 5,
): string | null {
  if (!attachments || attachments.length === 0) return null;
  const lines = attachments.slice(0, maxItems).map((attachment) => {
    const parts = [
      attachment.isInline ? "inline" : null,
      attachment.contentType ?? null,
      formatBytes(attachment.size),
    ].filter(Boolean);
    const suffix = parts.length ? ` (${parts.join(", ")})` : "";
    return `• ${attachment.name}${suffix}`;
  });
  if (attachments.length > maxItems) {
    lines.push(`…还有 ${attachments.length - maxItems} 个附件`);
  }
  return `*附件* (${attachments.length})\n${lines.join("\n")}`;
}

function formatProvider(providerType: MailboxBundle["connection"]["providerType"]): string {
  return providerType === "ms_oauth2api" ? "msOauth2api" : "graph_native";
}

export function mailboxStatusLine(bundle: MailboxBundle): string {
  const route = bundle.route ? `<#${bundle.route.slackChannelId}>` : "未配置";
  const lastSync = bundle.syncState?.lastSyncAt ?? "never";
  const pollingOnly = bundle.connection.providerType === "ms_oauth2api";
  const lease = pollingOnly ? "polling" : (bundle.lease?.expiresAt ?? "missing");
  const subscription = pollingOnly ? "polling" : (bundle.lease?.status ?? "missing");
  return `provider=${formatProvider(bundle.connection.providerType)} folders=${monitoredFoldersText(bundle)} route=${route} sync=${lastSync} sub=${subscription} lease=${lease}`;
}
