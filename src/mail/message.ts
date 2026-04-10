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

function readHtmlAttribute(tag: string, name: string): string | undefined {
  const pattern = new RegExp(`${name}\\s*=\\s*(?:"([^"]*)"|'([^']*)'|([^\\s>]+))`, "i");
  const matched = tag.match(pattern);
  const value = matched?.[1] ?? matched?.[2] ?? matched?.[3];
  return value ? decodeHtmlEntities(value) : undefined;
}

export function htmlToPlainText(input: string | undefined): string {
  if (!input) return "";
  return decodeHtmlEntities(
    input
      .replace(/<style[\s\S]*?<\/style>/gi, " ")
      .replace(/<script[\s\S]*?<\/script>/gi, " ")
      .replace(/<head[\s\S]*?<\/head>/gi, " ")
      .replace(/<a\b[^>]*href\s*=\s*("[^"]*"|'[^']*'|[^\s>]+)[^>]*>([\s\S]*?)<\/a>/gi, (full, _href, text) => {
        const href = readHtmlAttribute(full, "href");
        const label = decodeHtmlEntities(String(text ?? "")).replace(/<[^>]+>/g, " ").trim();
        if (!href) return label;
        if (!label) return href;
        return label === href ? href : `${label} (${href})`;
      })
      .replace(/<img\b[^>]*>/gi, (tag) => {
        const src = readHtmlAttribute(tag, "src");
        const alt = readHtmlAttribute(tag, "alt");
        if (src?.startsWith("cid:")) {
          return alt ? ` [内联图片：${alt}] ` : " [内联图片] ";
        }
        if (src) {
          return alt ? ` [图片：${alt} ${src}] ` : ` [图片：${src}] `;
        }
        return alt ? ` [图片：${alt}] ` : " [图片] ";
      })
      .replace(/<br\s*\/?>/gi, "\n")
      .replace(/<\/(p|div|section|article|tr|table|h[1-6])>/gi, "\n")
      .replace(/<li[^>]*>/gi, "• ")
      .replace(/<\/li>/gi, "\n")
      .replace(/<[^>]+>/g, " "),
  )
    .replace(/\r/g, "")
    .replace(/\n[ \t]+/g, "\n")
    .replace(/[ \t]+\n/g, "\n")
    .replace(/\n{3,}/g, "\n\n")
    .replace(/[ \t]{2,}/g, " ")
    .trim();
}

export function notificationBodyText(message: MailMessageSummary): string {
  const body = message.bodyText?.trim() || message.bodyPreview?.trim() || "";
  if (!body) {
    const inlineImageCount = (message.attachments ?? []).filter((attachment) =>
      attachment.isInline && attachment.contentType?.startsWith("image/")
    ).length;
    return inlineImageCount > 0 ? `此邮件正文主要由图片组成，包含 ${inlineImageCount} 张内联图片。` : "";
  }
  const text = message.bodyContentType === "html" ? htmlToPlainText(body) : body;
  if (text.trim()) return text;
  const inlineImageCount = (message.attachments ?? []).filter((attachment) =>
    attachment.isInline && attachment.contentType?.startsWith("image/")
  ).length;
  return inlineImageCount > 0 ? `此邮件正文主要由图片组成，包含 ${inlineImageCount} 张内联图片。` : text;
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
