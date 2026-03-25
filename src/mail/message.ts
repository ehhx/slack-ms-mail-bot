import type { MailMessageSummary, MailboxBundle } from "./types.ts";

export function formatMailboxRef(mailboxId: string): string {
  return mailboxId.length > 8 ? mailboxId.slice(0, 8) : mailboxId;
}

export function buildDedupeKey(mailboxId: string, message: MailMessageSummary): string {
  return message.internetMessageId?.trim() || `${mailboxId}:${message.messageId}`;
}

export function toPreviewText(input: string | undefined, maxChars: number): string {
  const raw = (input ?? "").replace(/\s+/g, " ").trim();
  if (!raw) return "(No preview available)";
  if (raw.length <= maxChars) return raw;
  return `${raw.slice(0, Math.max(0, maxChars - 1)).trimEnd()}…`;
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
  return `provider=${formatProvider(bundle.connection.providerType)} route=${route} sync=${lastSync} sub=${subscription} lease=${lease}`;
}
