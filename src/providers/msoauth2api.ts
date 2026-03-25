import type { AppConfig } from "../config.ts";
import type { MailMessageSummary } from "../mail/types.ts";

interface MsOauth2ApiMailRecord {
  send?: string;
  subject?: string;
  text?: string;
  html?: string;
  date?: string;
}

export class MsOauth2ApiError extends Error {
  readonly status: number;
  readonly body: string;

  constructor(message: string, status: number, body: string) {
    super(message);
    this.status = status;
    this.body = body;
  }
}

function normalizeRecordDate(input?: string): string | undefined {
  if (!input) return undefined;
  const parsed = new Date(input);
  if (Number.isNaN(parsed.getTime())) return input;
  return parsed.toISOString();
}

async function stableMessageId(record: MsOauth2ApiMailRecord): Promise<string> {
  const payload = [
    record.send ?? "",
    record.subject ?? "",
    record.date ?? "",
    record.text?.slice(0, 200) ?? "",
  ].join("|");
  const digest = await crypto.subtle.digest(
    "SHA-256",
    new TextEncoder().encode(payload),
  );
  return Array.from(new Uint8Array(digest))
    .map((byte) => byte.toString(16).padStart(2, "0"))
    .join("");
}

function toMessageSummary(
  record: MsOauth2ApiMailRecord,
  messageId: string,
): MailMessageSummary {
  return {
    messageId,
    subject: record.subject?.trim() || "(no subject)",
    fromAddress: record.send?.trim() || undefined,
    bodyPreview: record.text ?? record.html ?? undefined,
    receivedDateTime: normalizeRecordDate(record.date),
  };
}

export async function fetchMsOauth2ApiMessages(input: {
  config: AppConfig;
  refreshToken: string;
  emailAddress: string;
  fetchImpl?: typeof fetch;
}): Promise<MailMessageSummary[]> {
  const baseUrl = input.config.msOauth2apiBaseUrl;
  if (!baseUrl) {
    throw new Error("MSOAUTH2API_BASE_URL is not configured");
  }

  const url = new URL("/api/mail-all", baseUrl);
  url.searchParams.set("refresh_token", input.refreshToken);
  url.searchParams.set("client_id", input.config.microsoftClientId);
  url.searchParams.set("email", input.emailAddress);
  url.searchParams.set("mailbox", input.config.msOauth2apiMailbox);
  if (input.config.msOauth2apiPassword) {
    url.searchParams.set("password", input.config.msOauth2apiPassword);
  }

  const response = await (input.fetchImpl ?? fetch)(url.toString(), {
    method: "GET",
    headers: { accept: "application/json" },
  });

  const raw = await response.text();
  if (!response.ok) {
    throw new MsOauth2ApiError(
      `msOauth2api request failed with HTTP ${response.status}`,
      response.status,
      raw,
    );
  }

  const parsed = raw ? JSON.parse(raw) : [];
  const records = Array.isArray(parsed) ? parsed as MsOauth2ApiMailRecord[] : [];
  const messages: MailMessageSummary[] = [];
  for (const record of records) {
    const messageId = await stableMessageId(record);
    messages.push(toMessageSummary(record, messageId));
  }

  messages.sort((a, b) => (a.receivedDateTime ?? "").localeCompare(b.receivedDateTime ?? ""));
  return messages;
}
