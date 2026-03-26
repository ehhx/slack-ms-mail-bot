import type { MailProviderType } from "./mail/types.ts";

export interface AppConfig {
  slackSigningSecret: string;
  slackBotToken: string;
  appBaseUrl: string;
  kvPath: string | null;
  slackApiTimeoutMs: number;
  mailPreviewMaxChars: number;
  graphApiBaseUrl: string;
  microsoftClientId: string;
  microsoftClientSecret: string;
  microsoftRedirectUri: string;
  microsoftAuthTenant: string;
  tokenEncryptionKey: string;
  webhookClientState: string;
  graphSubscriptionRenewalWindowMinutes: number;
  graphSubscriptionMaxMinutes: number;
  mailSyncPollIntervalMinutes: number;
  mailProviderDefault: MailProviderType;
  msOauth2apiBaseUrl: string | null;
  msOauth2apiPassword: string | null;
  msOauth2apiMailboxes: Array<"INBOX" | "Junk">;
}

let cachedConfig: AppConfig | null = null;

function readEnv(name: string): string | null {
  try {
    return Deno.env.get(name) ?? null;
  } catch {
    return null;
  }
}

function requireEnv(name: string): string {
  const value = readEnv(name);
  if (!value) {
    throw new Error(`Missing required env var: ${name}`);
  }
  return value;
}

function readIntEnv(name: string, fallback: number): number {
  const raw = readEnv(name);
  if (!raw) return fallback;
  const parsed = Number.parseInt(raw, 10);
  return Number.isFinite(parsed) ? parsed : fallback;
}

function normalizeBaseUrl(url: string): string {
  return url.replace(/\/+$/, "");
}

function readMailProviderDefault(): MailProviderType {
  const raw = (readEnv("MAIL_PROVIDER_DEFAULT") ?? "graph_native").trim().toLowerCase();
  if (raw === "msoauth2api") return "ms_oauth2api";
  return "graph_native";
}

function readMsOauth2ApiMailboxes(): Array<"INBOX" | "Junk"> {
  const raw = (readEnv("MSOAUTH2API_MAILBOX") ?? "INBOX,Junk")
    .split(",")
    .map((value) => value.trim())
    .filter(Boolean);
  const next: Array<"INBOX" | "Junk"> = [];
  for (const value of raw) {
    const normalized = value.toLowerCase();
    if (normalized === "junk" && !next.includes("Junk")) {
      next.push("Junk");
      continue;
    }
    if (!next.includes("INBOX")) {
      next.push("INBOX");
    }
  }
  return next.length > 0 ? next : ["INBOX", "Junk"];
}

async function sha256Hex(input: string): Promise<string> {
  const digest = await crypto.subtle.digest(
    "SHA-256",
    new TextEncoder().encode(input),
  );
  return Array.from(new Uint8Array(digest))
    .map((b) => b.toString(16).padStart(2, "0"))
    .join("");
}

export async function getConfigAsync(): Promise<AppConfig> {
  if (cachedConfig) return cachedConfig;

  const tokenEncryptionKey = requireEnv("TOKEN_ENCRYPTION_KEY");
  const derivedWebhookState = await sha256Hex(tokenEncryptionKey);

  cachedConfig = {
    slackSigningSecret: requireEnv("SLACK_SIGNING_SECRET"),
    slackBotToken: requireEnv("SLACK_BOT_TOKEN"),
    appBaseUrl: normalizeBaseUrl(requireEnv("APP_BASE_URL")),
    kvPath: readEnv("KV_PATH"),
    slackApiTimeoutMs: readIntEnv("SLACK_API_TIMEOUT_MS", 15000),
    mailPreviewMaxChars: readIntEnv("MAIL_PREVIEW_MAX_CHARS", 220),
    graphApiBaseUrl: normalizeBaseUrl(
      readEnv("GRAPH_API_BASE_URL") ?? "https://graph.microsoft.com/v1.0",
    ),
    microsoftClientId: requireEnv("MICROSOFT_CLIENT_ID"),
    microsoftClientSecret: requireEnv("MICROSOFT_CLIENT_SECRET"),
    microsoftRedirectUri: requireEnv("MICROSOFT_REDIRECT_URI"),
    microsoftAuthTenant: readEnv("MICROSOFT_AUTH_TENANT") ?? "common",
    tokenEncryptionKey,
    webhookClientState: readEnv("GRAPH_WEBHOOK_CLIENT_STATE") ?? derivedWebhookState,
    graphSubscriptionRenewalWindowMinutes: readIntEnv(
      "GRAPH_SUBSCRIPTION_RENEWAL_WINDOW_MINUTES",
      180,
    ),
    graphSubscriptionMaxMinutes: readIntEnv(
      "GRAPH_SUBSCRIPTION_MAX_MINUTES",
      4230,
    ),
    mailSyncPollIntervalMinutes: readIntEnv(
      "MAIL_SYNC_POLL_INTERVAL_MINUTES",
      15,
    ),
    mailProviderDefault: readMailProviderDefault(),
    msOauth2apiBaseUrl: readEnv("MSOAUTH2API_BASE_URL")
      ? normalizeBaseUrl(readEnv("MSOAUTH2API_BASE_URL")!)
      : null,
    msOauth2apiPassword: readEnv("MSOAUTH2API_PASSWORD"),
    msOauth2apiMailboxes: readMsOauth2ApiMailboxes(),
  };

  return cachedConfig;
}

export function clearConfigCache(): void {
  cachedConfig = null;
}
