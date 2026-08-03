import type { MailProviderType } from "../mail/types.ts";

export type LarkMailCommand =
  | { kind: "help" }
  | { kind: "whoami" }
  | { kind: "connect"; providerType?: MailProviderType }
  | { kind: "claim"; mailbox: string }
  | { kind: "list" }
  | { kind: "status" }
  | { kind: "test"; mailbox: string }
  | { kind: "disconnect"; mailbox: string }
  | { kind: "sync"; mailbox: string }
  | { kind: "provider"; mailbox: string; providerType: MailProviderType }
  | { kind: "route"; mailbox: string };

export interface LarkTextMessageEvent {
  tenantId: string;
  senderOpenId: string;
  chatId: string;
  chatType: "group" | "p2p" | string;
  text: string;
}

type JsonRecord = Record<string, unknown>;

function isRecord(value: unknown): value is JsonRecord {
  return Boolean(value) && typeof value === "object";
}

function parseProviderType(input: string | undefined): MailProviderType | null {
  const raw = (input ?? "").trim().toLowerCase();
  if (!raw) return null;
  if (raw === "graph" || raw === "graph_native") return "graph_native";
  if (raw === "msoauth2api" || raw === "ms_oauth2api") return "ms_oauth2api";
  return null;
}

function stripLeadingBotMention(input: string): string {
  return input.trim().replace(/^@[_a-zA-Z0-9-]+\s+/, "");
}

export function parseLarkMailCommand(input: string): LarkMailCommand | null {
  const normalized = stripLeadingBotMention(input).replace(/^\//, "");
  const matched = normalized.match(/^mail(?:\s+(.*))?$/iu);
  if (!matched) return null;

  const raw = (matched[1] ?? "").trim();
  if (!raw) return { kind: "help" };
  const [head, ...rest] = raw.split(/\s+/);
  const tail = rest.join(" ").trim();

  switch (head.toLowerCase()) {
    case "help":
      return { kind: "help" };
    case "whoami":
      return { kind: "whoami" };
    case "connect": {
      const providerType = parseProviderType(tail);
      return providerType
        ? { kind: "connect", providerType }
        : { kind: "connect" };
    }
    case "claim":
      return tail ? { kind: "claim", mailbox: tail } : { kind: "help" };
    case "list":
      return { kind: "list" };
    case "status":
      return { kind: "status" };
    case "test":
      return tail ? { kind: "test", mailbox: tail } : { kind: "help" };
    case "disconnect":
      return tail ? { kind: "disconnect", mailbox: tail } : { kind: "help" };
    case "sync":
      return tail ? { kind: "sync", mailbox: tail } : { kind: "help" };
    case "route":
      return tail ? { kind: "route", mailbox: tail } : { kind: "help" };
    case "provider": {
      const [mailbox, providerRaw] = tail.split(/\s+/, 2);
      const providerType = parseProviderType(providerRaw);
      return mailbox && providerType
        ? { kind: "provider", mailbox, providerType }
        : { kind: "help" };
    }
    default:
      return { kind: "help" };
  }
}

export function parseLarkPayload(bodyText: string): JsonRecord {
  const parsed = JSON.parse(bodyText) as unknown;
  if (!isRecord(parsed)) throw new Error("Invalid Lark event payload");
  if (typeof parsed.encrypt === "string") {
    throw new Error("Encrypted Lark events are not enabled by this deployment");
  }
  return parsed;
}

export function getLarkUrlVerificationChallenge(
  payload: JsonRecord,
): string | null {
  return payload.type === "url_verification" &&
      typeof payload.challenge === "string"
    ? payload.challenge
    : null;
}

export function getLarkVerificationToken(payload: JsonRecord): string | null {
  if (typeof payload.token === "string") return payload.token;
  const header = isRecord(payload.header) ? payload.header : null;
  return typeof header?.token === "string" ? header.token : null;
}

export function parseLarkTextMessageEvent(
  payload: JsonRecord,
): LarkTextMessageEvent | null {
  const header = isRecord(payload.header) ? payload.header : null;
  const event = isRecord(payload.event) ? payload.event : null;
  if (header?.event_type !== "im.message.receive_v1" || !event) return null;

  const message = isRecord(event.message) ? event.message : null;
  const sender = isRecord(event.sender) ? event.sender : null;
  const senderId = sender && isRecord(sender.sender_id)
    ? sender.sender_id
    : null;
  if (
    message?.message_type !== "text" ||
    typeof message.chat_id !== "string" ||
    typeof message.content !== "string" ||
    typeof senderId?.open_id !== "string" ||
    sender?.sender_type === "app"
  ) {
    return null;
  }

  let content: JsonRecord;
  try {
    content = JSON.parse(message.content) as JsonRecord;
  } catch {
    return null;
  }
  if (typeof content.text !== "string") return null;

  return {
    tenantId: typeof header.tenant_key === "string"
      ? header.tenant_key
      : String(header.app_id ?? "lark"),
    senderOpenId: senderId.open_id,
    chatId: message.chat_id,
    chatType: typeof message.chat_type === "string"
      ? message.chat_type
      : "group",
    text: content.text,
  };
}
