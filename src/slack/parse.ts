import type { MailProviderType } from "../mail/types.ts";

export interface SlackSlashCommand {
  team_id: string;
  channel_id: string;
  channel_name?: string;
  user_id: string;
  command: string;
  text: string;
  response_url: string;
  trigger_id?: string;
}

export type MailCommand =
  | { kind: "help" }
  | { kind: "connect"; providerType?: MailProviderType }
  | { kind: "list" }
  | { kind: "status" }
  | { kind: "test"; mailbox: string }
  | { kind: "disconnect"; mailbox: string }
  | { kind: "sync"; mailbox: string }
  | { kind: "provider"; mailbox: string; providerType: MailProviderType }
  | { kind: "route"; mailbox: string; channelId: string; channelName?: string };

export interface SlackInteractivityPayload {
  type: string;
  user: { id: string };
  team: { id: string };
  channel?: { id: string };
  actions: Array<{
    action_id: string;
    value?: string;
  }>;
}

export function parseForm(body: string): Record<string, string> {
  const params = new URLSearchParams(body);
  const out: Record<string, string> = {};
  for (const [k, v] of params.entries()) out[k] = v;
  return out;
}

export function parseSlashCommand(body: string): SlackSlashCommand {
  const form = parseForm(body);
  return {
    team_id: form.team_id,
    channel_id: form.channel_id,
    channel_name: form.channel_name,
    user_id: form.user_id,
    command: form.command,
    text: form.text ?? "",
    response_url: form.response_url,
    trigger_id: form.trigger_id,
  };
}

function parseSlackChannelRef(input: string): { channelId: string; channelName?: string } | null {
  const trimmed = input.trim();
  const mention = trimmed.match(/^<#([A-Z0-9]+)(?:\|([^>]+))?>$/i);
  if (mention) {
    return { channelId: mention[1], channelName: mention[2] };
  }
  if (/^[CGD][A-Z0-9]+$/i.test(trimmed)) {
    return { channelId: trimmed };
  }
  return null;
}

function parseProviderType(input: string | undefined): MailProviderType | null {
  const raw = (input ?? "").trim().toLowerCase();
  if (!raw) return null;
  if (raw === "graph" || raw === "graph_native") return "graph_native";
  if (raw === "msoauth2api" || raw === "ms_oauth2api") return "ms_oauth2api";
  return null;
}

export function parseMailCommand(text: string): MailCommand {
  const raw = (text ?? "").trim();
  if (!raw) return { kind: "help" };

  const [head, ...rest] = raw.split(/\s+/);
  const tail = rest.join(" ").trim();

  switch (head.toLowerCase()) {
    case "help":
      return { kind: "help" };
    case "connect": {
      const providerType = parseProviderType(tail);
      return providerType ? { kind: "connect", providerType } : { kind: "connect" };
    }
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
    case "provider": {
      const [mailbox, providerRaw] = tail.split(/\s+/, 2);
      const providerType = parseProviderType(providerRaw);
      if (!mailbox || !providerType) return { kind: "help" };
      return { kind: "provider", mailbox, providerType };
    }
    case "route": {
      const [mailbox, channelToken] = tail.split(/\s+/, 2);
      if (!mailbox || !channelToken) return { kind: "help" };
      const parsedChannel = parseSlackChannelRef(channelToken);
      if (!parsedChannel) return { kind: "help" };
      return {
        kind: "route",
        mailbox,
        channelId: parsedChannel.channelId,
        channelName: parsedChannel.channelName,
      };
    }
    default:
      return { kind: "help" };
  }
}

export function parseInteractivityPayload(
  body: string,
): SlackInteractivityPayload {
  const form = parseForm(body);
  const payloadRaw = form.payload;
  if (!payloadRaw) throw new Error("Missing interactivity payload");
  return JSON.parse(payloadRaw) as SlackInteractivityPayload;
}
