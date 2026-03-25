import { getConfigAsync } from "../config.ts";

export class SlackApiError extends Error {
  readonly body: string;

  constructor(message: string, body: string) {
    super(message);
    this.body = body;
  }
}

async function fetchWithTimeout(
  input: string,
  init: RequestInit,
  timeoutMs: number,
): Promise<Response> {
  const controller = new AbortController();
  const timeout = setTimeout(() => controller.abort(), timeoutMs);
  try {
    return await fetch(input, { ...init, signal: controller.signal });
  } finally {
    clearTimeout(timeout);
  }
}

async function slackApi(method: string, body: Record<string, unknown>): Promise<Record<string, unknown>> {
  const config = await getConfigAsync();
  const response = await fetchWithTimeout(`https://slack.com/api/${method}`, {
    method: "POST",
    headers: {
      "content-type": "application/json; charset=utf-8",
      authorization: `Bearer ${config.slackBotToken}`,
    },
    body: JSON.stringify(body),
  }, config.slackApiTimeoutMs);

  const raw = await response.text().catch(() => "");
  if (!response.ok) {
    throw new SlackApiError(`Slack API ${method} failed with HTTP ${response.status}`, raw);
  }

  const data = raw ? JSON.parse(raw) as Record<string, unknown> : {};
  if (data.ok === false) {
    throw new SlackApiError(
      `Slack API ${method} returned error ${(data.error as string) ?? "unknown"}`,
      raw,
    );
  }
  return data;
}

export async function postChannelMessage(
  channel: string,
  text: string,
  blocks?: unknown[],
): Promise<void> {
  await slackApi("chat.postMessage", {
    channel,
    text,
    ...(blocks ? { blocks } : {}),
  });
}

export async function postEphemeralMessage(
  channel: string,
  user: string,
  text: string,
  blocks?: unknown[],
): Promise<void> {
  await slackApi("chat.postEphemeral", {
    channel,
    user,
    text,
    blocks: blocks ?? [{ type: "section", text: { type: "mrkdwn", text } }],
  });
}
