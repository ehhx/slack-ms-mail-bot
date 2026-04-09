import { getConfigAsync } from "../config.ts";
import type { MailInlineImage } from "../mail/types.ts";

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

function base64ToBytes(input: string): Uint8Array {
  const binary = atob(input);
  const bytes = new Uint8Array(binary.length);
  for (let index = 0; index < binary.length; index++) {
    bytes[index] = binary.charCodeAt(index);
  }
  return bytes;
}

export async function postChannelMessage(
  channel: string,
  text: string,
  blocks?: unknown[],
): Promise<{ ts?: string }> {
  const result = await slackApi("chat.postMessage", {
    channel,
    text,
    ...(blocks ? { blocks } : {}),
  });
  return { ts: typeof result.ts === "string" ? result.ts : undefined };
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

export async function uploadInlineImageToSlack(input: {
  channel: string;
  threadTs?: string;
  image: MailInlineImage;
}): Promise<void> {
  const config = await getConfigAsync();
  const bytes = base64ToBytes(input.image.dataBase64);
  const prepare = await slackApi("files.getUploadURLExternal", {
    filename: input.image.name,
    length: bytes.byteLength,
    alt_txt: input.image.name,
  });
  const uploadUrl = typeof prepare.upload_url === "string" ? prepare.upload_url : null;
  const fileId = typeof prepare.file_id === "string" ? prepare.file_id : null;
  if (!uploadUrl || !fileId) {
    throw new SlackApiError(
      "Slack API files.getUploadURLExternal returned invalid payload",
      JSON.stringify(prepare),
    );
  }

  const uploadResponse = await fetchWithTimeout(uploadUrl, {
    method: "POST",
    headers: {
      "content-type": input.image.contentType,
    },
    body: bytes,
  }, config.slackApiTimeoutMs);
  const uploadRaw = await uploadResponse.text().catch(() => "");
  if (!uploadResponse.ok) {
    throw new SlackApiError(
      `Slack external upload failed with HTTP ${uploadResponse.status}`,
      uploadRaw,
    );
  }

  await slackApi("files.completeUploadExternal", {
    files: [{ id: fileId, title: input.image.name }],
    channel_id: input.channel,
    ...(input.threadTs ? { thread_ts: input.threadTs } : {}),
    initial_comment: `🖼️ ${input.image.name}`,
  });
}
