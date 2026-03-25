export interface SlackResponseMessage {
  response_type?: "ephemeral" | "in_channel";
  text: string;
  blocks?: unknown[];
  replace_original?: boolean;
}

export function jsonResponse(payload: unknown, status = 200): Response {
  return new Response(JSON.stringify(payload), {
    status,
    headers: { "content-type": "application/json; charset=utf-8" },
  });
}

export function slackEphemeral(
  text: string,
  blocks?: unknown[],
): SlackResponseMessage {
  return { response_type: "ephemeral", text, blocks };
}

export function slackInChannel(
  text: string,
  blocks?: unknown[],
): SlackResponseMessage {
  return { response_type: "in_channel", text, blocks };
}

export async function postToResponseUrl(
  responseUrl: string,
  message: SlackResponseMessage,
): Promise<void> {
  await fetch(responseUrl, {
    method: "POST",
    headers: { "content-type": "application/json; charset=utf-8" },
    body: JSON.stringify(message),
  });
}
