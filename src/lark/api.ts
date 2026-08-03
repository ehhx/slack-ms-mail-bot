import { type AppConfig, getConfigAsync } from "../config.ts";

export class LarkApiError extends Error {
  constructor(
    message: string,
    readonly body: string,
  ) {
    super(message);
  }
}

interface CachedTenantToken {
  value: string;
  expiresAtMs: number;
}

let cachedTenantToken: CachedTenantToken | null = null;

async function fetchWithTimeout(
  input: string,
  init: RequestInit,
  timeoutMs: number,
  fetchImpl: typeof fetch,
): Promise<Response> {
  const controller = new AbortController();
  const timeout = setTimeout(() => controller.abort(), timeoutMs);
  try {
    return await fetchImpl(input, { ...init, signal: controller.signal });
  } finally {
    clearTimeout(timeout);
  }
}

function larkUrl(config: AppConfig, path: string): string {
  return `${config.larkApiBaseUrl}${path}`;
}

async function getTenantAccessToken(
  config: AppConfig,
  fetchImpl: typeof fetch,
): Promise<string> {
  if (cachedTenantToken && cachedTenantToken.expiresAtMs > Date.now()) {
    return cachedTenantToken.value;
  }

  const response = await fetchWithTimeout(
    larkUrl(config, "/auth/v3/tenant_access_token/internal"),
    {
      method: "POST",
      headers: { "content-type": "application/json; charset=utf-8" },
      body: JSON.stringify({
        app_id: config.larkAppId,
        app_secret: config.larkAppSecret,
      }),
    },
    config.larkApiTimeoutMs,
    fetchImpl,
  );
  const raw = await response.text().catch(() => "");
  if (!response.ok) {
    throw new LarkApiError(
      `Lark tenant token request failed with HTTP ${response.status}`,
      raw,
    );
  }

  const payload = raw ? JSON.parse(raw) as Record<string, unknown> : {};
  const code = Number(payload.code ?? 0);
  const token = typeof payload.tenant_access_token === "string"
    ? payload.tenant_access_token
    : null;
  if (code !== 0 || !token) {
    throw new LarkApiError(
      `Lark tenant token request failed: ${
        String(payload.msg ?? "unknown error")
      }`,
      raw,
    );
  }

  const expiresInSeconds = Number(payload.expire ?? 7200);
  cachedTenantToken = {
    value: token,
    // 提前一分钟刷新，避免边缘运行时在长同步尾部命中过期 token。
    expiresAtMs: Date.now() + Math.max(60, expiresInSeconds - 60) * 1000,
  };
  return token;
}

export async function postLarkCard(input: {
  chatId: string;
  card: Record<string, unknown>;
  idempotencyKey?: string;
  fetchImpl?: typeof fetch;
}): Promise<{ messageId?: string }> {
  const config = await getConfigAsync();
  const fetchImpl = input.fetchImpl ?? fetch;
  const tenantAccessToken = await getTenantAccessToken(config, fetchImpl);
  const response = await fetchWithTimeout(
    larkUrl(config, "/im/v1/messages?receive_id_type=chat_id"),
    {
      method: "POST",
      headers: {
        "content-type": "application/json; charset=utf-8",
        authorization: `Bearer ${tenantAccessToken}`,
      },
      body: JSON.stringify({
        receive_id: input.chatId,
        msg_type: "interactive",
        content: JSON.stringify(input.card),
        uuid: input.idempotencyKey ?? crypto.randomUUID(),
      }),
    },
    config.larkApiTimeoutMs,
    fetchImpl,
  );
  const raw = await response.text().catch(() => "");
  if (!response.ok) {
    throw new LarkApiError(
      `Lark message request failed with HTTP ${response.status}`,
      raw,
    );
  }

  const payload = raw ? JSON.parse(raw) as Record<string, unknown> : {};
  if (Number(payload.code ?? 0) !== 0) {
    throw new LarkApiError(
      `Lark message request failed: ${String(payload.msg ?? "unknown error")}`,
      raw,
    );
  }

  const data = payload.data as Record<string, unknown> | undefined;
  return {
    messageId: typeof data?.message_id === "string"
      ? data.message_id
      : undefined,
  };
}

export function clearLarkTokenCache(): void {
  cachedTenantToken = null;
}
