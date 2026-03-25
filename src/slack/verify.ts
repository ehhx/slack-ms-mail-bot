function toHex(bytes: ArrayBuffer): string {
  const u8 = new Uint8Array(bytes);
  let out = "";
  for (const b of u8) out += b.toString(16).padStart(2, "0");
  return out;
}

function timingSafeEqual(a: string, b: string): boolean {
  if (a.length !== b.length) return false;
  let diff = 0;
  for (let i = 0; i < a.length; i++) diff |= a.charCodeAt(i) ^ b.charCodeAt(i);
  return diff === 0;
}

export async function computeSlackSignatureV0(
  signingSecret: string,
  timestamp: string,
  rawBody: string,
): Promise<string> {
  const key = await crypto.subtle.importKey(
    "raw",
    new TextEncoder().encode(signingSecret),
    { name: "HMAC", hash: "SHA-256" },
    false,
    ["sign"],
  );

  const base = `v0:${timestamp}:${rawBody}`;
  const signature = await crypto.subtle.sign(
    "HMAC",
    key,
    new TextEncoder().encode(base),
  );
  return `v0=${toHex(signature)}`;
}

export async function verifySlackRequest(
  request: Request,
  rawBody: string,
  signingSecret: string,
  options?: { nowMs?: number; toleranceSeconds?: number },
): Promise<{ ok: true } | { ok: false; error: string }> {
  const timestamp = request.headers.get("x-slack-request-timestamp");
  const signature = request.headers.get("x-slack-signature");
  if (!timestamp || !signature) {
    return { ok: false, error: "Missing Slack signature headers" };
  }

  const nowMs = options?.nowMs ?? Date.now();
  const toleranceSeconds = options?.toleranceSeconds ?? 60 * 5;
  const tsSeconds = Number.parseInt(timestamp, 10);
  if (!Number.isFinite(tsSeconds)) {
    return { ok: false, error: "Invalid Slack timestamp" };
  }

  const ageSeconds = Math.abs(Math.floor(nowMs / 1000) - tsSeconds);
  if (ageSeconds > toleranceSeconds) {
    return { ok: false, error: "Slack request timestamp too old" };
  }

  const expected = await computeSlackSignatureV0(signingSecret, timestamp, rawBody);
  if (!timingSafeEqual(expected, signature)) {
    return { ok: false, error: "Slack signature mismatch" };
  }

  return { ok: true };
}
