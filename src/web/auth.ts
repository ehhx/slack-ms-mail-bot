import type { AppConfig } from "../config.ts";

export const WEB_SESSION_COOKIE_NAME = "mail_admin_session";
const WEB_SESSION_TTL_MS = 7 * 24 * 60 * 60 * 1000;

function textBytes(input: string): Uint8Array {
  return new TextEncoder().encode(input);
}

function base64UrlEncode(bytes: Uint8Array): string {
  let binary = "";
  for (const value of bytes) {
    binary += String.fromCharCode(value);
  }
  return btoa(binary).replace(/\+/g, "-").replace(/\//g, "_").replace(/=+$/g, "");
}

async function signSessionPayload(payload: string, secret: string): Promise<string> {
  const key = await crypto.subtle.importKey(
    "raw",
    textBytes(secret),
    { name: "HMAC", hash: "SHA-256" },
    false,
    ["sign"],
  );
  const signature = await crypto.subtle.sign("HMAC", key, textBytes(payload));
  return base64UrlEncode(new Uint8Array(signature));
}

function timingSafeEqual(left: string, right: string): boolean {
  const leftBytes = textBytes(left);
  const rightBytes = textBytes(right);
  const length = Math.max(leftBytes.length, rightBytes.length);
  let diff = leftBytes.length ^ rightBytes.length;
  for (let index = 0; index < length; index++) {
    diff |= (leftBytes[index] ?? 0) ^ (rightBytes[index] ?? 0);
  }
  return diff === 0;
}

function parseCookies(header: string | null): Record<string, string> {
  if (!header) return {};
  return header
    .split(";")
    .map((part) => part.trim())
    .filter(Boolean)
    .reduce<Record<string, string>>((acc, part) => {
      const separator = part.indexOf("=");
      if (separator <= 0) return acc;
      const key = part.slice(0, separator).trim();
      const value = part.slice(separator + 1).trim();
      acc[key] = decodeURIComponent(value);
      return acc;
    }, {});
}

function buildCookie(
  value: string,
  maxAgeSeconds: number,
  secure: boolean,
): string {
  return `${WEB_SESSION_COOKIE_NAME}=${encodeURIComponent(value)}; Path=/; HttpOnly; SameSite=Lax; Max-Age=${maxAgeSeconds}${secure ? "; Secure" : ""}`;
}

export function isWebConsoleEnabled(config: AppConfig): boolean {
  return Boolean(config.webAdminPassword);
}

export async function verifyWebAdminPassword(
  password: string,
  config: AppConfig,
): Promise<boolean> {
  if (!config.webAdminPassword) return false;
  return timingSafeEqual(password, config.webAdminPassword);
}

export async function buildWebAdminSessionCookie(config: AppConfig): Promise<string> {
  const expiresAt = Date.now() + WEB_SESSION_TTL_MS;
  const payload = `${WEB_SESSION_COOKIE_NAME}:${expiresAt}`;
  const signature = await signSessionPayload(payload, config.webSessionSecret);
  return buildCookie(
    `${expiresAt}.${signature}`,
    Math.floor(WEB_SESSION_TTL_MS / 1000),
    config.appBaseUrl.startsWith("https://"),
  );
}

export function buildClearWebAdminSessionCookie(config: AppConfig): string {
  return buildCookie("", 0, config.appBaseUrl.startsWith("https://"));
}

export async function isWebAdminAuthenticated(
  request: Request,
  config: AppConfig,
): Promise<boolean> {
  if (!config.webAdminPassword) return false;
  const cookies = parseCookies(request.headers.get("cookie"));
  const raw = cookies[WEB_SESSION_COOKIE_NAME];
  if (!raw) return false;
  const separator = raw.indexOf(".");
  if (separator <= 0) return false;
  const expiresAtRaw = raw.slice(0, separator);
  const signature = raw.slice(separator + 1);
  const expiresAt = Number.parseInt(expiresAtRaw, 10);
  if (!Number.isFinite(expiresAt) || expiresAt <= Date.now()) return false;
  const expected = await signSessionPayload(
    `${WEB_SESSION_COOKIE_NAME}:${expiresAt}`,
    config.webSessionSecret,
  );
  return timingSafeEqual(signature, expected);
}
