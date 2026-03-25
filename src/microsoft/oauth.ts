import type { AppConfig } from "../config.ts";

export const MICROSOFT_SCOPES = ["offline_access", "Mail.Read", "User.Read"];

export interface MicrosoftTokenSet {
  accessToken: string;
  refreshToken?: string;
  tokenType: string;
  scope: string;
  expiresIn: number;
  expiresAt: string;
}

function tokenEndpoint(config: AppConfig): string {
  return `https://login.microsoftonline.com/${config.microsoftAuthTenant}/oauth2/v2.0/token`;
}

export function buildMicrosoftAuthorizeUrl(
  config: AppConfig,
  state: string,
): string {
  const url = new URL(
    `https://login.microsoftonline.com/${config.microsoftAuthTenant}/oauth2/v2.0/authorize`,
  );
  url.searchParams.set("client_id", config.microsoftClientId);
  url.searchParams.set("response_type", "code");
  url.searchParams.set("redirect_uri", config.microsoftRedirectUri);
  url.searchParams.set("response_mode", "query");
  url.searchParams.set("scope", MICROSOFT_SCOPES.join(" "));
  url.searchParams.set("state", state);
  return url.toString();
}

async function exchangeTokenForm(
  config: AppConfig,
  form: URLSearchParams,
  fetchImpl: typeof fetch = fetch,
): Promise<MicrosoftTokenSet> {
  const response = await fetchImpl(tokenEndpoint(config), {
    method: "POST",
    headers: {
      "content-type": "application/x-www-form-urlencoded",
    },
    body: form.toString(),
  });

  const raw = await response.text();
  if (!response.ok) {
    throw new Error(`Microsoft token exchange failed: ${response.status} ${raw.slice(0, 300)}`);
  }

  const data = JSON.parse(raw) as {
    access_token: string;
    refresh_token?: string;
    token_type: string;
    scope: string;
    expires_in: number;
  };

  return {
    accessToken: data.access_token,
    refreshToken: data.refresh_token,
    tokenType: data.token_type,
    scope: data.scope,
    expiresIn: data.expires_in,
    expiresAt: new Date(Date.now() + data.expires_in * 1000).toISOString(),
  };
}

export function exchangeAuthorizationCode(
  config: AppConfig,
  code: string,
  fetchImpl: typeof fetch = fetch,
): Promise<MicrosoftTokenSet> {
  const form = new URLSearchParams({
    client_id: config.microsoftClientId,
    client_secret: config.microsoftClientSecret,
    redirect_uri: config.microsoftRedirectUri,
    grant_type: "authorization_code",
    code,
    scope: MICROSOFT_SCOPES.join(" "),
  });
  return exchangeTokenForm(config, form, fetchImpl);
}

export function refreshAccessToken(
  config: AppConfig,
  refreshToken: string,
  fetchImpl: typeof fetch = fetch,
): Promise<MicrosoftTokenSet> {
  const form = new URLSearchParams({
    client_id: config.microsoftClientId,
    client_secret: config.microsoftClientSecret,
    redirect_uri: config.microsoftRedirectUri,
    grant_type: "refresh_token",
    refresh_token: refreshToken,
    scope: MICROSOFT_SCOPES.join(" "),
  });
  return exchangeTokenForm(config, form, fetchImpl);
}
