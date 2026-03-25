import type { AppConfig } from "../config.ts";
import type { MailMessageSummary } from "../mail/types.ts";

export interface MicrosoftGraphUser {
  id: string;
  displayName: string;
  mail?: string;
  userPrincipalName?: string;
}

export interface MicrosoftGraphMailFolder {
  id: string;
  displayName: string;
}

export interface MicrosoftGraphSubscription {
  id: string;
  resource: string;
  expirationDateTime: string;
  clientState?: string;
}

export class GraphApiError extends Error {
  readonly status: number;
  readonly body: string;

  constructor(message: string, status: number, body: string) {
    super(message);
    this.status = status;
    this.body = body;
  }
}

function encodePathSegment(value: string): string {
  return encodeURIComponent(value).replace(/%2F/g, "/");
}

export class MicrosoftGraphClient {
  private readonly config: AppConfig;
  private readonly accessToken: string;
  private readonly fetchImpl: typeof fetch;

  constructor(
    config: AppConfig,
    accessToken: string,
    fetchImpl: typeof fetch = fetch,
  ) {
    this.config = config;
    this.accessToken = accessToken;
    this.fetchImpl = fetchImpl;
  }

  private async request<T>(
    pathOrUrl: string,
    init: RequestInit = {},
    json = true,
  ): Promise<T> {
    const url = pathOrUrl.startsWith("http")
      ? pathOrUrl
      : `${this.config.graphApiBaseUrl}${pathOrUrl}`;
    const headers = new Headers(init.headers);
    headers.set("authorization", `Bearer ${this.accessToken}`);
    if (!headers.has("content-type") && init.body) {
      headers.set("content-type", "application/json; charset=utf-8");
    }
    headers.set("accept", "application/json");

    const response = await this.fetchImpl(url, { ...init, headers });
    const raw = await response.text();
    if (!response.ok) {
      throw new GraphApiError(
        `Microsoft Graph request failed: ${response.status}`,
        response.status,
        raw,
      );
    }

    if (!json) return raw as T;
    return raw ? JSON.parse(raw) as T : ({} as T);
  }

  async getCurrentUser(): Promise<MicrosoftGraphUser> {
    return await this.request<MicrosoftGraphUser>(
      "/me?$select=id,displayName,mail,userPrincipalName",
    );
  }

  async getInboxFolder(): Promise<MicrosoftGraphMailFolder> {
    return await this.request<MicrosoftGraphMailFolder>(
      "/me/mailFolders/inbox?$select=id,displayName",
    );
  }

  async createSubscription(input: {
    resource: string;
    notificationUrl: string;
    lifecycleNotificationUrl?: string;
    clientState: string;
    expirationDateTime: string;
  }): Promise<MicrosoftGraphSubscription> {
    return await this.request<MicrosoftGraphSubscription>("/subscriptions", {
      method: "POST",
      body: JSON.stringify({
        changeType: "created",
        notificationUrl: input.notificationUrl,
        lifecycleNotificationUrl: input.lifecycleNotificationUrl,
        resource: input.resource,
        expirationDateTime: input.expirationDateTime,
        clientState: input.clientState,
      }),
    });
  }

  async renewSubscription(
    subscriptionId: string,
    expirationDateTime: string,
  ): Promise<MicrosoftGraphSubscription> {
    return await this.request<MicrosoftGraphSubscription>(
      `/subscriptions/${encodePathSegment(subscriptionId)}`,
      {
        method: "PATCH",
        body: JSON.stringify({ expirationDateTime }),
      },
    );
  }

  async deleteSubscription(subscriptionId: string): Promise<void> {
    await this.request<unknown>(
      `/subscriptions/${encodePathSegment(subscriptionId)}`,
      { method: "DELETE" },
      false,
    );
  }

  async collectMessageDelta(input: {
    folderId: string;
    deltaLink?: string;
  }): Promise<{ messages: MailMessageSummary[]; deltaLink: string }> {
    const messages: MailMessageSummary[] = [];
    let nextUrl = input.deltaLink ?? `${this.config.graphApiBaseUrl}/me/mailFolders/${encodePathSegment(input.folderId)}/messages/delta?$select=id,subject,bodyPreview,receivedDateTime,webLink,internetMessageId,from&changeType=created`;
    let latestDeltaLink: string | null = input.deltaLink ?? null;

    while (nextUrl) {
      const page = await this.request<{
        value?: Array<Record<string, unknown>>;
        "@odata.nextLink"?: string;
        "@odata.deltaLink"?: string;
      }>(nextUrl);

      for (const item of page.value ?? []) {
        if (item["@removed"]) continue;
        const from = item.from as { emailAddress?: { name?: string; address?: string } } | undefined;
        messages.push({
          messageId: String(item.id ?? ""),
          internetMessageId: item.internetMessageId ? String(item.internetMessageId) : undefined,
          subject: item.subject ? String(item.subject) : "(no subject)",
          fromName: from?.emailAddress?.name,
          fromAddress: from?.emailAddress?.address,
          bodyPreview: item.bodyPreview ? String(item.bodyPreview) : undefined,
          receivedDateTime: item.receivedDateTime ? String(item.receivedDateTime) : undefined,
          webLink: item.webLink ? String(item.webLink) : undefined,
        });
      }

      if (page["@odata.deltaLink"]) {
        latestDeltaLink = page["@odata.deltaLink"];
      }
      nextUrl = page["@odata.nextLink"] ?? "";
      if (!nextUrl) break;
    }

    if (!latestDeltaLink) {
      throw new Error("Microsoft Graph delta query did not return a deltaLink");
    }

    return { messages, deltaLink: latestDeltaLink };
  }
}

export function buildMailboxResource(folderId: string): string {
  return `/me/mailFolders/${encodePathSegment(folderId)}/messages`;
}
