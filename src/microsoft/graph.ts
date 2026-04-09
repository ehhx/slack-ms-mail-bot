import type { AppConfig } from "../config.ts";
import type {
  MailAttachmentSummary,
  MailInlineImage,
  MailFolderKind,
  MailMessageSummary,
} from "../mail/types.ts";

export interface MicrosoftGraphUser {
  id: string;
  displayName: string;
  mail?: string;
  userPrincipalName?: string;
}

export interface MicrosoftGraphMailFolder {
  id: string;
  displayName: string;
  wellKnownName?: string;
}

export interface MicrosoftGraphSubscription {
  id: string;
  resource: string;
  expirationDateTime: string;
  clientState?: string;
}

interface GraphItemBody {
  contentType?: "text" | "html";
  content?: string;
}

interface GraphAttachmentRecord {
  "@odata.type"?: string;
  id?: string;
  name?: string;
  contentType?: string;
  size?: number;
  isInline?: boolean;
  contentId?: string;
  contentBytes?: string;
}

interface GraphCollectionPage {
  value?: Array<Record<string, unknown>>;
  "@odata.nextLink"?: string;
  "@odata.deltaLink"?: string;
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

  async getMailFolder(folderName: string): Promise<MicrosoftGraphMailFolder> {
    return await this.request<MicrosoftGraphMailFolder>(
      // mailFolder 资源在 v1.0 下不支持选择 wellKnownName；
      // 这里只需要拿到 folder id 和 displayName，用于后续 delta query 与 Slack 展示。
      `/me/mailFolders/${encodePathSegment(folderName)}?$select=id,displayName`,
    );
  }

  async getInboxFolder(): Promise<MicrosoftGraphMailFolder> {
    return await this.getMailFolder("inbox");
  }

  async getJunkFolder(): Promise<MicrosoftGraphMailFolder> {
    return await this.getMailFolder("junkemail");
  }

  async getMessageDetail(messageId: string): Promise<Partial<MailMessageSummary>> {
    const message = await this.request<Record<string, unknown>>(
      `/me/messages/${encodePathSegment(messageId)}?$select=id,subject,body,uniqueBody,receivedDateTime,webLink,internetMessageId,from,hasAttachments`,
    );
    const from = message.from as { emailAddress?: { name?: string; address?: string } } | undefined;
    const body = (message.uniqueBody as GraphItemBody | undefined) ??
      (message.body as GraphItemBody | undefined);
    const hasAttachments = Boolean(message.hasAttachments);
    const attachments = hasAttachments
      ? await this.listMessageAttachments(messageId)
      : undefined;

    return {
      messageId: String(message.id ?? messageId),
      internetMessageId: message.internetMessageId
        ? String(message.internetMessageId)
        : undefined,
      subject: message.subject ? String(message.subject) : "(no subject)",
      fromName: from?.emailAddress?.name,
      fromAddress: from?.emailAddress?.address,
      bodyText: body?.content ? String(body.content) : undefined,
      bodyContentType: body?.contentType === "html" ? "html" : "text",
      receivedDateTime: message.receivedDateTime ? String(message.receivedDateTime) : undefined,
      webLink: message.webLink ? String(message.webLink) : undefined,
      hasAttachments,
      attachments,
    };
  }

  private mapMessageRecord(
    item: Record<string, unknown>,
    input: { folderKind?: MailFolderKind; folderName?: string },
  ): MailMessageSummary {
    const from = item.from as { emailAddress?: { name?: string; address?: string } } | undefined;
    return {
      messageId: String(item.id ?? ""),
      internetMessageId: item.internetMessageId ? String(item.internetMessageId) : undefined,
      subject: item.subject ? String(item.subject) : "(no subject)",
      fromName: from?.emailAddress?.name,
      fromAddress: from?.emailAddress?.address,
      bodyPreview: item.bodyPreview ? String(item.bodyPreview) : undefined,
      receivedDateTime: item.receivedDateTime ? String(item.receivedDateTime) : undefined,
      webLink: item.webLink ? String(item.webLink) : undefined,
      hasAttachments: Boolean(item.hasAttachments),
      folderKind: input.folderKind,
      folderName: input.folderName,
    };
  }

  async listFolderMessages(input: {
    folderId: string;
    folderKind?: MailFolderKind;
    folderName?: string;
    top?: number;
    pageUrl?: string;
  }): Promise<{ messages: MailMessageSummary[]; nextPageUrl?: string }> {
    const top = Math.max(1, Math.min(input.top ?? 25, 100));
    const withOrderBy =
      `${this.config.graphApiBaseUrl}/me/mailFolders/${encodePathSegment(input.folderId)}/messages?$select=id,subject,bodyPreview,receivedDateTime,webLink,internetMessageId,from,hasAttachments&$orderby=receivedDateTime%20desc&$top=${top}`;
    const fallbackUrl =
      `${this.config.graphApiBaseUrl}/me/mailFolders/${encodePathSegment(input.folderId)}/messages?$select=id,subject,bodyPreview,receivedDateTime,webLink,internetMessageId,from,hasAttachments&$top=${top}`;
    const requestUrl = input.pageUrl ?? withOrderBy;

    try {
      const page = await this.request<GraphCollectionPage>(requestUrl);
      return {
        messages: (page.value ?? []).map((item) => this.mapMessageRecord(item, input)),
        nextPageUrl: page["@odata.nextLink"] || undefined,
      };
    } catch (error) {
      if (input.pageUrl || !(error instanceof GraphApiError) || error.status !== 400) {
        throw error;
      }
      const page = await this.request<GraphCollectionPage>(fallbackUrl);
      return {
        messages: (page.value ?? []).map((item) => this.mapMessageRecord(item, input)),
        nextPageUrl: page["@odata.nextLink"] || undefined,
      };
    }
  }

  async listMessageAttachments(messageId: string): Promise<MailAttachmentSummary[]> {
    const attachments: MailAttachmentSummary[] = [];
    let nextUrl =
      `${this.config.graphApiBaseUrl}/me/messages/${encodePathSegment(messageId)}/attachments?$select=id,name,contentType,size,isInline,contentId`;

    while (nextUrl) {
      const page = await this.request<{
        value?: GraphAttachmentRecord[];
        "@odata.nextLink"?: string;
      }>(nextUrl);

      for (const item of page.value ?? []) {
        attachments.push({
          attachmentId: item.id ? String(item.id) : undefined,
          name: item.name ? String(item.name) : "(unnamed attachment)",
          contentType: item.contentType ? String(item.contentType) : undefined,
          size: typeof item.size === "number" ? item.size : undefined,
          isInline: Boolean(item.isInline),
          contentId: item.contentId ? String(item.contentId) : undefined,
        });
      }

      nextUrl = page["@odata.nextLink"] ?? "";
      if (!nextUrl) break;
    }

    return attachments;
  }

  async getInlineImageAttachmentContent(
    messageId: string,
    attachmentId: string,
  ): Promise<MailInlineImage | null> {
    const attachment = await this.request<GraphAttachmentRecord>(
      `/me/messages/${encodePathSegment(messageId)}/attachments/${encodePathSegment(attachmentId)}?$select=id,name,contentType,size,isInline,contentId,contentBytes`,
    );
    if (
      attachment["@odata.type"] &&
      attachment["@odata.type"] !== "#microsoft.graph.fileAttachment"
    ) {
      return null;
    }
    if (!attachment.contentBytes || !attachment.contentType?.startsWith("image/")) {
      return null;
    }
    return {
      attachmentId: String(attachment.id ?? attachmentId),
      name: attachment.name ? String(attachment.name) : "inline-image",
      contentType: String(attachment.contentType),
      size: typeof attachment.size === "number" ? attachment.size : undefined,
      contentId: attachment.contentId ? String(attachment.contentId) : undefined,
      dataBase64: String(attachment.contentBytes),
    };
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
    folderKind?: MailFolderKind;
    folderName?: string;
    deltaLink?: string;
  }): Promise<{ messages: MailMessageSummary[]; deltaLink: string }> {
    const messages: MailMessageSummary[] = [];
    let nextUrl = input.deltaLink ?? `${this.config.graphApiBaseUrl}/me/mailFolders/${encodePathSegment(input.folderId)}/messages/delta?$select=id,subject,bodyPreview,receivedDateTime,webLink,internetMessageId,from&changeType=created`;
    let latestDeltaLink: string | null = input.deltaLink ?? null;

    while (nextUrl) {
      const page = await this.request<GraphCollectionPage>(nextUrl);

      for (const item of page.value ?? []) {
        if (item["@removed"]) continue;
        messages.push(this.mapMessageRecord(item, input));
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

export function buildMailboxMessagesResource(): string {
  return "/me/messages";
}
