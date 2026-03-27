export type MailboxConnectionStatus = "active" | "needs_reauth" | "disconnected";
export type LeaseStatus = "active" | "missing" | "degraded";
export type MailProviderType = "graph_native" | "ms_oauth2api";
export type MailFolderKind = "inbox" | "junk";

export interface MailboxFolderSyncState {
  folderId: string;
  folderName: string;
  deltaLink?: string;
  lastMessageReceivedAt?: string;
}

export interface MailAttachmentSummary {
  attachmentId?: string;
  name: string;
  contentType?: string;
  size?: number;
  isInline?: boolean;
  contentId?: string;
}

export interface MailboxConnection {
  mailboxId: string;
  teamId: string;
  authorizedByUserId: string;
  graphUserId?: string;
  emailAddress: string;
  displayName: string;
  tenantId?: string;
  inboxFolderId?: string;
  junkFolderId?: string;
  encryptedRefreshToken: string;
  accessTokenExpiresAt: string;
  providerType: MailProviderType;
  createdAt: string;
  updatedAt: string;
  status: MailboxConnectionStatus;
  lastError?: string;
}

export interface MailboxRoute {
  mailboxId: string;
  slackChannelId: string;
  slackChannelName?: string;
  updatedAt: string;
}

export interface MailboxSyncState {
  mailboxId: string;
  deltaLink?: string;
  lastSyncAt?: string;
  lastNotificationAt?: string;
  lastMessageReceivedAt?: string;
  folderStates?: Partial<Record<MailFolderKind, MailboxFolderSyncState>>;
  lastError?: string;
  updatedAt: string;
}

export interface MailboxSubscriptionLease {
  mailboxId: string;
  subscriptionId?: string;
  resource: string;
  clientState: string;
  expiresAt?: string;
  status: LeaseStatus;
  updatedAt: string;
  lastError?: string;
}

export interface DeliveredMailRecord {
  mailboxId: string;
  dedupeKey: string;
  messageId: string;
  internetMessageId?: string;
  subject: string;
  slackChannelId: string;
  deliveredAt: string;
}

export interface OAuthState {
  state: string;
  teamId: string;
  userId: string;
  channelId: string;
  channelName?: string;
  providerType: MailProviderType;
  createdAt: string;
  expiresAt: string;
}

export interface SyncJob {
  mailboxId: string;
  reason: string;
  attemptCount: number;
  enqueuedAt: string;
  requestedByUserId?: string;
}

export interface MailboxBundle {
  connection: MailboxConnection;
  route: MailboxRoute | null;
  syncState: MailboxSyncState | null;
  lease: MailboxSubscriptionLease | null;
}

export interface MailMessageSummary {
  messageId: string;
  internetMessageId?: string;
  subject: string;
  fromName?: string;
  fromAddress?: string;
  bodyPreview?: string;
  bodyText?: string;
  bodyContentType?: "text" | "html";
  receivedDateTime?: string;
  webLink?: string;
  hasAttachments?: boolean;
  attachments?: MailAttachmentSummary[];
  folderKind?: MailFolderKind;
  folderName?: string;
}
