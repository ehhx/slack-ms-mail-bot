import { type AppConfig, getConfigAsync } from "../config.ts";
import {
  buildMailboxMessagesResource,
  GraphApiError,
  MicrosoftGraphClient,
} from "../microsoft/graph.ts";
import {
  buildMicrosoftAuthorizeUrl,
  exchangeAuthorizationCode,
  type MicrosoftTokenSet,
  refreshAccessToken,
} from "../microsoft/oauth.ts";
import type { GraphWebhookNotification } from "../microsoft/webhook.ts";
import {
  fetchMsOauth2ApiMessages,
  MsOauth2ApiError,
} from "../providers/msoauth2api.ts";
import { postLarkCard } from "../lark/api.ts";
import { buildLarkMailNotificationCard } from "../lark/ui.ts";
import { getKv, pruneExpiredState } from "../store/kv.ts";
import {
  deleteMailbox,
  deleteOAuthState,
  deleteSyncJob,
  enqueueSyncJob,
  findMailboxIdByEmail,
  getMailboxBundle,
  getMailboxIdBySubscription,
  getOAuthState,
  hasDeliveredRecord,
  listAllMailboxBundles,
  listMailboxBundles,
  listSyncJobs,
  markSyncJobAttempt,
  resolveMailboxBundle,
  saveDeliveredRecord,
  saveMailboxBundle,
  saveMailboxRoute,
  saveMailboxSyncState,
  saveOAuthState,
} from "../store/mailbox.ts";
import { decryptSecret, encryptSecret } from "./crypto.ts";
import { buildDedupeKey, formatFolderLabel, toPreviewText } from "./message.ts";
import type {
  MailboxBundle,
  MailboxConnection,
  MailboxFolderSyncState,
  MailboxRoute,
  MailboxSubscriptionLease,
  MailboxSyncState,
  MailFolderKind,
  MailInlineImage,
  MailMessageSummary,
  MailProviderType,
} from "./types.ts";

function nowIso(): string {
  return new Date().toISOString();
}

const MAX_INLINE_IMAGE_BYTES = 10 * 1024 * 1024;
const WEB_MESSAGE_LIST_LIMIT = 25;
const WEB_INLINE_IMAGE_LIMIT = 4;
const WEB_PAGE_CURSOR_PREFIX = "graph-page:";
const GRAPH_ACCESS_TOKEN_MARGIN_MS = 2 * 60 * 1000;
const WEB_LIST_CACHE_TTL_MS = 20 * 1000;
const WEB_DETAIL_CACHE_TTL_MS = 5 * 60 * 1000;
const WEB_LIST_CACHE_MAX = 24;
const WEB_DETAIL_CACHE_MAX = 40;

interface CachedGraphSession {
  accessToken: string;
  tokenSet: MicrosoftTokenSet;
  encryptedRefreshToken: string;
  connectionUpdatedAt: string;
  expiresAtMs: number;
}

interface TimedCacheValue<T> {
  value: T;
  expiresAtMs: number;
}

type WebListResult = {
  bundle: MailboxBundle;
  folder: { kind: MailFolderKind; folderId: string; folderName: string };
  messages: MailMessageSummary[];
  nextPageCursor?: string;
};

type WebDetailResult = {
  bundle: MailboxBundle;
  folderKind: MailFolderKind;
  message: MailMessageSummary;
};

const graphSessionCache = new Map<string, CachedGraphSession>();
const graphSessionRequests = new Map<string, Promise<CachedGraphSession>>();
const webListCache = new Map<string, TimedCacheValue<WebListResult>>();
const webDetailCache = new Map<string, TimedCacheValue<WebDetailResult>>();

function readTimedCache<T>(
  cache: Map<string, TimedCacheValue<T>>,
  key: string,
): T | null {
  const entry = cache.get(key);
  if (!entry) return null;
  if (entry.expiresAtMs <= Date.now()) {
    cache.delete(key);
    return null;
  }
  cache.delete(key);
  cache.set(key, entry);
  return entry.value;
}

function writeTimedCache<T>(
  cache: Map<string, TimedCacheValue<T>>,
  key: string,
  value: T,
  ttlMs: number,
  maxSize: number,
): void {
  cache.delete(key);
  cache.set(key, { value, expiresAtMs: Date.now() + ttlMs });
  while (cache.size > maxSize) {
    const oldestKey = cache.keys().next().value;
    if (typeof oldestKey !== "string") break;
    cache.delete(oldestKey);
  }
}

export function clearWebMailRuntimeCaches(): void {
  graphSessionCache.clear();
  graphSessionRequests.clear();
  webListCache.clear();
  webDetailCache.clear();
}

function isExpired(iso: string | undefined, marginMs = 0): boolean {
  if (!iso) return true;
  const date = new Date(iso);
  if (Number.isNaN(date.getTime())) return true;
  return date.getTime() <= Date.now() + marginMs;
}

function compareIso(
  left: string | undefined,
  right: string | undefined,
): number | null {
  if (!left || !right) return null;
  const leftMs = Date.parse(left);
  const rightMs = Date.parse(right);
  if (Number.isNaN(leftMs) || Number.isNaN(rightMs)) {
    return left.localeCompare(right);
  }
  return leftMs - rightMs;
}

function latestReceivedDate(
  messages: MailMessageSummary[],
  fallback?: string,
): string | undefined {
  let latest = fallback;
  for (const message of messages) {
    if (!message.receivedDateTime) continue;
    const compared = compareIso(latest, message.receivedDateTime);
    if (compared === null || compared < 0) {
      latest = message.receivedDateTime;
    }
  }
  return latest;
}

function isHistoricalMessage(
  watermark: string | undefined,
  message: MailMessageSummary,
): boolean {
  if (!watermark || !message.receivedDateTime) return false;
  const compared = compareIso(message.receivedDateTime, watermark);
  return compared !== null ? compared < 0 : false;
}

const GRAPH_WATCHED_FOLDERS = [
  { kind: "inbox", folderName: "Inbox", wellKnownName: "inbox" },
  { kind: "junk", folderName: "Junk", wellKnownName: "junkemail" },
] as const satisfies ReadonlyArray<{
  kind: MailFolderKind;
  folderName: string;
  wellKnownName: string;
}>;

interface ResolvedMailboxFolder {
  kind: MailFolderKind;
  folderId: string;
  folderName: string;
}

interface GraphFolderDeltaResult extends ResolvedMailboxFolder {
  deltaLink: string;
  messages: MailMessageSummary[];
}

function getFolderId(
  connection: MailboxConnection,
  kind: MailFolderKind,
): string | undefined {
  return kind === "junk" ? connection.junkFolderId : connection.inboxFolderId;
}

function setFolderId(
  connection: MailboxConnection,
  kind: MailFolderKind,
  folderId: string,
): MailboxConnection {
  return kind === "junk"
    ? { ...connection, junkFolderId: folderId }
    : { ...connection, inboxFolderId: folderId };
}

function cloneFolderStates(
  folderStates: MailboxSyncState["folderStates"],
): MailboxSyncState["folderStates"] {
  if (!folderStates) return undefined;
  const next: Partial<Record<MailFolderKind, MailboxFolderSyncState>> = {};
  for (
    const [key, value] of Object.entries(folderStates) as Array<
      [MailFolderKind, MailboxFolderSyncState | undefined]
    >
  ) {
    if (!value) continue;
    next[key] = { ...value };
  }
  return next;
}

function getFolderState(
  syncState: MailboxSyncState | null | undefined,
  folder: ResolvedMailboxFolder,
): MailboxFolderSyncState | undefined {
  const current = syncState?.folderStates?.[folder.kind];
  if (current) {
    return {
      ...current,
      folderId: current.folderId || folder.folderId,
      folderName: current.folderName || folder.folderName,
    };
  }
  if (folder.kind === "inbox") {
    return {
      folderId: folder.folderId,
      folderName: folder.folderName,
      deltaLink: syncState?.deltaLink,
      lastMessageReceivedAt: syncState?.lastMessageReceivedAt,
    };
  }
  return syncState?.lastMessageReceivedAt
    ? {
      folderId: folder.folderId,
      folderName: folder.folderName,
      lastMessageReceivedAt: syncState.lastMessageReceivedAt,
    }
    : undefined;
}

function buildFolderStates(
  previousSyncState: MailboxSyncState | null | undefined,
  results: GraphFolderDeltaResult[],
): Partial<Record<MailFolderKind, MailboxFolderSyncState>> {
  const next = cloneFolderStates(previousSyncState?.folderStates) ?? {};
  for (const result of results) {
    next[result.kind] = {
      folderId: result.folderId,
      folderName: result.folderName,
      deltaLink: result.deltaLink,
      lastMessageReceivedAt: latestReceivedDate(
        result.messages,
        getFolderState(previousSyncState, result)?.lastMessageReceivedAt,
      ),
    };
  }
  return next;
}

function latestFolderStateDate(
  folderStates:
    | Partial<Record<MailFolderKind, MailboxFolderSyncState>>
    | undefined,
  fallback?: string,
): string | undefined {
  let latest = fallback;
  for (const state of Object.values(folderStates ?? {})) {
    if (!state?.lastMessageReceivedAt) continue;
    const compared = compareIso(latest, state.lastMessageReceivedAt);
    if (compared === null || compared < 0) {
      latest = state.lastMessageReceivedAt;
    }
  }
  return latest;
}

function buildGraphSyncState(
  mailboxId: string,
  previousSyncState: MailboxSyncState | null | undefined,
  results: GraphFolderDeltaResult[],
): MailboxSyncState {
  const folderStates = buildFolderStates(previousSyncState, results);
  const inboxState = folderStates.inbox;
  return {
    mailboxId,
    deltaLink: inboxState?.deltaLink,
    lastSyncAt: nowIso(),
    lastNotificationAt: previousSyncState?.lastNotificationAt,
    lastMessageReceivedAt: latestFolderStateDate(
      folderStates,
      previousSyncState?.lastMessageReceivedAt,
    ),
    folderStates,
    updatedAt: nowIso(),
    lastError: undefined,
  };
}

async function resolveGraphFolders(
  graph: MicrosoftGraphClient,
  connection: MailboxConnection,
): Promise<
  { connection: MailboxConnection; folders: ResolvedMailboxFolder[] }
> {
  let nextConnection = connection;
  const folders: ResolvedMailboxFolder[] = [];

  for (const spec of GRAPH_WATCHED_FOLDERS) {
    let folderId = getFolderId(nextConnection, spec.kind);
    if (!folderId) {
      const folder = await graph.getMailFolder(spec.wellKnownName);
      folderId = folder.id;
      nextConnection = setFolderId(nextConnection, spec.kind, folder.id);
    }
    folders.push({
      kind: spec.kind,
      folderId,
      folderName: spec.folderName,
    });
  }

  return { connection: nextConnection, folders };
}

async function collectGraphFolderDeltas(
  graph: MicrosoftGraphClient,
  folders: ResolvedMailboxFolder[],
  syncState: MailboxSyncState | null | undefined,
): Promise<Array<GraphFolderDeltaResult & { hadDeltaLink: boolean }>> {
  const results: Array<GraphFolderDeltaResult & { hadDeltaLink: boolean }> = [];

  for (const folder of folders) {
    const previousState = getFolderState(syncState, folder);
    const delta = await graph.collectMessageDelta({
      folderId: folder.folderId,
      folderKind: folder.kind,
      folderName: folder.folderName,
      deltaLink: previousState?.deltaLink,
    });
    results.push({
      ...folder,
      deltaLink: delta.deltaLink,
      messages: delta.messages,
      hadDeltaLink: Boolean(previousState?.deltaLink),
    });
  }

  return results;
}

function buildLeaseResource(_connection: MailboxConnection): string {
  return buildMailboxMessagesResource();
}

function buildMissingLease(
  connection: MailboxConnection,
  config: AppConfig,
): MailboxSubscriptionLease {
  return {
    mailboxId: connection.mailboxId,
    resource: buildLeaseResource(connection),
    clientState: config.webhookClientState,
    status: "missing",
    updatedAt: nowIso(),
    lastError: undefined,
  };
}

function subscriptionExpiry(config: AppConfig): string {
  // Outlook message subscriptions 当前仍受约 3 天上限约束，因此这里继续限制在 4230 分钟内。
  const maxMinutes = Math.max(
    1,
    Math.min(config.graphSubscriptionMaxMinutes, 4230),
  );
  return new Date(Date.now() + maxMinutes * 60 * 1000).toISOString();
}

async function issueAccessToken(
  config: AppConfig,
  connection: MailboxConnection,
  fetchImpl: typeof fetch = fetch,
): Promise<{ tokenSet: MicrosoftTokenSet; encryptedRefreshToken: string }> {
  const refreshToken = await decryptSecret(
    connection.encryptedRefreshToken,
    config.tokenEncryptionKey,
  );
  const tokenSet = await refreshAccessToken(config, refreshToken, fetchImpl);
  const nextRefresh = tokenSet.refreshToken ?? refreshToken;
  return {
    tokenSet,
    encryptedRefreshToken: await encryptSecret(
      nextRefresh,
      config.tokenEncryptionKey,
    ),
  };
}

function buildNotificationUrl(config: AppConfig): string {
  return new URL("/graph/webhook", config.appBaseUrl).toString();
}

async function ensureGraphContext(
  bundle: MailboxBundle,
  config: AppConfig,
  fetchImpl: typeof fetch = fetch,
): Promise<{
  graph: MicrosoftGraphClient;
  tokenSet: MicrosoftTokenSet;
  connection: MailboxConnection;
}> {
  let session: CachedGraphSession | null = null;

  // 自定义 fetch 通常来自测试；避免测试间共享运行时 token，也让 mock 调用保持可预测。
  if (fetchImpl === fetch) {
    const cached = graphSessionCache.get(bundle.connection.mailboxId);
    if (
      cached && cached.expiresAtMs > Date.now() + GRAPH_ACCESS_TOKEN_MARGIN_MS
    ) {
      session = cached;
    } else {
      let pending = graphSessionRequests.get(bundle.connection.mailboxId);
      if (!pending) {
        pending = (async () => {
          const { tokenSet, encryptedRefreshToken } = await issueAccessToken(
            config,
            bundle.connection,
            fetchImpl,
          );
          const created: CachedGraphSession = {
            accessToken: tokenSet.accessToken,
            tokenSet,
            encryptedRefreshToken,
            connectionUpdatedAt: nowIso(),
            expiresAtMs: Date.parse(tokenSet.expiresAt),
          };
          graphSessionCache.set(bundle.connection.mailboxId, created);
          return created;
        })().finally(() => {
          graphSessionRequests.delete(bundle.connection.mailboxId);
        });
        graphSessionRequests.set(bundle.connection.mailboxId, pending);
      }
      session = await pending;
    }
  }

  if (!session) {
    const { tokenSet, encryptedRefreshToken } = await issueAccessToken(
      config,
      bundle.connection,
      fetchImpl,
    );
    session = {
      accessToken: tokenSet.accessToken,
      tokenSet,
      encryptedRefreshToken,
      connectionUpdatedAt: nowIso(),
      expiresAtMs: Date.parse(tokenSet.expiresAt),
    };
  }

  const connection: MailboxConnection = {
    ...bundle.connection,
    encryptedRefreshToken: session.encryptedRefreshToken,
    accessTokenExpiresAt: session.tokenSet.expiresAt,
    updatedAt: session.connectionUpdatedAt,
    status: "active",
    lastError: undefined,
  };

  return {
    graph: new MicrosoftGraphClient(config, session.accessToken, fetchImpl),
    tokenSet: session.tokenSet,
    connection,
  };
}

async function persistBundle(bundle: MailboxBundle): Promise<void> {
  const kv = await getKv();
  await saveMailboxBundle(kv, bundle);
}

async function updateBundleWithError(
  bundle: MailboxBundle,
  error: unknown,
  kind: "connection" | "lease" | "sync" = "connection",
): Promise<void> {
  const message = error instanceof Error ? error.message : String(error);
  const next: MailboxBundle = {
    ...bundle,
    connection: {
      ...bundle.connection,
      updatedAt: nowIso(),
      ...(kind === "connection"
        ? {
          status: "needs_reauth" as const,
          lastError: message,
        }
        : {}),
    },
    lease: bundle.lease
      ? {
        ...bundle.lease,
        updatedAt: nowIso(),
        ...(kind === "lease"
          ? { status: "degraded" as const, lastError: message }
          : {}),
      }
      : bundle.lease,
    syncState: bundle.syncState
      ? {
        ...bundle.syncState,
        updatedAt: nowIso(),
        ...(kind === "sync" ? { lastError: message } : {}),
      }
      : bundle.syncState,
  };
  await persistBundle(next);
}

async function seedDeliveredMessages(
  kv: Deno.Kv,
  input: {
    connection: MailboxConnection;
    route: MailboxRoute | null;
    messages: MailMessageSummary[];
  },
): Promise<void> {
  for (const message of input.messages) {
    const dedupeKey = buildDedupeKey(input.connection.mailboxId, message);
    await saveDeliveredRecord(kv, {
      mailboxId: input.connection.mailboxId,
      dedupeKey,
      messageId: message.messageId,
      internetMessageId: message.internetMessageId,
      subject: message.subject,
      deliveryChatId: input.route?.chatId ?? "",
      deliveredAt: nowIso(),
    });
  }
}

async function buildMsOauth2ApiBaselineState(
  kv: Deno.Kv,
  input: {
    config: AppConfig;
    connection: MailboxConnection;
    route: MailboxRoute | null;
    previousSyncState?: MailboxSyncState | null;
    fetchImpl?: typeof fetch;
  },
): Promise<MailboxSyncState> {
  const refreshToken = await decryptSecret(
    input.connection.encryptedRefreshToken,
    input.config.tokenEncryptionKey,
  );
  const messages = await fetchMsOauth2ApiMessages({
    config: input.config,
    refreshToken,
    emailAddress: input.connection.emailAddress,
    fetchImpl: input.fetchImpl,
  });

  // msOauth2api 只有全量拉取接口。这里在建立/切换 provider 时先把当前可见消息做基线入库，
  // 防止后续第一次轮询把历史邮件整批推到 Lark。
  await seedDeliveredMessages(kv, {
    connection: input.connection,
    route: input.route,
    messages,
  });

  return {
    mailboxId: input.connection.mailboxId,
    lastSyncAt: nowIso(),
    lastNotificationAt: input.previousSyncState?.lastNotificationAt,
    lastMessageReceivedAt: latestReceivedDate(
      messages,
      input.previousSyncState?.lastMessageReceivedAt,
    ),
    updatedAt: nowIso(),
    lastError: undefined,
  };
}

function normalizeProviderType(
  providerType: MailProviderType | undefined,
  config: AppConfig,
): MailProviderType {
  return providerType ?? config.mailProviderDefault;
}

export async function createConnectUrl(input: {
  teamId: string;
  userId: string;
  channelId: string;
  channelName?: string;
  providerType?: MailProviderType;
}): Promise<{ authorizeUrl: string; providerType: MailProviderType }> {
  const config = await getConfigAsync();
  const kv = await getKv();
  const state = crypto.randomUUID();
  const providerType = normalizeProviderType(input.providerType, config);

  await saveOAuthState(kv, {
    state,
    teamId: input.teamId,
    userId: input.userId,
    channelId: input.channelId,
    channelName: input.channelName,
    providerType,
  });

  return {
    authorizeUrl: buildMicrosoftAuthorizeUrl(config, state),
    providerType,
  };
}

async function createSubscriptionForMailbox(
  graph: MicrosoftGraphClient,
  config: AppConfig,
  connection: MailboxConnection,
): Promise<MailboxSubscriptionLease> {
  const created = await graph.createSubscription({
    resource: buildLeaseResource(connection),
    notificationUrl: buildNotificationUrl(config),
    lifecycleNotificationUrl: buildNotificationUrl(config),
    clientState: config.webhookClientState,
    expirationDateTime: subscriptionExpiry(config),
  });
  return {
    mailboxId: connection.mailboxId,
    resource: created.resource,
    clientState: config.webhookClientState,
    subscriptionId: created.id,
    expiresAt: created.expirationDateTime,
    status: "active",
    updatedAt: nowIso(),
  };
}

export async function completeOAuthCallback(
  code: string,
  state: string,
  fetchImpl: typeof fetch = fetch,
): Promise<MailboxBundle> {
  const config = await getConfigAsync();
  const kv = await getKv();
  const oauthState = await getOAuthState(kv, state);
  if (!oauthState) throw new Error("OAuth state not found or expired");
  if (isExpired(oauthState.expiresAt)) {
    await deleteOAuthState(kv, state);
    throw new Error("OAuth state expired");
  }

  const tokenSet = await exchangeAuthorizationCode(config, code, fetchImpl);
  if (!tokenSet.refreshToken) {
    throw new Error("Microsoft OAuth response did not include a refresh token");
  }

  const graph = new MicrosoftGraphClient(
    config,
    tokenSet.accessToken,
    fetchImpl,
  );
  const user = await graph.getCurrentUser();
  const emailAddress = user.mail || user.userPrincipalName;
  if (!emailAddress) {
    throw new Error("Microsoft account does not expose a usable mail address");
  }

  const existingId = await findMailboxIdByEmail(kv, emailAddress);
  const existingBundle = existingId
    ? await getMailboxBundle(kv, existingId)
    : null;
  const mailboxId = existingBundle?.connection.mailboxId ?? crypto.randomUUID();
  const encryptedRefreshToken = await encryptSecret(
    tokenSet.refreshToken,
    config.tokenEncryptionKey,
  );
  const providerType = normalizeProviderType(
    oauthState.providerType ?? existingBundle?.connection.providerType,
    config,
  );

  const connection: MailboxConnection = {
    mailboxId,
    teamId: oauthState.teamId,
    authorizedByUserId: oauthState.userId,
    graphUserId: user.id,
    emailAddress,
    displayName: user.displayName || emailAddress,
    encryptedRefreshToken,
    accessTokenExpiresAt: tokenSet.expiresAt,
    providerType,
    createdAt: existingBundle?.connection.createdAt ?? nowIso(),
    updatedAt: nowIso(),
    status: "active",
    lastError: undefined,
  };
  const { connection: resolvedConnection, folders } = await resolveGraphFolders(
    graph,
    connection,
  );

  const route: MailboxRoute = {
    mailboxId,
    platform: "lark",
    chatId: oauthState.channelId,
    chatName: oauthState.channelName,
    updatedAt: nowIso(),
  };

  let syncState: MailboxSyncState;
  let lease: MailboxSubscriptionLease;

  if (providerType === "graph_native") {
    const baselines = await collectGraphFolderDeltas(graph, folders, null);
    syncState = buildGraphSyncState(
      mailboxId,
      existingBundle?.syncState,
      baselines,
    );
    lease = await createSubscriptionForMailbox(
      graph,
      config,
      resolvedConnection,
    );
  } else {
    syncState = await buildMsOauth2ApiBaselineState(kv, {
      config,
      connection: resolvedConnection,
      route,
      previousSyncState: existingBundle?.syncState,
      fetchImpl,
    });
    lease = buildMissingLease(resolvedConnection, config);
  }

  const bundle: MailboxBundle = {
    connection: resolvedConnection,
    route,
    syncState,
    lease,
  };
  await saveMailboxBundle(kv, bundle);
  await deleteOAuthState(kv, state);
  return bundle;
}

export async function listMailboxes(teamId: string): Promise<MailboxBundle[]> {
  const kv = await getKv();
  return await listMailboxBundles(kv, teamId);
}

/**
 * 将旧 Slack 工作区的既有 Outlook 连接认领到当前 Lark 租户和群聊。
 * OAuth refresh token、Graph subscription 和同步游标都会保留，因此无需重复授权。
 */
export async function claimMailboxForLark(input: {
  teamId: string;
  userId: string;
  chatId: string;
  chatName?: string;
  mailbox: string;
}): Promise<MailboxBundle> {
  const kv = await getKv();
  const directId = await findMailboxIdByEmail(kv, input.mailbox);
  const directBundle = directId ? await getMailboxBundle(kv, directId) : null;
  const bundle = directBundle ??
    (await listAllMailboxBundles(kv)).find((item) =>
      item.connection.mailboxId.startsWith(input.mailbox.trim()) ||
      item.connection.emailAddress.toLowerCase() ===
        input.mailbox.trim().toLowerCase()
    ) ?? null;
  if (!bundle) throw new Error("Mailbox not found");

  if (
    bundle.route?.platform === "lark" &&
    bundle.connection.teamId !== input.teamId
  ) {
    throw new Error("Mailbox is already claimed by another Lark tenant");
  }

  const next: MailboxBundle = {
    ...bundle,
    connection: {
      ...bundle.connection,
      teamId: input.teamId,
      authorizedByUserId: input.userId,
      updatedAt: nowIso(),
    },
    route: {
      mailboxId: bundle.connection.mailboxId,
      platform: "lark",
      chatId: input.chatId,
      chatName: input.chatName,
      updatedAt: nowIso(),
    },
  };
  await saveMailboxBundle(kv, next);
  return next;
}

function resolveFolderKind(
  input: MailFolderKind | string | undefined,
): MailFolderKind {
  return input === "junk" ? "junk" : "inbox";
}

function encodeGraphPathSegment(value: string): string {
  return encodeURIComponent(value).replace(/%2F/g, "/");
}

function encodeOpaqueCursor(input: string): string {
  const bytes = new TextEncoder().encode(input);
  let binary = "";
  for (const byte of bytes) {
    binary += String.fromCharCode(byte);
  }
  return btoa(binary).replace(/\+/g, "-").replace(/\//g, "_").replace(
    /=+$/g,
    "",
  );
}

function decodeOpaqueCursor(input: string): string | null {
  if (!input) return null;
  try {
    const normalized = input.replace(/-/g, "+").replace(/_/g, "/");
    const padded = normalized + "=".repeat((4 - normalized.length % 4) % 4);
    const binary = atob(padded);
    const bytes = Uint8Array.from(binary, (char) => char.charCodeAt(0));
    return new TextDecoder().decode(bytes);
  } catch {
    return null;
  }
}

function buildGraphFolderMessagesPath(
  config: AppConfig,
  folderId: string,
): string {
  const graphBase = new URL(config.graphApiBaseUrl);
  const basePath = graphBase.pathname.replace(/\/+$/, "");
  return `${basePath}/me/mailFolders/${
    encodeGraphPathSegment(folderId)
  }/messages`;
}

function assertSafeGraphPageUrl(
  config: AppConfig,
  folderId: string,
  pageUrl: string,
): string {
  const graphBase = new URL(config.graphApiBaseUrl);
  const parsed = new URL(pageUrl);
  if (parsed.origin !== graphBase.origin) {
    throw new InvalidWebMailPageCursorError("分页游标域名无效。");
  }
  if (parsed.pathname !== buildGraphFolderMessagesPath(config, folderId)) {
    throw new InvalidWebMailPageCursorError("分页游标路径无效。");
  }
  if (
    !parsed.searchParams.has("$skiptoken") && !parsed.searchParams.has("$skip")
  ) {
    throw new InvalidWebMailPageCursorError("分页游标缺少分页参数。");
  }
  return parsed.toString();
}

export class InvalidWebMailPageCursorError extends Error {}

export function encodeWebMailPageCursor(pageUrl: string): string {
  return `${WEB_PAGE_CURSOR_PREFIX}${encodeOpaqueCursor(pageUrl)}`;
}

export function decodeWebMailPageCursor(
  cursor: string,
  config: AppConfig,
  folderId: string,
): string {
  if (!cursor.startsWith(WEB_PAGE_CURSOR_PREFIX)) {
    throw new InvalidWebMailPageCursorError("分页游标前缀无效。");
  }
  const decoded = decodeOpaqueCursor(
    cursor.slice(WEB_PAGE_CURSOR_PREFIX.length),
  );
  if (!decoded) {
    throw new InvalidWebMailPageCursorError("分页游标无法解码。");
  }
  return assertSafeGraphPageUrl(config, folderId, decoded);
}

function requireResolvedFolder(
  folders: ResolvedMailboxFolder[],
  kind: MailFolderKind,
): ResolvedMailboxFolder {
  const folder = folders.find((entry) => entry.kind === kind);
  if (!folder) {
    throw new Error(`Mail folder not found: ${kind}`);
  }
  return folder;
}

function connectionChanged(
  previous: MailboxConnection,
  next: MailboxConnection,
): boolean {
  return previous.encryptedRefreshToken !== next.encryptedRefreshToken ||
    previous.accessTokenExpiresAt !== next.accessTokenExpiresAt ||
    previous.inboxFolderId !== next.inboxFolderId ||
    previous.junkFolderId !== next.junkFolderId ||
    previous.status !== next.status ||
    previous.lastError !== next.lastError ||
    previous.updatedAt !== next.updatedAt;
}

async function getGraphMailboxAccessForRead(
  mailboxId: string,
  fetchImpl: typeof fetch = fetch,
): Promise<{
  kv: Deno.Kv;
  config: AppConfig;
  bundle: MailboxBundle;
  graph: MicrosoftGraphClient;
  folders: ResolvedMailboxFolder[];
}> {
  const config = await getConfigAsync();
  const kv = await getKv();
  const bundle = await getMailboxBundle(kv, mailboxId);
  if (!bundle) throw new Error("Mailbox not found");
  if (bundle.connection.providerType !== "graph_native") {
    throw new Error("Web console 当前只支持 Graph Native 邮箱");
  }

  let graphContext;
  try {
    graphContext = await ensureGraphContext(bundle, config, fetchImpl);
  } catch (error) {
    await updateBundleWithError(bundle, error, "connection");
    throw error;
  }
  const { connection, folders } = await resolveGraphFolders(
    graphContext.graph,
    graphContext.connection,
  );
  const nextBundle: MailboxBundle = {
    ...bundle,
    connection,
  };
  // 同一个边缘实例内会复用 access token。仅在 token、文件夹或状态真正变化时写 KV，
  // 避免每次切换邮件都产生一次远程 KV 写入。
  if (connectionChanged(bundle.connection, connection)) {
    await saveMailboxBundle(kv, nextBundle);
  }

  return {
    kv,
    config,
    bundle: nextBundle,
    graph: graphContext.graph,
    folders,
  };
}

export async function listAllMailboxBundlesForWeb(): Promise<MailboxBundle[]> {
  const kv = await getKv();
  return await listAllMailboxBundles(kv);
}

export async function loadMailboxWebView(input: {
  mailboxId: string;
  folderKind?: MailFolderKind | string;
  messageId?: string | null;
  limit?: number;
  pageCursor?: string | null;
  fetchImpl?: typeof fetch;
}): Promise<{
  bundle: MailboxBundle;
  folder: { kind: MailFolderKind; folderId: string; folderName: string };
  messages: MailMessageSummary[];
  selectedMessage: MailMessageSummary | null;
  nextPageCursor?: string;
}> {
  const { bundle, config, graph, folders } = await getGraphMailboxAccessForRead(
    input.mailboxId,
    input.fetchImpl,
  );
  const folder = requireResolvedFolder(
    folders,
    resolveFolderKind(input.folderKind),
  );
  const pageUrl = input.pageCursor
    ? decodeWebMailPageCursor(input.pageCursor, config, folder.folderId)
    : undefined;
  const page = await graph.listFolderMessages({
    folderId: folder.folderId,
    folderKind: folder.kind,
    folderName: folder.folderName,
    top: input.limit ?? WEB_MESSAGE_LIST_LIMIT,
    pageUrl,
  });

  let selectedMessage: MailMessageSummary | null = null;
  if (input.messageId) {
    const baseMessage = page.messages.find((message) =>
      message.messageId === input.messageId
    ) ?? {
      messageId: input.messageId,
      subject: "(loading)",
      folderKind: folder.kind,
      folderName: folder.folderName,
    };
    selectedMessage = await enrichGraphMessage(
      graph,
      baseMessage,
      WEB_INLINE_IMAGE_LIMIT,
    );
  }

  return {
    bundle,
    folder,
    messages: page.messages,
    selectedMessage,
    nextPageCursor: page.nextPageUrl
      ? encodeWebMailPageCursor(page.nextPageUrl)
      : undefined,
  };
}

export async function listMailboxMessagesForWeb(input: {
  mailboxId: string;
  folderKind?: MailFolderKind | string;
  limit?: number;
  pageCursor?: string | null;
  fetchImpl?: typeof fetch;
  forceRefresh?: boolean;
}): Promise<WebListResult> {
  const folderKind = resolveFolderKind(input.folderKind);
  const cacheKey = JSON.stringify([
    input.mailboxId,
    folderKind,
    input.limit ?? WEB_MESSAGE_LIST_LIMIT,
    input.pageCursor ?? "",
  ]);
  const useRuntimeCache = (input.fetchImpl ?? fetch) === fetch;
  if (useRuntimeCache && !input.forceRefresh) {
    const cached = readTimedCache(webListCache, cacheKey);
    if (cached) return cached;
  }

  const page = await loadMailboxWebView(input);
  const result: WebListResult = {
    bundle: page.bundle,
    folder: page.folder,
    messages: page.messages,
    nextPageCursor: page.nextPageCursor,
  };
  if (useRuntimeCache) {
    writeTimedCache(
      webListCache,
      cacheKey,
      result,
      WEB_LIST_CACHE_TTL_MS,
      WEB_LIST_CACHE_MAX,
    );
  }
  return result;
}

export async function getMailboxMessageForWeb(input: {
  mailboxId: string;
  messageId: string;
  folderKind?: MailFolderKind | string;
  fetchImpl?: typeof fetch;
}): Promise<WebDetailResult> {
  const folderKind = resolveFolderKind(input.folderKind);
  const cacheKey = `${input.mailboxId}:${folderKind}:${input.messageId}`;
  const useRuntimeCache = (input.fetchImpl ?? fetch) === fetch;
  if (useRuntimeCache) {
    const cached = readTimedCache(webDetailCache, cacheKey);
    if (cached) return cached;
  }

  const { bundle, graph } = await getGraphMailboxAccessForRead(
    input.mailboxId,
    input.fetchImpl,
  );
  const message = await enrichGraphMessage(
    graph,
    {
      messageId: input.messageId,
      subject: "(loading)",
      folderKind,
      folderName: formatFolderLabel(folderKind),
    },
    WEB_INLINE_IMAGE_LIMIT,
  );

  const result: WebDetailResult = {
    bundle,
    folderKind,
    message,
  };
  if (useRuntimeCache) {
    writeTimedCache(
      webDetailCache,
      cacheKey,
      result,
      WEB_DETAIL_CACHE_TTL_MS,
      WEB_DETAIL_CACHE_MAX,
    );
  }
  return result;
}

export async function updateMailboxRoute(input: {
  teamId: string;
  mailbox: string;
  chatId: string;
  chatName?: string;
}): Promise<MailboxBundle> {
  const kv = await getKv();
  const bundle = await resolveMailboxBundle(kv, input.teamId, input.mailbox);
  if (!bundle) throw new Error("Mailbox not found");
  const route: MailboxRoute = {
    mailboxId: bundle.connection.mailboxId,
    platform: "lark",
    chatId: input.chatId,
    chatName: input.chatName,
    updatedAt: nowIso(),
  };
  await saveMailboxRoute(kv, route);
  return { ...bundle, route };
}

export async function updateMailboxProvider(input: {
  teamId: string;
  mailbox: string;
  providerType: MailProviderType;
  fetchImpl?: typeof fetch;
}): Promise<MailboxBundle> {
  const config = await getConfigAsync();
  const kv = await getKv();
  const bundle = await resolveMailboxBundle(kv, input.teamId, input.mailbox);
  if (!bundle) throw new Error("Mailbox not found");
  if (bundle.connection.providerType === input.providerType) {
    return bundle;
  }

  const fetchImpl = input.fetchImpl ?? fetch;
  const baseConnection: MailboxConnection = {
    ...bundle.connection,
    providerType: input.providerType,
    updatedAt: nowIso(),
    status: "active",
    lastError: undefined,
  };

  if (input.providerType === "ms_oauth2api") {
    try {
      if (bundle.lease?.subscriptionId) {
        const { graph } = await ensureGraphContext(bundle, config, fetchImpl);
        await graph.deleteSubscription(bundle.lease.subscriptionId);
      }
    } catch (error) {
      console.error(
        "Failed to delete Graph subscription during provider switch",
        error,
      );
    }

    const syncState = await buildMsOauth2ApiBaselineState(kv, {
      config,
      connection: baseConnection,
      route: bundle.route,
      previousSyncState: bundle.syncState,
      fetchImpl,
    });

    const nextBundle: MailboxBundle = {
      ...bundle,
      connection: baseConnection,
      syncState,
      lease: buildMissingLease(baseConnection, config),
    };
    await saveMailboxBundle(kv, nextBundle);
    return nextBundle;
  }

  let graphContext;
  try {
    graphContext = await ensureGraphContext(
      {
        ...bundle,
        connection: baseConnection,
      },
      config,
      fetchImpl,
    );
  } catch (error) {
    await updateBundleWithError(
      {
        ...bundle,
        connection: baseConnection,
      },
      error,
      "connection",
    );
    throw error;
  }

  const { connection, folders } = await resolveGraphFolders(
    graphContext.graph,
    graphContext.connection,
  );
  const baselines = await collectGraphFolderDeltas(
    graphContext.graph,
    folders,
    null,
  );
  const syncState = buildGraphSyncState(
    connection.mailboxId,
    bundle.syncState,
    baselines,
  );
  const lease = await createSubscriptionForMailbox(
    graphContext.graph,
    config,
    connection,
  );
  const nextBundle: MailboxBundle = {
    ...bundle,
    connection,
    syncState,
    lease,
  };
  await saveMailboxBundle(kv, nextBundle);
  return nextBundle;
}

export async function enqueueMailboxSync(input: {
  mailboxId: string;
  reason: string;
  requestedByUserId?: string;
}): Promise<void> {
  const kv = await getKv();
  await enqueueSyncJob(kv, input);
}

export async function queueMailboxSyncByMailboxRef(input: {
  teamId: string;
  mailbox: string;
  reason: string;
  requestedByUserId?: string;
}): Promise<MailboxBundle> {
  const kv = await getKv();
  const bundle = await resolveMailboxBundle(kv, input.teamId, input.mailbox);
  if (!bundle) throw new Error("Mailbox not found");
  await enqueueSyncJob(kv, {
    mailboxId: bundle.connection.mailboxId,
    reason: input.reason,
    requestedByUserId: input.requestedByUserId,
  });
  return bundle;
}

async function sendMailNotification(
  mailbox: MailboxBundle,
  message: MailMessageSummary,
  maxPreviewChars: number,
): Promise<void> {
  if (!mailbox.route) throw new Error("Mailbox route is not configured");
  if (mailbox.route.platform !== "lark") {
    throw new Error(
      "Mailbox route has not been migrated to Lark; run `mail claim <mailbox>` in the target group",
    );
  }

  await postLarkCard({
    chatId: mailbox.route.chatId,
    card: buildLarkMailNotificationCard(mailbox, message, maxPreviewChars),
    idempotencyKey: `${mailbox.connection.mailboxId}:${
      buildDedupeKey(mailbox.connection.mailboxId, message)
    }`,
  });
}

async function loadInlineImagesForGraphMessage(
  graph: MicrosoftGraphClient,
  message: MailMessageSummary,
  maxItems: number,
): Promise<MailInlineImage[]> {
  const inlineAttachments = (message.attachments ?? [])
    .filter((attachment) =>
      Boolean(
        attachment.attachmentId &&
          attachment.isInline &&
          attachment.contentType?.startsWith("image/") &&
          (!attachment.size || attachment.size <= MAX_INLINE_IMAGE_BYTES),
      )
    )
    .slice(0, Math.max(0, maxItems));
  const inlineImages: MailInlineImage[] = [];
  for (const attachment of inlineAttachments) {
    try {
      const content = await graph.getInlineImageAttachmentContent(
        message.messageId,
        attachment.attachmentId!,
      );
      if (content) {
        inlineImages.push(content);
      }
    } catch (error) {
      console.error("Failed to read inline image attachment", {
        messageId: message.messageId,
        attachmentId: attachment.attachmentId,
        error,
      });
    }
  }
  return inlineImages;
}

function shouldLoadInlineImagesForMessage(
  message: MailMessageSummary,
  inlineImageLimit: number,
): boolean {
  if (inlineImageLimit <= 0) return false;
  if (
    !message.attachments?.some((attachment) =>
      attachment.isInline && attachment.contentType?.startsWith("image/")
    )
  ) {
    return false;
  }
  if (message.bodyContentType !== "html") return false;
  return message.bodyText?.toLowerCase().includes("cid:") ?? false;
}

async function enrichGraphMessage(
  graph: MicrosoftGraphClient,
  message: MailMessageSummary,
  inlineImageLimit: number,
): Promise<MailMessageSummary> {
  const detail = await graph.getMessageDetail(message.messageId);
  const merged: MailMessageSummary = {
    ...message,
    ...detail,
    folderKind: message.folderKind,
    folderName: message.folderName,
  };
  const inlineImages =
    shouldLoadInlineImagesForMessage(merged, inlineImageLimit)
      ? await loadInlineImagesForGraphMessage(
        graph,
        merged,
        inlineImageLimit,
      )
      : [];
  return {
    ...merged,
    inlineImages,
  };
}

async function enrichGraphMessageForNotification(
  graph: MicrosoftGraphClient,
  message: MailMessageSummary,
): Promise<MailMessageSummary> {
  try {
    // Lark 首版通知只发送正文摘要和附件元数据，不上传邮件内联图片。
    return await enrichGraphMessage(graph, message, 0);
  } catch (error) {
    console.error(
      "Failed to enrich Graph message detail",
      message.messageId,
      error,
    );
    return message;
  }
}

async function syncGraphMailbox(
  bundle: MailboxBundle,
  config: AppConfig,
  kv: Deno.Kv,
  fetchImpl: typeof fetch,
): Promise<{ delivered: number; skipped: number }> {
  let graphContext;
  try {
    graphContext = await ensureGraphContext(bundle, config, fetchImpl);
  } catch (error) {
    await updateBundleWithError(bundle, error, "connection");
    throw error;
  }

  const { connection, folders } = await resolveGraphFolders(
    graphContext.graph,
    graphContext.connection,
  );
  const workingBundle: MailboxBundle = {
    ...bundle,
    connection,
  };

  try {
    const deltas = await collectGraphFolderDeltas(
      graphContext.graph,
      folders,
      bundle.syncState,
    );

    let delivered = 0;
    let skipped = 0;
    const deliverableMessages = deltas
      .flatMap((delta) => {
        if (!delta.hadDeltaLink) {
          skipped += delta.messages.length;
          return [];
        }
        return delta.messages;
      })
      .sort((left, right) =>
        (left.receivedDateTime ?? "").localeCompare(
          right.receivedDateTime ?? "",
        )
      );

    for (const message of deliverableMessages) {
      const initialDedupeKey = buildDedupeKey(
        bundle.connection.mailboxId,
        message,
      );
      const alreadyDelivered = await hasDeliveredRecord(
        kv,
        bundle.connection.mailboxId,
        initialDedupeKey,
      );
      if (alreadyDelivered) {
        skipped++;
        continue;
      }
      const enrichedMessage = await enrichGraphMessageForNotification(
        graphContext.graph,
        message,
      );
      const dedupeKey = buildDedupeKey(
        bundle.connection.mailboxId,
        enrichedMessage,
      );
      if (dedupeKey !== initialDedupeKey) {
        const deliveredAfterEnrich = await hasDeliveredRecord(
          kv,
          bundle.connection.mailboxId,
          dedupeKey,
        );
        if (deliveredAfterEnrich) {
          skipped++;
          continue;
        }
      }
      await sendMailNotification(
        workingBundle,
        enrichedMessage,
        config.mailPreviewMaxChars,
      );
      await saveDeliveredRecord(kv, {
        mailboxId: bundle.connection.mailboxId,
        dedupeKey,
        messageId: enrichedMessage.messageId,
        internetMessageId: enrichedMessage.internetMessageId,
        subject: enrichedMessage.subject,
        deliveryChatId: workingBundle.route?.chatId ?? "",
        deliveredAt: nowIso(),
      });
      delivered++;
    }

    const nextSyncState = buildGraphSyncState(
      bundle.connection.mailboxId,
      bundle.syncState,
      deltas,
    );

    const nextBundle: MailboxBundle = {
      ...workingBundle,
      syncState: nextSyncState,
      lease: workingBundle.lease
        ? { ...workingBundle.lease, lastError: undefined, updatedAt: nowIso() }
        : workingBundle.lease,
    };
    await saveMailboxBundle(kv, nextBundle);
    if (
      !nextBundle.lease?.subscriptionId ||
      nextBundle.lease.resource !== buildLeaseResource(nextBundle.connection)
    ) {
      await ensureSubscriptionForBundle(nextBundle, fetchImpl);
    }

    return { delivered, skipped };
  } catch (error) {
    const kind =
      error instanceof GraphApiError && [401, 403].includes(error.status)
        ? "connection"
        : "sync";
    await updateBundleWithError(workingBundle, error, kind);
    throw error;
  }
}

async function syncMsOauth2ApiMailbox(
  bundle: MailboxBundle,
  config: AppConfig,
  kv: Deno.Kv,
  fetchImpl: typeof fetch,
): Promise<{ delivered: number; skipped: number }> {
  const connection: MailboxConnection = {
    ...bundle.connection,
    updatedAt: nowIso(),
    status: "active",
    lastError: undefined,
  };
  const workingBundle: MailboxBundle = {
    ...bundle,
    connection,
    lease: bundle.lease ?? buildMissingLease(connection, config),
  };

  try {
    const refreshToken = await decryptSecret(
      connection.encryptedRefreshToken,
      config.tokenEncryptionKey,
    );
    const messages = await fetchMsOauth2ApiMessages({
      config,
      refreshToken,
      emailAddress: connection.emailAddress,
      fetchImpl,
    });

    let delivered = 0;
    let skipped = 0;
    for (const message of messages) {
      const dedupeKey = buildDedupeKey(connection.mailboxId, message);
      const alreadyDelivered = await hasDeliveredRecord(
        kv,
        connection.mailboxId,
        dedupeKey,
      );
      if (
        alreadyDelivered ||
        isHistoricalMessage(bundle.syncState?.lastMessageReceivedAt, message)
      ) {
        skipped++;
        continue;
      }

      await sendMailNotification(
        workingBundle,
        message,
        config.mailPreviewMaxChars,
      );
      await saveDeliveredRecord(kv, {
        mailboxId: connection.mailboxId,
        dedupeKey,
        messageId: message.messageId,
        internetMessageId: message.internetMessageId,
        subject: message.subject,
        deliveryChatId: workingBundle.route?.chatId ?? "",
        deliveredAt: nowIso(),
      });
      delivered++;
    }

    const nextSyncState: MailboxSyncState = {
      mailboxId: connection.mailboxId,
      lastSyncAt: nowIso(),
      lastNotificationAt: bundle.syncState?.lastNotificationAt,
      lastMessageReceivedAt: latestReceivedDate(
        messages,
        bundle.syncState?.lastMessageReceivedAt,
      ),
      updatedAt: nowIso(),
      lastError: undefined,
    };
    const nextLease = {
      ...(workingBundle.lease ?? buildMissingLease(connection, config)),
      resource: buildLeaseResource(connection),
      clientState: config.webhookClientState,
      subscriptionId: undefined,
      expiresAt: undefined,
      status: "missing" as const,
      updatedAt: nowIso(),
      lastError: undefined,
    };

    await saveMailboxBundle(kv, {
      ...workingBundle,
      syncState: nextSyncState,
      lease: nextLease,
    });

    return { delivered, skipped };
  } catch (error) {
    const kind =
      error instanceof MsOauth2ApiError && [401, 403].includes(error.status)
        ? "connection"
        : "sync";
    await updateBundleWithError(workingBundle, error, kind);
    throw error;
  }
}

export async function syncMailbox(
  mailboxId: string,
  fetchImpl: typeof fetch = fetch,
): Promise<{ delivered: number; skipped: number }> {
  const config = await getConfigAsync();
  const kv = await getKv();
  const bundle = await getMailboxBundle(kv, mailboxId);
  if (!bundle) throw new Error("Mailbox not found");

  if (bundle.connection.providerType === "ms_oauth2api") {
    return await syncMsOauth2ApiMailbox(bundle, config, kv, fetchImpl);
  }
  return await syncGraphMailbox(bundle, config, kv, fetchImpl);
}

export async function processQueuedSyncs(
  limit = 10,
  fetchImpl: typeof fetch = fetch,
): Promise<void> {
  const kv = await getKv();
  const jobs = await listSyncJobs(kv);
  for (const job of jobs.slice(0, limit)) {
    try {
      await markSyncJobAttempt(kv, job.mailboxId);
      await syncMailbox(job.mailboxId, fetchImpl);
      await deleteSyncJob(kv, job.mailboxId);
    } catch (error) {
      console.error("sync job failed", job.mailboxId, error);
    }
  }
}

async function ensureSubscriptionForBundle(
  bundle: MailboxBundle,
  fetchImpl: typeof fetch = fetch,
): Promise<void> {
  if (bundle.connection.providerType === "ms_oauth2api") {
    return;
  }

  const config = await getConfigAsync();
  const kv = await getKv();
  let graphContext;
  try {
    graphContext = await ensureGraphContext(bundle, config, fetchImpl);
  } catch (error) {
    await updateBundleWithError(bundle, error, "connection");
    return;
  }

  const baseBundle: MailboxBundle = {
    ...bundle,
    connection: graphContext.connection,
  };
  const renewalWindowMs = config.graphSubscriptionRenewalWindowMinutes * 60 *
    1000;
  const expectedResource = buildLeaseResource(baseBundle.connection);
  const requiresRecreate = bundle.lease?.resource !== expectedResource;
  const requiresRenew = !bundle.lease?.subscriptionId ||
    requiresRecreate ||
    isExpired(bundle.lease.expiresAt, renewalWindowMs);
  if (!requiresRenew) {
    if (connectionChanged(bundle.connection, baseBundle.connection)) {
      await persistBundle(baseBundle);
    }
    return;
  }

  try {
    const nextExpiry = subscriptionExpiry(config);
    if (requiresRecreate && bundle.lease?.subscriptionId) {
      try {
        await graphContext.graph.deleteSubscription(
          bundle.lease.subscriptionId,
        );
      } catch (error) {
        const graphError = error instanceof GraphApiError ? error : null;
        if (!graphError || ![404, 410].includes(graphError.status)) {
          throw error;
        }
      }
    }
    const renewed = bundle.lease?.subscriptionId && !requiresRecreate
      ? await graphContext.graph.renewSubscription(
        bundle.lease.subscriptionId,
        nextExpiry,
      )
      : await graphContext.graph.createSubscription({
        resource: expectedResource,
        notificationUrl: buildNotificationUrl(config),
        lifecycleNotificationUrl: buildNotificationUrl(config),
        clientState: config.webhookClientState,
        expirationDateTime: nextExpiry,
      });

    const lease: MailboxSubscriptionLease = {
      mailboxId: bundle.connection.mailboxId,
      resource: renewed.resource,
      clientState: config.webhookClientState,
      subscriptionId: renewed.id,
      expiresAt: renewed.expirationDateTime,
      status: "active",
      updatedAt: nowIso(),
      lastError: undefined,
    };

    await saveMailboxBundle(kv, { ...baseBundle, lease });
  } catch (error) {
    const graphError = error instanceof GraphApiError ? error : null;
    if (graphError && [404, 410].includes(graphError.status)) {
      try {
        const recreatedLease = await createSubscriptionForMailbox(
          graphContext.graph,
          config,
          baseBundle.connection,
        );
        await saveMailboxBundle(kv, { ...baseBundle, lease: recreatedLease });
        return;
      } catch (recreateError) {
        await updateBundleWithError(baseBundle, recreateError, "lease");
        return;
      }
    }
    await updateBundleWithError(baseBundle, error, "lease");
  }
}

async function enqueueMaintenanceSyncs(
  bundles: MailboxBundle[],
): Promise<void> {
  const config = await getConfigAsync();
  const kv = await getKv();
  for (const bundle of bundles) {
    const lastSyncAgeMs = bundle.syncState?.lastSyncAt
      ? Date.now() - new Date(bundle.syncState.lastSyncAt).getTime()
      : Number.POSITIVE_INFINITY;
    if (lastSyncAgeMs >= config.mailSyncPollIntervalMinutes * 60 * 1000) {
      await enqueueSyncJob(kv, {
        mailboxId: bundle.connection.mailboxId,
        reason: "maintenance_poll",
      });
    }
  }
}

export async function renewExpiringSubscriptions(
  fetchImpl: typeof fetch = fetch,
  existingBundles?: MailboxBundle[],
): Promise<void> {
  const kv = await getKv();
  const bundles = existingBundles ?? await listAllMailboxBundles(kv);
  for (const bundle of bundles) {
    if (bundle.connection.providerType === "ms_oauth2api") continue;
    await ensureSubscriptionForBundle(bundle, fetchImpl);
  }
}

export async function runMaintenance(
  fetchImpl: typeof fetch = fetch,
): Promise<void> {
  const kv = await getKv();
  const bundles = await listAllMailboxBundles(kv);
  await enqueueMaintenanceSyncs(bundles);
  await renewExpiringSubscriptions(fetchImpl, bundles);
  await processQueuedSyncs(10, fetchImpl);
  await pruneExpiredState(200);
}

export async function sendTestNotification(input: {
  teamId: string;
  mailbox: string;
}): Promise<MailboxBundle> {
  const config = await getConfigAsync();
  const kv = await getKv();
  const bundle = await resolveMailboxBundle(kv, input.teamId, input.mailbox);
  if (!bundle) throw new Error("Mailbox not found");
  if (!bundle.route) throw new Error("Mailbox route is not configured");
  await sendMailNotification(bundle, {
    messageId: crypto.randomUUID(),
    subject: "Test notification from Lark Outlook Mail Bot",
    fromName: bundle.connection.displayName,
    fromAddress: bundle.connection.emailAddress,
    bodyPreview: toPreviewText(
      `This is a test notification for ${bundle.connection.emailAddress}. New emails for this mailbox will be delivered to this Lark chat.`,
      config.mailPreviewMaxChars,
    ),
    receivedDateTime: nowIso(),
    webLink: new URL("https://outlook.office.com/mail/").toString(),
    folderKind: "inbox",
    folderName: "Inbox",
  }, config.mailPreviewMaxChars);
  return bundle;
}

export async function disconnectMailbox(input: {
  teamId: string;
  mailbox: string;
  fetchImpl?: typeof fetch;
}): Promise<MailboxBundle> {
  const kv = await getKv();
  const bundle = await resolveMailboxBundle(kv, input.teamId, input.mailbox);
  if (!bundle) throw new Error("Mailbox not found");

  try {
    if (bundle.lease?.subscriptionId) {
      const config = await getConfigAsync();
      const { graph } = await ensureGraphContext(
        bundle,
        config,
        input.fetchImpl ?? fetch,
      );
      await graph.deleteSubscription(bundle.lease.subscriptionId);
    }
  } catch (error) {
    console.error(
      "Failed to delete Graph subscription during disconnect",
      error,
    );
  }

  await deleteMailbox(kv, bundle.connection.mailboxId);
  return bundle;
}

export async function processGraphNotifications(
  notifications: GraphWebhookNotification[],
): Promise<{ queued: number; ignored: number }> {
  const config = await getConfigAsync();
  const kv = await getKv();
  let queued = 0;
  let ignored = 0;

  for (const notification of notifications) {
    if (notification.clientState !== config.webhookClientState) {
      ignored++;
      continue;
    }
    const mailboxId = await getMailboxIdBySubscription(
      kv,
      notification.subscriptionId,
    );
    if (!mailboxId) {
      ignored++;
      continue;
    }

    const bundle = await getMailboxBundle(kv, mailboxId);
    if (!bundle || bundle.connection.providerType === "ms_oauth2api") {
      ignored++;
      continue;
    }

    const nextSyncState: MailboxSyncState = {
      mailboxId,
      deltaLink: bundle.syncState?.deltaLink,
      lastSyncAt: bundle.syncState?.lastSyncAt,
      lastMessageReceivedAt: bundle.syncState?.lastMessageReceivedAt,
      folderStates: cloneFolderStates(bundle.syncState?.folderStates),
      lastNotificationAt: nowIso(),
      updatedAt: nowIso(),
      lastError: undefined,
    };
    await saveMailboxSyncState(kv, nextSyncState);
    await enqueueSyncJob(kv, {
      mailboxId,
      reason: notification.lifecycleEvent
        ? `graph_${notification.lifecycleEvent}`
        : (notification.changeType ?? "graph_notification"),
    });
    queued++;
  }

  return { queued, ignored };
}
