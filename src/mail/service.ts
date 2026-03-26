import { getConfigAsync, type AppConfig } from "../config.ts";
import {
  buildMailboxMessagesResource,
  GraphApiError,
  MicrosoftGraphClient,
} from "../microsoft/graph.ts";
import {
  buildMicrosoftAuthorizeUrl,
  exchangeAuthorizationCode,
  refreshAccessToken,
  type MicrosoftTokenSet,
} from "../microsoft/oauth.ts";
import type { GraphWebhookNotification } from "../microsoft/webhook.ts";
import { fetchMsOauth2ApiMessages, MsOauth2ApiError } from "../providers/msoauth2api.ts";
import { postChannelMessage } from "../slack/api.ts";
import { buildMailNotificationBlocks } from "../slack/ui.ts";
import { getKv } from "../store/kv.ts";
import {
  deleteMailbox,
  deleteOAuthState,
  deleteSyncJob,
  enqueueSyncJob,
  findMailboxIdByEmail,
  getMailboxBundle,
  getMailboxIdBySubscription,
  getOAuthState,
  listMailboxBundles,
  listSyncJobs,
  markSyncJobAttempt,
  resolveMailboxBundle,
  saveDeliveredRecord,
  saveMailboxBundle,
  saveMailboxRoute,
  saveMailboxSyncState,
  saveOAuthState,
  hasDeliveredRecord,
} from "../store/mailbox.ts";
import { decryptSecret, encryptSecret } from "./crypto.ts";
import { buildDedupeKey, formatFolderLabel, toPreviewText } from "./message.ts";
import type {
  MailFolderKind,
  MailboxBundle,
  MailboxConnection,
  MailboxFolderSyncState,
  MailboxRoute,
  MailboxSubscriptionLease,
  MailboxSyncState,
  MailMessageSummary,
  MailProviderType,
} from "./types.ts";

function nowIso(): string {
  return new Date().toISOString();
}

function isExpired(iso: string | undefined, marginMs = 0): boolean {
  if (!iso) return true;
  const date = new Date(iso);
  if (Number.isNaN(date.getTime())) return true;
  return date.getTime() <= Date.now() + marginMs;
}

function compareIso(left: string | undefined, right: string | undefined): number | null {
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

function getFolderId(connection: MailboxConnection, kind: MailFolderKind): string | undefined {
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
  for (const [key, value] of Object.entries(folderStates) as Array<
    [MailFolderKind, MailboxFolderSyncState | undefined]
  >) {
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
  folderStates: Partial<Record<MailFolderKind, MailboxFolderSyncState>> | undefined,
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
): Promise<{ connection: MailboxConnection; folders: ResolvedMailboxFolder[] }> {
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
  const maxMinutes = Math.max(1, Math.min(config.graphSubscriptionMaxMinutes, 4230));
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
    encryptedRefreshToken: await encryptSecret(nextRefresh, config.tokenEncryptionKey),
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
  const { tokenSet, encryptedRefreshToken } = await issueAccessToken(
    config,
    bundle.connection,
    fetchImpl,
  );

  const connection: MailboxConnection = {
    ...bundle.connection,
    encryptedRefreshToken,
    accessTokenExpiresAt: tokenSet.expiresAt,
    updatedAt: nowIso(),
    status: "active",
    lastError: undefined,
  };

  return {
    graph: new MicrosoftGraphClient(config, tokenSet.accessToken, fetchImpl),
    tokenSet,
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
      slackChannelId: input.route?.slackChannelId ?? "",
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
  // 防止后续第一次轮询把历史邮件整批推到 Slack。
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

  const graph = new MicrosoftGraphClient(config, tokenSet.accessToken, fetchImpl);
  const user = await graph.getCurrentUser();
  const emailAddress = user.mail || user.userPrincipalName;
  if (!emailAddress) {
    throw new Error("Microsoft account does not expose a usable mail address");
  }

  const existingId = await findMailboxIdByEmail(kv, emailAddress);
  const existingBundle = existingId ? await getMailboxBundle(kv, existingId) : null;
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
  const { connection: resolvedConnection, folders } = await resolveGraphFolders(graph, connection);

  const route: MailboxRoute = {
    mailboxId,
    slackChannelId: oauthState.channelId,
    slackChannelName: oauthState.channelName,
    updatedAt: nowIso(),
  };

  let syncState: MailboxSyncState;
  let lease: MailboxSubscriptionLease;

  if (providerType === "graph_native") {
    const baselines = await collectGraphFolderDeltas(graph, folders, null);
    syncState = buildGraphSyncState(mailboxId, existingBundle?.syncState, baselines);
    lease = await createSubscriptionForMailbox(graph, config, resolvedConnection);
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

  const bundle: MailboxBundle = { connection: resolvedConnection, route, syncState, lease };
  await saveMailboxBundle(kv, bundle);
  await deleteOAuthState(kv, state);
  return bundle;
}

export async function listMailboxes(teamId: string): Promise<MailboxBundle[]> {
  const kv = await getKv();
  return await listMailboxBundles(kv, teamId);
}

export async function updateMailboxRoute(input: {
  teamId: string;
  mailbox: string;
  channelId: string;
  channelName?: string;
}): Promise<MailboxBundle> {
  const kv = await getKv();
  const bundle = await resolveMailboxBundle(kv, input.teamId, input.mailbox);
  if (!bundle) throw new Error("Mailbox not found");
  const route: MailboxRoute = {
    mailboxId: bundle.connection.mailboxId,
    slackChannelId: input.channelId,
    slackChannelName: input.channelName,
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
      console.error("Failed to delete Graph subscription during provider switch", error);
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
  const baselines = await collectGraphFolderDeltas(graphContext.graph, folders, null);
  const syncState = buildGraphSyncState(connection.mailboxId, bundle.syncState, baselines);
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
  const text =
    `📬 [${formatFolderLabel(message.folderKind, message.folderName)}] ${message.subject || "(no subject)"} — ${message.fromName || message.fromAddress || "Unknown sender"}`;
  const blocks = buildMailNotificationBlocks(mailbox, message, maxPreviewChars);
  await postChannelMessage(mailbox.route.slackChannelId, text, blocks);
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
        (left.receivedDateTime ?? "").localeCompare(right.receivedDateTime ?? "")
      );

    for (const message of deliverableMessages) {
      const dedupeKey = buildDedupeKey(bundle.connection.mailboxId, message);
      const alreadyDelivered = await hasDeliveredRecord(kv, bundle.connection.mailboxId, dedupeKey);
      if (alreadyDelivered) {
        skipped++;
        continue;
      }
      await sendMailNotification(workingBundle, message, config.mailPreviewMaxChars);
      await saveDeliveredRecord(kv, {
        mailboxId: bundle.connection.mailboxId,
        dedupeKey,
        messageId: message.messageId,
        internetMessageId: message.internetMessageId,
        subject: message.subject,
        slackChannelId: bundle.route?.slackChannelId ?? "",
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
    if (!nextBundle.lease?.subscriptionId ||
      nextBundle.lease.resource !== buildLeaseResource(nextBundle.connection)) {
      await ensureSubscriptionForBundle(nextBundle, fetchImpl);
    }

    return { delivered, skipped };
  } catch (error) {
    const kind = error instanceof GraphApiError && [401, 403].includes(error.status)
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
      const alreadyDelivered = await hasDeliveredRecord(kv, connection.mailboxId, dedupeKey);
      if (alreadyDelivered || isHistoricalMessage(bundle.syncState?.lastMessageReceivedAt, message)) {
        skipped++;
        continue;
      }

      await sendMailNotification(workingBundle, message, config.mailPreviewMaxChars);
      await saveDeliveredRecord(kv, {
        mailboxId: connection.mailboxId,
        dedupeKey,
        messageId: message.messageId,
        internetMessageId: message.internetMessageId,
        subject: message.subject,
        slackChannelId: bundle.route?.slackChannelId ?? "",
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
    const kind = error instanceof MsOauth2ApiError && [401, 403].includes(error.status)
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
  const renewalWindowMs = config.graphSubscriptionRenewalWindowMinutes * 60 * 1000;
  const expectedResource = buildLeaseResource(baseBundle.connection);
  const requiresRecreate = bundle.lease?.resource !== expectedResource;
  const requiresRenew = !bundle.lease?.subscriptionId ||
    requiresRecreate ||
    isExpired(bundle.lease.expiresAt, renewalWindowMs);
  if (!requiresRenew) {
    await persistBundle(baseBundle);
    return;
  }

  try {
    const nextExpiry = subscriptionExpiry(config);
    if (requiresRecreate && bundle.lease?.subscriptionId) {
      try {
        await graphContext.graph.deleteSubscription(bundle.lease.subscriptionId);
      } catch (error) {
        const graphError = error instanceof GraphApiError ? error : null;
        if (!graphError || ![404, 410].includes(graphError.status)) {
          throw error;
        }
      }
    }
    const renewed = bundle.lease?.subscriptionId && !requiresRecreate
      ? await graphContext.graph.renewSubscription(bundle.lease.subscriptionId, nextExpiry)
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

async function enqueueMaintenanceSyncs(): Promise<void> {
  const config = await getConfigAsync();
  const kv = await getKv();
  for await (const entry of kv.list<string>({ prefix: ["mailbox_email"] })) {
    const mailboxId = entry.value;
    if (!mailboxId) continue;
    const bundle = await getMailboxBundle(kv, mailboxId);
    if (!bundle) continue;
    const lastSyncAgeMs = bundle.syncState?.lastSyncAt
      ? Date.now() - new Date(bundle.syncState.lastSyncAt).getTime()
      : Number.POSITIVE_INFINITY;
    if (lastSyncAgeMs >= config.mailSyncPollIntervalMinutes * 60 * 1000) {
      await enqueueSyncJob(kv, {
        mailboxId,
        reason: "maintenance_poll",
      });
    }
  }
}

export async function renewExpiringSubscriptions(
  fetchImpl: typeof fetch = fetch,
): Promise<void> {
  const kv = await getKv();
  const bundles = await Promise.all(
    (await listSyncTargets()).map((mailboxId) => getMailboxBundle(kv, mailboxId)),
  );
  for (const bundle of bundles) {
    if (!bundle || bundle.connection.providerType === "ms_oauth2api") continue;
    await ensureSubscriptionForBundle(bundle, fetchImpl);
  }
}

async function listSyncTargets(): Promise<string[]> {
  const kv = await getKv();
  const mailboxIds: string[] = [];
  for await (const entry of kv.list<string>({ prefix: ["mailbox_email"] })) {
    if (entry.value) mailboxIds.push(entry.value);
  }
  return mailboxIds;
}

export async function runMaintenance(fetchImpl: typeof fetch = fetch): Promise<void> {
  await enqueueMaintenanceSyncs();
  await renewExpiringSubscriptions(fetchImpl);
  await processQueuedSyncs(10, fetchImpl);
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
    subject: "Test notification from Slack Outlook Mail Bot",
    fromName: bundle.connection.displayName,
    fromAddress: bundle.connection.emailAddress,
    bodyPreview: toPreviewText(
      `This is a test notification for ${bundle.connection.emailAddress}. New emails for this mailbox will be delivered here.`,
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
      const { graph } = await ensureGraphContext(bundle, config, input.fetchImpl ?? fetch);
      await graph.deleteSubscription(bundle.lease.subscriptionId);
    }
  } catch (error) {
    console.error("Failed to delete Graph subscription during disconnect", error);
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
    const mailboxId = await getMailboxIdBySubscription(kv, notification.subscriptionId);
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
