import { deleteByPrefix } from "./kv.ts";
import type {
  DeliveredMailRecord,
  MailboxBundle,
  MailboxConnection,
  MailboxRoute,
  MailboxSubscriptionLease,
  MailboxSyncState,
  OAuthState,
  SyncJob,
} from "../mail/types.ts";

const OAUTH_STATE_TTL_MS = 15 * 60 * 1000;
const DELIVERED_RECORD_TTL_MS = 90 * 24 * 60 * 60 * 1000;

function nowIso(): string {
  return new Date().toISOString();
}

function normalizeEmail(email: string): string {
  return email.trim().toLowerCase();
}

function connectionKey(mailboxId: string): Deno.KvKey {
  return ["mailbox_connection", mailboxId];
}

function emailKey(email: string): Deno.KvKey {
  return ["mailbox_email", normalizeEmail(email)];
}

function teamIndexKey(teamId: string, email: string, mailboxId: string): Deno.KvKey {
  return ["team_mailbox", teamId, normalizeEmail(email), mailboxId];
}

function routeKey(mailboxId: string): Deno.KvKey {
  return ["mailbox_route", mailboxId];
}

function syncStateKey(mailboxId: string): Deno.KvKey {
  return ["mailbox_sync", mailboxId];
}

function leaseKey(mailboxId: string): Deno.KvKey {
  return ["mailbox_lease", mailboxId];
}

function subscriptionKey(subscriptionId: string): Deno.KvKey {
  return ["subscription_mailbox", subscriptionId];
}

function deliveredPrefix(mailboxId: string): Deno.KvKey {
  return ["delivered_mail", mailboxId];
}

function deliveredKey(mailboxId: string, dedupeKey: string): Deno.KvKey {
  return ["delivered_mail", mailboxId, dedupeKey];
}

function oauthStateKey(state: string): Deno.KvKey {
  return ["oauth_state", state];
}

function syncQueueKey(mailboxId: string): Deno.KvKey {
  return ["sync_queue", mailboxId];
}

export async function saveOAuthState(
  kv: Deno.Kv,
  input: Omit<OAuthState, "createdAt" | "expiresAt"> & { expiresAt?: string },
): Promise<OAuthState> {
  const createdAt = nowIso();
  const expiresAt = input.expiresAt ?? new Date(Date.now() + OAUTH_STATE_TTL_MS).toISOString();
  const value: OAuthState = { ...input, createdAt, expiresAt };
  await kv.set(oauthStateKey(value.state), value, { expireIn: OAUTH_STATE_TTL_MS });
  return value;
}

export async function getOAuthState(
  kv: Deno.Kv,
  state: string,
): Promise<OAuthState | null> {
  const res = await kv.get<OAuthState>(oauthStateKey(state));
  return res.value ?? null;
}

export async function deleteOAuthState(kv: Deno.Kv, state: string): Promise<void> {
  await kv.delete(oauthStateKey(state));
}

export async function saveMailboxBundle(
  kv: Deno.Kv,
  bundle: MailboxBundle,
): Promise<void> {
  const existing = await getMailboxConnection(kv, bundle.connection.mailboxId);
  const atomic = kv.atomic()
    .set(connectionKey(bundle.connection.mailboxId), bundle.connection)
    .set(emailKey(bundle.connection.emailAddress), bundle.connection.mailboxId)
    .set(
      teamIndexKey(
        bundle.connection.teamId,
        bundle.connection.emailAddress,
        bundle.connection.mailboxId,
      ),
      { mailboxId: bundle.connection.mailboxId },
    );

  if (existing && normalizeEmail(existing.emailAddress) !== normalizeEmail(bundle.connection.emailAddress)) {
    atomic.delete(emailKey(existing.emailAddress));
    atomic.delete(teamIndexKey(existing.teamId, existing.emailAddress, existing.mailboxId));
  }

  if (bundle.route) atomic.set(routeKey(bundle.connection.mailboxId), bundle.route);
  if (bundle.syncState) atomic.set(syncStateKey(bundle.connection.mailboxId), bundle.syncState);

  const existingLease = await getMailboxLease(kv, bundle.connection.mailboxId);
  if (existingLease?.subscriptionId && existingLease.subscriptionId !== bundle.lease?.subscriptionId) {
    atomic.delete(subscriptionKey(existingLease.subscriptionId));
  }
  if (bundle.lease) {
    atomic.set(leaseKey(bundle.connection.mailboxId), bundle.lease);
    if (bundle.lease.subscriptionId) {
      atomic.set(subscriptionKey(bundle.lease.subscriptionId), bundle.connection.mailboxId);
    }
  }

  const result = await atomic.commit();
  if (!result.ok) throw new Error("Failed to save mailbox bundle");
}

export async function getMailboxConnection(
  kv: Deno.Kv,
  mailboxId: string,
): Promise<MailboxConnection | null> {
  const res = await kv.get<MailboxConnection>(connectionKey(mailboxId));
  return res.value ?? null;
}

export async function getMailboxRoute(
  kv: Deno.Kv,
  mailboxId: string,
): Promise<MailboxRoute | null> {
  const res = await kv.get<MailboxRoute>(routeKey(mailboxId));
  return res.value ?? null;
}

export async function getMailboxSyncState(
  kv: Deno.Kv,
  mailboxId: string,
): Promise<MailboxSyncState | null> {
  const res = await kv.get<MailboxSyncState>(syncStateKey(mailboxId));
  return res.value ?? null;
}

export async function getMailboxLease(
  kv: Deno.Kv,
  mailboxId: string,
): Promise<MailboxSubscriptionLease | null> {
  const res = await kv.get<MailboxSubscriptionLease>(leaseKey(mailboxId));
  return res.value ?? null;
}

export async function getMailboxBundle(
  kv: Deno.Kv,
  mailboxId: string,
): Promise<MailboxBundle | null> {
  const [connection, route, syncState, lease] = await kv.getMany<
    MailboxConnection | MailboxRoute | MailboxSyncState | MailboxSubscriptionLease
  >([
    connectionKey(mailboxId),
    routeKey(mailboxId),
    syncStateKey(mailboxId),
    leaseKey(mailboxId),
  ]);

  if (!connection.value) return null;
  return {
    connection: connection.value as MailboxConnection,
    route: (route.value as MailboxRoute | null) ?? null,
    syncState: (syncState.value as MailboxSyncState | null) ?? null,
    lease: (lease.value as MailboxSubscriptionLease | null) ?? null,
  };
}

export async function findMailboxIdByEmail(
  kv: Deno.Kv,
  emailAddress: string,
): Promise<string | null> {
  const res = await kv.get<string>(emailKey(emailAddress));
  return res.value ?? null;
}

export async function getMailboxIdBySubscription(
  kv: Deno.Kv,
  subscriptionId: string,
): Promise<string | null> {
  const res = await kv.get<string>(subscriptionKey(subscriptionId));
  return res.value ?? null;
}

export async function listMailboxBundles(
  kv: Deno.Kv,
  teamId: string,
): Promise<MailboxBundle[]> {
  const ids: string[] = [];
  for await (const entry of kv.list({ prefix: ["team_mailbox", teamId] })) {
    ids.push(String(entry.key[3]));
  }

  const bundles = await Promise.all(ids.map((id) => getMailboxBundle(kv, id)));
  return bundles.filter((value): value is MailboxBundle => Boolean(value));
}

export async function listAllMailboxBundles(
  kv: Deno.Kv,
): Promise<MailboxBundle[]> {
  const ids: string[] = [];
  for await (const entry of kv.list({ prefix: ["mailbox_connection"] })) {
    ids.push(String(entry.key[1]));
  }

  const bundles = await Promise.all(ids.map((id) => getMailboxBundle(kv, id)));
  return bundles
    .filter((value): value is MailboxBundle => Boolean(value))
    .sort((left, right) =>
      left.connection.emailAddress.localeCompare(right.connection.emailAddress)
    );
}

export async function resolveMailboxBundle(
  kv: Deno.Kv,
  teamId: string,
  input: string,
): Promise<MailboxBundle | null> {
  const trimmed = input.trim();
  if (!trimmed) return null;

  const byEmailId = await findMailboxIdByEmail(kv, trimmed);
  if (byEmailId) {
    const bundle = await getMailboxBundle(kv, byEmailId);
    if (bundle?.connection.teamId === teamId) return bundle;
  }

  const bundles = await listMailboxBundles(kv, teamId);
  return bundles.find((bundle) =>
    bundle.connection.mailboxId.startsWith(trimmed) ||
    normalizeEmail(bundle.connection.emailAddress) === normalizeEmail(trimmed)
  ) ?? null;
}

export async function saveMailboxRoute(
  kv: Deno.Kv,
  route: MailboxRoute,
): Promise<void> {
  await kv.set(routeKey(route.mailboxId), route);
}

export async function saveMailboxSyncState(
  kv: Deno.Kv,
  state: MailboxSyncState,
): Promise<void> {
  await kv.set(syncStateKey(state.mailboxId), state);
}

export async function saveMailboxLease(
  kv: Deno.Kv,
  lease: MailboxSubscriptionLease,
): Promise<void> {
  const existing = await getMailboxLease(kv, lease.mailboxId);
  const atomic = kv.atomic().set(leaseKey(lease.mailboxId), lease);
  if (existing?.subscriptionId && existing.subscriptionId !== lease.subscriptionId) {
    atomic.delete(subscriptionKey(existing.subscriptionId));
  }
  if (lease.subscriptionId) {
    atomic.set(subscriptionKey(lease.subscriptionId), lease.mailboxId);
  }
  const result = await atomic.commit();
  if (!result.ok) throw new Error("Failed to save mailbox lease");
}

export async function hasDeliveredRecord(
  kv: Deno.Kv,
  mailboxId: string,
  dedupeKey: string,
): Promise<boolean> {
  const res = await kv.get<DeliveredMailRecord>(deliveredKey(mailboxId, dedupeKey));
  return Boolean(res.value);
}

export async function saveDeliveredRecord(
  kv: Deno.Kv,
  record: DeliveredMailRecord,
): Promise<void> {
  await kv.set(deliveredKey(record.mailboxId, record.dedupeKey), record, {
    expireIn: DELIVERED_RECORD_TTL_MS,
  });
}

export async function enqueueSyncJob(
  kv: Deno.Kv,
  input: Omit<SyncJob, "attemptCount" | "enqueuedAt"> & { attemptCount?: number; enqueuedAt?: string },
): Promise<SyncJob> {
  const current = await kv.get<SyncJob>(syncQueueKey(input.mailboxId));
  const next: SyncJob = {
    mailboxId: input.mailboxId,
    reason: input.reason,
    requestedByUserId: input.requestedByUserId,
    attemptCount: input.attemptCount ?? (current.value?.attemptCount ?? 0),
    enqueuedAt: input.enqueuedAt ?? current.value?.enqueuedAt ?? nowIso(),
  };
  await kv.set(syncQueueKey(input.mailboxId), next);
  return next;
}

export async function markSyncJobAttempt(
  kv: Deno.Kv,
  mailboxId: string,
): Promise<void> {
  const current = await kv.get<SyncJob>(syncQueueKey(mailboxId));
  if (!current.value) return;
  await kv.set(syncQueueKey(mailboxId), {
    ...current.value,
    attemptCount: current.value.attemptCount + 1,
  });
}

export async function listSyncJobs(kv: Deno.Kv): Promise<SyncJob[]> {
  const jobs: SyncJob[] = [];
  for await (const entry of kv.list<SyncJob>({ prefix: ["sync_queue"] })) {
    if (entry.value) jobs.push(entry.value);
  }
  jobs.sort((a, b) => a.enqueuedAt.localeCompare(b.enqueuedAt));
  return jobs;
}

export async function deleteSyncJob(kv: Deno.Kv, mailboxId: string): Promise<void> {
  await kv.delete(syncQueueKey(mailboxId));
}

export async function deleteMailbox(kv: Deno.Kv, mailboxId: string): Promise<void> {
  const bundle = await getMailboxBundle(kv, mailboxId);
  if (!bundle) return;

  const atomic = kv.atomic()
    .delete(connectionKey(mailboxId))
    .delete(routeKey(mailboxId))
    .delete(syncStateKey(mailboxId))
    .delete(leaseKey(mailboxId))
    .delete(emailKey(bundle.connection.emailAddress))
    .delete(teamIndexKey(bundle.connection.teamId, bundle.connection.emailAddress, mailboxId))
    .delete(syncQueueKey(mailboxId));

  if (bundle.lease?.subscriptionId) {
    atomic.delete(subscriptionKey(bundle.lease.subscriptionId));
  }

  const result = await atomic.commit();
  if (!result.ok) throw new Error("Failed to delete mailbox bundle");

  await deleteByPrefix(kv, deliveredPrefix(mailboxId));
}
