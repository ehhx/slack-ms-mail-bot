import type { PersistenceBackend } from "../config.ts";
import {
  createSupabaseKvStoreFromEnv,
  SupabaseKvStore,
} from "./supabase_kv.ts";

export type LegacyDenoKvMigrationKind =
  | "oauth_state"
  | "mailbox_connection"
  | "mailbox_email"
  | "team_mailbox"
  | "mailbox_route"
  | "mailbox_sync"
  | "mailbox_lease"
  | "subscription_mailbox"
  | "delivered_mail"
  | "sync_queue";

export interface LegacyDenoKvMigrationRequest {
  kind?: LegacyDenoKvMigrationKind;
  cursor?: string;
}

export interface LegacyDenoKvMigrationResult {
  kind: LegacyDenoKvMigrationKind;
  imported: number;
  skipped: number;
  nextCursor: string | null;
}

const DELIVERED_RECORD_TTL_MS = 90 * 24 * 60 * 60 * 1000;

export function isLegacyDenoKvMigrationKind(
  value: unknown,
): value is LegacyDenoKvMigrationKind {
  return value === "oauth_state" || value === "mailbox_connection" ||
    value === "mailbox_email" || value === "team_mailbox" ||
    value === "mailbox_route" || value === "mailbox_sync" ||
    value === "mailbox_lease" || value === "subscription_mailbox" ||
    value === "delivered_mail" || value === "sync_queue";
}

export async function importLegacyDenoKvPage(
  request: LegacyDenoKvMigrationRequest,
  batchSize: number,
): Promise<LegacyDenoKvMigrationResult> {
  if (!isLegacyDenoKvMigrationKind(request.kind)) {
    throw new Error("invalid migration kind");
  }

  const target = createSupabaseKvStoreFromEnv();
  if (!target) throw new Error("Supabase persistent store is unavailable");

  const legacyKv = await Deno.openKv();
  try {
    return await importLegacyDenoKvPageFromStore(
      legacyKv,
      target,
      request.kind,
      request.cursor,
      batchSize,
    );
  } finally {
    legacyKv.close();
  }
}

export function persistenceBackendSupportsLegacyMigration(
  backend: PersistenceBackend,
): boolean {
  return backend === "dual_write" || backend === "supabase";
}

export async function importLegacyDenoKvPageFromStore(
  legacyKv: Deno.Kv,
  target: SupabaseKvStore,
  kind: LegacyDenoKvMigrationKind,
  cursor: string | undefined,
  batchSize: number,
): Promise<LegacyDenoKvMigrationResult> {
  const entries = legacyKv.list<unknown>(
    { prefix: legacyDenoKvPrefix(kind) },
    { limit: Math.max(1, Math.min(batchSize, 500)), cursor },
  );
  let imported = 0;
  let skipped = 0;

  for await (const entry of entries) {
    const options = legacyDenoKvImportOptions(kind, entry.value);
    if (!options) {
      skipped++;
      continue;
    }
    if (await target.setIfAbsent(entry.key, entry.value, options)) {
      imported++;
    } else {
      skipped++;
    }
  }

  return {
    kind,
    imported,
    skipped,
    nextCursor: entries.cursor || null,
  };
}

function legacyDenoKvPrefix(kind: LegacyDenoKvMigrationKind): Deno.KvKey {
  return [kind];
}

function legacyDenoKvImportOptions(
  kind: LegacyDenoKvMigrationKind,
  value: unknown,
): { expireIn?: number } | null {
  if (kind === "oauth_state") {
    if (!isRecord(value) || typeof value.expiresAt !== "string") return null;
    const expireIn = Date.parse(value.expiresAt) - Date.now();
    return expireIn > 0 ? { expireIn } : null;
  }

  if (kind === "delivered_mail") {
    if (!isRecord(value) || typeof value.deliveredAt !== "string") return null;
    const expireIn = Date.parse(value.deliveredAt) + DELIVERED_RECORD_TTL_MS -
      Date.now();
    return expireIn > 0 ? { expireIn } : null;
  }

  return {};
}

function isRecord(value: unknown): value is Record<string, unknown> {
  return Boolean(value) && typeof value === "object" && !Array.isArray(value);
}
