/**
 * Supabase-backed compatibility layer for the subset of Deno KV used by the
 * mail service. Keeping the existing key encoding makes the migration cursor
 * based and lets the business layer switch backends without changing its
 * mailbox semantics.
 */

export type MailKvKey = Deno.KvKey;

interface StoredState {
  value: unknown;
  version: number | string | null;
}

interface StoredListEntry {
  state_key: string;
  state_value: unknown;
  state_version: number | string;
}

interface Mutation {
  kind: "set" | "delete";
  state_key: string;
  state_value?: unknown;
  expires_at?: string | null;
}

interface ExpiringOptions {
  expireIn?: number;
}

const DEFAULT_REQUEST_TIMEOUT_MS = 10_000;
const MAX_LIST_PAGE_SIZE = 1_000;

export class SupabaseKvStore {
  constructor(
    private readonly apiUrl: string,
    private readonly serviceRoleKey: string,
    private readonly requestTimeoutMs = DEFAULT_REQUEST_TIMEOUT_MS,
    private readonly fetchImpl: typeof fetch = fetch,
  ) {}

  async get<T>(key: MailKvKey): Promise<{
    key: MailKvKey;
    value: T | null;
    versionstamp: string | null;
  }> {
    const state = await this.rpc<StoredState>("mail_state_get", {
      p_state_key: encodeKey(key),
    });
    return {
      key,
      value: (state?.value ?? null) as T | null,
      versionstamp: normalizeVersionstamp(state?.version),
    };
  }

  async getMany(keys: readonly MailKvKey[]): Promise<
    Array<{
      key: MailKvKey;
      value: unknown;
      versionstamp: string | null;
    }>
  > {
    const encodedKeys = keys.map((key) => encodeKey(key));
    const rows = await this.rpc<
      Array<{
        state_key: string;
        state_value: unknown;
        state_version: number | string | null;
      }>
    >("mail_state_get_many", { p_state_keys: encodedKeys });
    return (rows ?? []).map((row, index) => ({
      key: keys[index],
      value: row.state_value ?? null,
      versionstamp: normalizeVersionstamp(row.state_version),
    }));
  }

  async set<T>(
    key: MailKvKey,
    value: T,
    options: ExpiringOptions = {},
  ): Promise<void> {
    await this.rpc("mail_state_set", {
      p_state_key: encodeKey(key),
      p_state_value: value,
      p_expires_at: resolveExpiresAt(options.expireIn),
    });
  }

  async setIfAbsent<T>(
    key: MailKvKey,
    value: T,
    options: ExpiringOptions = {},
  ): Promise<boolean> {
    const result = await this.rpc<boolean>("mail_state_set_if_absent", {
      p_state_key: encodeKey(key),
      p_state_value: value,
      p_expires_at: resolveExpiresAt(options.expireIn),
    });
    return result === true;
  }

  async delete(key: MailKvKey): Promise<void> {
    await this.rpc("mail_state_delete", {
      p_state_key: encodeKey(key),
    });
  }

  async *list<T>(
    selector: { prefix: MailKvKey },
    options: { limit?: number; cursor?: string } = {},
  ): AsyncGenerator<{
    key: MailKvKey;
    value: T;
    versionstamp: string;
  }> {
    const prefix = encodeKey(selector.prefix);
    let after = options.cursor ?? null;
    let remaining = Math.max(
      0,
      options.limit ?? Number.MAX_SAFE_INTEGER,
    );

    while (remaining > 0) {
      const pageSize = Math.max(
        1,
        Math.min(remaining, MAX_LIST_PAGE_SIZE),
      );
      const rows = await this.rpc<StoredListEntry[]>("mail_state_list", {
        p_prefix: prefix,
        p_limit: pageSize,
        p_after: after,
      });
      if (!rows?.length) return;

      for (const row of rows) {
        yield {
          key: decodeKey(row.state_key),
          value: row.state_value as T,
          versionstamp: String(row.state_version),
        };
        remaining--;
        after = row.state_key;
        if (remaining === 0) return;
      }

      if (rows.length < pageSize) return;
    }
  }

  atomic(): SupabaseAtomicOperation {
    return new SupabaseAtomicOperation(this);
  }

  async atomicBatch(mutations: Mutation[]): Promise<void> {
    for (const mutation of mutations) {
      if (mutation.kind === "set") {
        await this.rpc("mail_state_set", {
          p_state_key: mutation.state_key,
          p_state_value: mutation.state_value,
          p_expires_at: mutation.expires_at ?? null,
        });
        continue;
      }
      await this.rpc("mail_state_delete", {
        p_state_key: mutation.state_key,
      });
    }
  }

  async pruneExpired(limit = 100): Promise<number> {
    const result = await this.rpc<number>("mail_state_prune_expired", {
      p_limit: Math.max(1, Math.min(limit, MAX_LIST_PAGE_SIZE)),
    });
    return Number.isFinite(result) ? Number(result) : 0;
  }

  private async rpc<T>(
    name: string,
    body: Record<string, unknown>,
  ): Promise<T> {
    const controller = new AbortController();
    const timeout = setTimeout(() => controller.abort(), this.requestTimeoutMs);
    try {
      const response = await this.fetchImpl(
        `${this.apiUrl}/rest/v1/rpc/${name}`,
        {
          method: "POST",
          headers: {
            "Content-Type": "application/json",
            apikey: this.serviceRoleKey,
            Authorization: `Bearer ${this.serviceRoleKey}`,
          },
          body: JSON.stringify(body),
          signal: controller.signal,
        },
      );
      const responseText = await response.text();
      if (!response.ok) {
        throw new Error(
          `Supabase state RPC ${name} failed: HTTP ${response.status} ${
            responseText.slice(0, 300)
          }`,
        );
      }
      return responseText ? JSON.parse(responseText) as T : undefined as T;
    } finally {
      clearTimeout(timeout);
    }
  }
}

export class SupabaseAtomicOperation {
  private readonly mutations: Mutation[] = [];

  constructor(private readonly store: SupabaseKvStore) {}

  set<T>(key: MailKvKey, value: T, options: ExpiringOptions = {}): this {
    this.mutations.push({
      kind: "set",
      state_key: encodeKey(key),
      state_value: value,
      expires_at: resolveExpiresAt(options.expireIn),
    });
    return this;
  }

  delete(key: MailKvKey): this {
    this.mutations.push({ kind: "delete", state_key: encodeKey(key) });
    return this;
  }

  check(): this {
    throw new Error("Supabase atomic checks are not used by the mail store");
  }

  async commit(): Promise<{ ok: boolean; versionstamp: null }> {
    if (this.mutations.length === 0) {
      return { ok: true, versionstamp: null };
    }
    await this.store.atomicBatch(this.mutations);
    return { ok: true, versionstamp: null };
  }
}

export function createSupabaseKvStoreFromEnv(): SupabaseKvStore | null {
  const rawUrl = Deno.env.get("SUPABASE_URL")?.trim();
  const serviceRoleKey = Deno.env.get("SUPABASE_SERVICE_ROLE_KEY")?.trim();
  if (!rawUrl || !serviceRoleKey) return null;

  const parsedTimeout = Number.parseInt(
    Deno.env.get("SUPABASE_REQUEST_TIMEOUT_MS") ?? "10000",
    10,
  );
  const requestTimeoutMs = Number.isFinite(parsedTimeout)
    ? Math.max(1_000, parsedTimeout)
    : DEFAULT_REQUEST_TIMEOUT_MS;

  return new SupabaseKvStore(
    rawUrl.replace(/\/+$/, ""),
    serviceRoleKey,
    requestTimeoutMs,
  );
}

function resolveExpiresAt(expireIn: number | undefined): string | null {
  if (expireIn == null) return null;
  return new Date(Date.now() + Math.max(0, expireIn)).toISOString();
}

function normalizeVersionstamp(
  value: number | string | null | undefined,
): string | null {
  if (value == null) return null;
  return String(value);
}

function encodeKey(key: MailKvKey): string {
  return key.map(encodeKeyPart).join("/");
}

function encodeKeyPart(part: Deno.KvKeyPart): string {
  if (part instanceof Uint8Array) {
    return `u:${
      Array.from(part).map((value) => value.toString(16).padStart(2, "0")).join(
        "",
      )
    }`;
  }
  switch (typeof part) {
    case "string":
      return `s:${encodeURIComponent(part)}`;
    case "number":
      return `n:${part}`;
    case "boolean":
      return `b:${part ? "1" : "0"}`;
    case "bigint":
      return `i:${part}`;
    case "symbol":
      throw new TypeError("Symbol key parts are not supported by Supabase");
  }
}

function decodeKey(encoded: string): MailKvKey {
  if (!encoded) return [];
  return encoded.split("/").map((segment) => {
    const kind = segment.slice(0, 2);
    const raw = segment.slice(2);
    switch (kind) {
      case "s:":
        return decodeURIComponent(raw);
      case "n:":
        return Number(raw);
      case "b:":
        return raw === "1";
      case "i:":
        return BigInt(raw);
      case "u:": {
        if (raw.length % 2 !== 0) throw new Error("invalid persisted key");
        const bytes = new Uint8Array(raw.length / 2);
        for (let index = 0; index < bytes.length; index++) {
          bytes[index] = Number.parseInt(
            raw.slice(index * 2, index * 2 + 2),
            16,
          );
        }
        return bytes;
      }
      default:
        throw new Error("invalid persisted state key");
    }
  });
}
