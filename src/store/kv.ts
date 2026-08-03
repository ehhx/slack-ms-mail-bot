import { getConfigAsync } from "../config.ts";
import {
  createSupabaseKvStoreFromEnv,
  SupabaseKvStore,
} from "./supabase_kv.ts";

let kvPromise: Promise<Deno.Kv> | null = null;

export function getKv(): Promise<Deno.Kv> {
  if (kvPromise) return kvPromise;
  kvPromise = (async () => {
    const config = await getConfigAsync();
    if (config.persistenceBackend !== "deno_kv") {
      const supabase = createSupabaseKvStoreFromEnv();
      if (!supabase) {
        throw new Error(
          "SUPABASE_URL and SUPABASE_SERVICE_ROLE_KEY are required for the configured persistence backend",
        );
      }
      if (config.persistenceBackend === "supabase") {
        return supabase as unknown as Deno.Kv;
      }

      const legacyKv = config.kvPath
        ? await Deno.openKv(config.kvPath)
        : await Deno.openKv();
      return new DualWriteKvStore(legacyKv, supabase) as unknown as Deno.Kv;
    }
    if (config.kvPath) {
      return await Deno.openKv(config.kvPath);
    }
    return await Deno.openKv();
  })();
  return kvPromise;
}

export function setKvForTesting(kv: Deno.Kv | null): void {
  kvPromise = kv ? Promise.resolve(kv) : null;
}

export async function getPersistenceStatus(): Promise<{
  backend: "deno_kv" | "dual_write" | "supabase";
  available: boolean;
  error?: string;
}> {
  const config = await getConfigAsync();
  try {
    const kv = await getKv();
    if (kv instanceof DualWriteKvStore) {
      await kv.healthCheck();
    } else {
      await kv.get(["__mail_healthcheck"]);
    }
    return { backend: config.persistenceBackend, available: true };
  } catch (error) {
    return {
      backend: config.persistenceBackend,
      available: false,
      error: error instanceof Error ? error.message : String(error),
    };
  }
}

export async function pruneExpiredState(limit = 100): Promise<number> {
  if (!kvPromise) return 0;
  const kv = await kvPromise;
  if (kv instanceof SupabaseKvStore) return await kv.pruneExpired(limit);
  if (kv instanceof DualWriteKvStore) return await kv.pruneExpired(limit);
  return 0;
}

export async function deleteByPrefix(
  kv: Deno.Kv,
  prefix: Deno.KvKey,
): Promise<number> {
  let count = 0;
  for await (const entry of kv.list({ prefix })) {
    await kv.delete(entry.key);
    count++;
  }
  return count;
}

interface DualWriteMutation {
  kind: "set" | "delete";
  key: Deno.KvKey;
  value?: unknown;
  options?: { expireIn?: number };
}

class DualWriteKvStore {
  constructor(
    private readonly primary: Deno.Kv,
    private readonly secondary: SupabaseKvStore,
  ) {}

  get<T>(key: Deno.KvKey): Promise<{
    key: Deno.KvKey;
    value: T | null;
    versionstamp: string | null;
  }> {
    return this.primary.get<T>(key);
  }

  getMany(keys: readonly Deno.KvKey[]): Promise<
    Array<{
      key: Deno.KvKey;
      value: unknown;
      versionstamp: string | null;
    }>
  > {
    return this.primary.getMany(keys as Deno.KvKey[]) as Promise<
      Array<{
        key: Deno.KvKey;
        value: unknown;
        versionstamp: string | null;
      }>
    >;
  }

  list<T>(
    selector: { prefix: Deno.KvKey },
    options?: { limit?: number; cursor?: string },
  ): AsyncIterable<{
    key: Deno.KvKey;
    value: T;
    versionstamp: string;
  }> {
    return this.primary.list<T>(selector, options);
  }

  async set<T>(
    key: Deno.KvKey,
    value: T,
    options?: { expireIn?: number },
  ): Promise<void> {
    await this.secondary.set(key, value, options);
    await this.primary.set(key, value, options);
  }

  async delete(key: Deno.KvKey): Promise<void> {
    await this.secondary.delete(key);
    await this.primary.delete(key);
  }

  atomic(): DualWriteAtomicOperation {
    return new DualWriteAtomicOperation(this.primary, this.secondary);
  }

  async pruneExpired(limit: number): Promise<number> {
    return await this.secondary.pruneExpired(limit);
  }

  async healthCheck(): Promise<void> {
    await Promise.all([
      this.primary.get(["__mail_healthcheck"]),
      this.secondary.get(["__mail_healthcheck"]),
    ]);
  }

  close(): void {
    this.primary.close();
  }
}

class DualWriteAtomicOperation {
  private readonly mutations: DualWriteMutation[] = [];

  constructor(
    private readonly primary: Deno.Kv,
    private readonly secondary: SupabaseKvStore,
  ) {}

  set<T>(
    key: Deno.KvKey,
    value: T,
    options?: { expireIn?: number },
  ): this {
    this.mutations.push({ kind: "set", key, value, options });
    return this;
  }

  delete(key: Deno.KvKey): this {
    this.mutations.push({ kind: "delete", key });
    return this;
  }

  check(): this {
    throw new Error("Dual-write atomic checks are not used by the mail store");
  }

  async commit(): Promise<{ ok: boolean; versionstamp: string | null }> {
    const secondaryAtomic = this.secondary.atomic();
    const primaryAtomic = this.primary.atomic();
    for (const mutation of this.mutations) {
      if (mutation.kind === "set") {
        secondaryAtomic.set(
          mutation.key,
          mutation.value,
          mutation.options,
        );
        primaryAtomic.set(
          mutation.key,
          mutation.value,
          mutation.options,
        );
      } else {
        secondaryAtomic.delete(mutation.key);
        primaryAtomic.delete(mutation.key);
      }
    }

    await secondaryAtomic.commit();
    const result = await primaryAtomic.commit();
    return result.ok
      ? { ok: true, versionstamp: result.versionstamp }
      : { ok: false, versionstamp: null };
  }
}
