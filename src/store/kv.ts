import { getConfigAsync } from "../config.ts";

let kvPromise: Promise<Deno.Kv> | null = null;

export function getKv(): Promise<Deno.Kv> {
  if (kvPromise) return kvPromise;
  kvPromise = (async () => {
    const config = await getConfigAsync();
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
