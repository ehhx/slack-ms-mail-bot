export type WaitUntilLike = { waitUntil(promise: Promise<unknown>): void };

export function runBackground(
  ctx: WaitUntilLike | undefined,
  promise: Promise<unknown>,
): void {
  if (ctx?.waitUntil) {
    ctx.waitUntil(promise);
    return;
  }
  promise.catch((error) => console.error("background task failed", error));
}
