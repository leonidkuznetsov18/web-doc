/** Disposable historical-source diagnostics; no document values are accepted. */
const clock = `realm-${crypto.randomUUID()}`;
const fields = new Set([
  "event",
  "clock",
  "navigation",
  "format",
  "worker",
  "id",
  "operation",
  "kind",
  "timeoutMs",
  "now",
]);

export function safeMetadata(value: unknown): Record<string, string | number> {
  const out: Record<string, string | number> = {};
  if (value === null || typeof value !== "object") return out;
  for (const key of fields) {
    const field: unknown = Reflect.get(value, key);
    if (typeof field === "number" && Number.isFinite(field)) out[key] = field;
    if (
      typeof field === "string" &&
      field.length <= 128 &&
      /^[a-zA-Z0-9_.:/-]+$/.test(field)
    )
      out[key] = field;
  }
  return out;
}

export function recordPhase(
  event: string,
  details: unknown = {},
  worker?: object,
): void {
  const rawContext: unknown = Reflect.get(globalThis, "__action915Context");
  const context = safeMetadata(rawContext);
  const clean = safeMetadata(details);
  if (typeof clean.format === "string") {
    context.format = clean.format;
    if (rawContext !== null && typeof rawContext === "object")
      Reflect.set(rawContext, "format", clean.format);
  }
  const ids: unknown = Reflect.get(globalThis, "__action915WorkerIds");
  if (worker && ids instanceof WeakMap) {
    const id: unknown = ids.get(worker);
    if (typeof id === "number") clean.worker = id;
  }
  const metadata = {
    ...context,
    ...clean,
    event,
    clock: typeof context.clock === "string" ? context.clock : clock,
    now: performance.now(),
  };
  if (typeof document === "undefined") {
    // A separate diagnostic envelope is the only added worker traffic. The
    // native request/reply messages and worker URL are never intercepted.
    const send: unknown = Reflect.get(globalThis, "postMessage");
    if (typeof send === "function")
      Reflect.apply(send, globalThis, [{ action915: true, metadata }]);
  } else console.debug("ACTION915:" + JSON.stringify(metadata));
}
