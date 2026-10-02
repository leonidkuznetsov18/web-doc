import type { ViewerWarning } from "../contracts.js";
import type { WorkerRpcClient } from "../worker-client.js";
import type { EditWorkerOperation } from "../worker-protocol.js";
import type {
  EditEngine,
  EditEngineContext,
  BatchMode,
  EngineBatch,
  EngineChange,
  MaterializedDocument,
  MaterializeOptions,
  RestoreTarget,
} from "./engine.js";
import type {
  EditElement,
  EditFindOptions,
  EditOperation,
  ElementQuery,
  OperationIssue,
  OperationSchemaSet,
  PagePoint,
  TextTarget,
} from "./types.js";

/*
 * The main-thread side of an edit engine that lives in a worker: every core
 * engine call is one request over the worker RPC. Format clients extend it
 * with how their worker starts and with their own reads.
 */

export abstract class WorkerEngineClient implements EditEngine {
  abstract readonly schemas: OperationSchemaSet;
  protected readonly rpc: WorkerRpcClient;
  protected readonly context: EditEngineContext;
  /** Highest state id seen, for batches passed as plain arrays. */
  #stateId = 0;

  constructor(rpc: WorkerRpcClient, context: EditEngineContext) {
    this.rpc = rpc;
    this.context = context;
  }

  validate(
    operations: readonly EditOperation[],
    signal: AbortSignal,
    mode: BatchMode = {},
  ): Promise<readonly OperationIssue[]> {
    return this.request("edit-validate", { operations, mode }, signal);
  }

  /** Plain operation arrays, as the unit tests pass them, become the next batch. */
  apply(
    input: EngineBatch | readonly EditOperation[],
    signal: AbortSignal,
  ): Promise<EngineChange> {
    const batch: EngineBatch = Array.isArray(input)
      ? { stateId: ++this.#stateId, operations: input }
      : (input as EngineBatch);
    this.#stateId = Math.max(this.#stateId, batch.stateId);
    return this.request("edit-apply", { batch }, signal);
  }

  /** A bare signal, as the unit tests pass it, asks for the shown form. */
  async materialize(
    purposeOrSignal: "show" | "save" | AbortSignal = "show",
    options: MaterializeOptions = {},
    signal: AbortSignal = new AbortController().signal,
  ): Promise<Uint8Array> {
    return (await this.materializeDocument(purposeOrSignal, options, signal))
      .bytes;
  }

  async materializeDocument(
    purposeOrSignal: "show" | "save" | AbortSignal = "show",
    options: MaterializeOptions = {},
    signal: AbortSignal = new AbortController().signal,
  ): Promise<MaterializedDocument> {
    const purpose =
      purposeOrSignal instanceof AbortSignal ? "show" : purposeOrSignal;
    const own =
      purposeOrSignal instanceof AbortSignal ? purposeOrSignal : signal;
    const result = await this.request<{
      readonly data: ArrayBuffer;
      readonly warnings: readonly ViewerWarning[];
    }>("edit-materialize", { purpose, options }, own);
    return { bytes: new Uint8Array(result.data), warnings: result.warnings };
  }

  restore(
    input: RestoreTarget | readonly (readonly EditOperation[])[],
    signal: AbortSignal,
  ): Promise<void> {
    const target: RestoreTarget = Array.isArray(input)
      ? {
          batches: (input as readonly (readonly EditOperation[])[]).map(
            (operations, index) => ({ stateId: index + 1, operations }),
          ),
        }
      : (input as RestoreTarget);
    const base = target.base?.slice().buffer;
    return this.request(
      "edit-restore",
      { batches: target.batches, ...(base ? { base } : {}) },
      signal,
      base ? [base] : undefined,
    );
  }

  putAsset(id: string, data: Uint8Array, signal: AbortSignal): Promise<void> {
    const buffer = data.slice().buffer;
    return this.request("edit-put-asset", { id, data: buffer }, signal, [
      buffer,
    ]);
  }

  getElements(
    query: ElementQuery,
    signal: AbortSignal,
  ): Promise<readonly EditElement[]> {
    return this.request("edit-elements", { query }, signal);
  }

  getElement(
    id: string,
    signal: AbortSignal,
  ): Promise<EditElement | undefined> {
    return this.request("edit-element", { id }, signal);
  }

  elementsAt(
    pageIndex: number,
    point: PagePoint,
    signal: AbortSignal,
  ): Promise<readonly EditElement[]> {
    return this.request("edit-elements-at", { pageIndex, point }, signal);
  }

  findText(
    query: string,
    options: EditFindOptions,
    signal: AbortSignal,
  ): Promise<readonly TextTarget[]> {
    return this.request("edit-find-text", { query, options }, signal);
  }

  async dispose(): Promise<void> {
    // A healthy worker answers at once; a dead one never would, and nothing
    // is lost by terminating it without an answer.
    const request = this.request("edit-dispose", undefined, undefined).catch(
      () => undefined,
    );
    await Promise.race([
      request,
      new Promise((resolve) => setTimeout(resolve, DISPOSE_GRACE_MS)),
    ]);
    this.rpc.destroy();
  }

  protected request<T>(
    operation: EditWorkerOperation,
    payload: unknown,
    signal: AbortSignal | undefined,
    transfer?: Transferable[],
  ): Promise<T> {
    return this.rpc.request<T>(operation, payload, {
      ...(signal ? { signal } : {}),
      ...(transfer ? { transfer } : {}),
      timeoutMs: this.context.limits.maxOperationMs,
    });
  }
}

/** How long disposal waits for the worker's answer before terminating it. */
export const DISPOSE_GRACE_MS = 1000;

/** `value` resolved against `base`, else the packaged fallback. */
export function resolveAssetUrl(
  value: string | URL | undefined,
  base: URL | undefined,
  fallback: URL,
): URL {
  return value === undefined ? fallback : new URL(value, base);
}

/** A URL next to this module; data-driven so bundlers leave the package-owned directory alone. */
export function packageRelativeUrl(path: string, from: string): URL {
  return new URL(path, from);
}
