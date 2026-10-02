import type { WorkerLike } from "../../src/worker-client.js";
import {
  attachWorkerEndpoint,
  type WorkerOperationHandler,
  type WorkerScopeLike,
} from "../../src/worker-endpoint.js";
import type {
  WorkerInboundMessage,
  WorkerOutboundMessage,
} from "../../src/worker-protocol.js";

/**
 * An in-process stand-in for a Web Worker: the endpoint runs `handler` on the
 * same thread, messages cross asynchronously like a real port, and `crash()`
 * simulates the worker dying mid-request.
 */
export function loopbackWorker(handler: WorkerOperationHandler): {
  readonly worker: WorkerLike;
  readonly terminated: () => boolean;
  crash(): void;
} {
  const toWorker = new Set<
    (event: MessageEvent<WorkerInboundMessage>) => void
  >();
  const toMain = new Set<
    (event: MessageEvent<WorkerOutboundMessage>) => void
  >();
  const errors = new Set<(event: Event) => void>();
  let terminated = false;
  const scope: WorkerScopeLike = {
    postMessage(message) {
      if (terminated) return;
      queueMicrotask(() => {
        for (const listener of toMain)
          listener({ data: message } as MessageEvent<WorkerOutboundMessage>);
      });
    },
    addEventListener(_type, listener) {
      toWorker.add(listener);
    },
    removeEventListener(_type, listener) {
      toWorker.delete(listener);
    },
  };
  const detach = attachWorkerEndpoint(scope, handler);
  const worker = {
    postMessage(message: WorkerInboundMessage) {
      if (terminated) return;
      queueMicrotask(() => {
        for (const listener of toWorker)
          listener({ data: message } as MessageEvent<WorkerInboundMessage>);
      });
    },
    addEventListener(type: string, listener: never) {
      if (type === "message") toMain.add(listener);
      else errors.add(listener);
    },
    removeEventListener(type: string, listener: never) {
      if (type === "message") toMain.delete(listener);
      else errors.delete(listener);
    },
    terminate() {
      terminated = true;
      detach();
    },
  } as unknown as WorkerLike;
  return {
    worker,
    terminated: () => terminated,
    crash() {
      terminated = true;
      detach();
      for (const listener of errors) listener(new Event("error"));
    },
  };
}
