import type {
  DocumentAdapter,
  ResourceLimits,
  TextRun,
  ViewerEventMap,
} from "../../src/contracts.js";
import type {
  RestoreTarget,
  EngineBatch,
  EditEngine,
  EditEngineContext,
  EditEngineProvider,
  EditSessionCore,
  EngineChange,
} from "../../src/edit/engine.js";
import type { EditSession } from "../../src/edit/sessions.js";
import type {
  EditSessionHost,
  PreparedDocument,
} from "../../src/edit/session.js";
import type {
  EditableFormat,
  EditElement,
  EditFindOptions,
  EditOperation,
  ElementQuery,
  OperationIssue,
  OperationSchemaSet,
  PagePoint,
  TextTarget,
} from "../../src/edit/types.js";
import { abortError } from "../../src/errors.js";
import { defaultResourceLimits, resolveLimits } from "../../src/limits.js";

/*
 * A text document engine for exercising the editing core without a real
 * format: pages are strings, the file is their JSON encoding. Every failure
 * mode the core must survive can be scripted.
 */

export type FakeOperation =
  | {
      readonly op: "setText";
      readonly pageIndex: number;
      readonly text: string;
    }
  | { readonly op: "insertPage"; readonly index: number; readonly text: string }
  | { readonly op: "deletePage"; readonly pageIndex: number }
  /** Appends the byte length of a binary payload to a page's text. */
  | {
      readonly op: "stamp";
      readonly pageIndex: number;
      readonly data: Uint8Array | string;
    }
  /** Passes validation, fails in `apply`. */
  | { readonly op: "fail" }
  /** Passes validation, waits until the signal aborts. */
  | { readonly op: "hang" };

export const fakeSchemas: OperationSchemaSet = {
  format: "pdf",
  version: 1,
  operations: {
    setText: {
      type: "object",
      required: ["op", "pageIndex", "text"],
      additionalProperties: false,
      properties: {
        op: { const: "setText" },
        pageIndex: { type: "integer", minimum: 0 },
        text: { type: "string", maxLength: 100 },
      },
    },
    insertPage: {
      type: "object",
      required: ["op", "index", "text"],
      additionalProperties: false,
      properties: {
        op: { const: "insertPage" },
        index: { type: "integer", minimum: 0 },
        text: { type: "string" },
      },
    },
    deletePage: {
      type: "object",
      required: ["op", "pageIndex"],
      additionalProperties: false,
      properties: {
        op: { const: "deletePage" },
        pageIndex: { type: "integer", minimum: 0 },
      },
    },
    stamp: {
      type: "object",
      required: ["op", "pageIndex", "data"],
      additionalProperties: false,
      properties: {
        op: { const: "stamp" },
        pageIndex: { type: "integer", minimum: 0 },
        data: { "x-binary": true, type: "string", contentEncoding: "base64" },
      },
    },
    fail: { type: "object", properties: { op: { const: "fail" } } },
    hang: { type: "object", properties: { op: { const: "hang" } } },
  },
};

/** Files start with a PDF signature so format detection routes them to the adapter. */
const SIGNATURE = "%PDF-1.7\n";

export function encodePages(pages: readonly string[]): Uint8Array {
  return new TextEncoder().encode(SIGNATURE + JSON.stringify(pages));
}

export function decodePages(bytes: Uint8Array): string[] {
  const text = new TextDecoder().decode(bytes);
  if (!text.startsWith(SIGNATURE)) throw new Error("not a fake document");
  return JSON.parse(text.slice(SIGNATURE.length)) as string[];
}

export interface FakeEngineOptions {
  /** Throw from `materialize` while true. */
  failMaterialize?: boolean;
  /** Throw from `restore` while true. */
  failRestore?: boolean;
}

export class FakeEditEngine implements EditEngine {
  readonly schemas = fakeSchemas;
  readonly original: readonly string[];
  readonly options: FakeEngineOptions;
  readonly calls: string[] = [];
  pages: string[];
  disposed = false;

  constructor(original: Uint8Array, options: FakeEngineOptions = {}) {
    this.original = decodePages(original);
    this.pages = [...this.original];
    this.options = options;
  }

  async validate(
    operations: readonly EditOperation[],
  ): Promise<readonly OperationIssue[]> {
    this.calls.push("validate");
    const issues: OperationIssue[] = [];
    let pageCount = this.pages.length;
    operations.forEach((raw, operationIndex) => {
      const operation = raw as FakeOperation;
      if (operation.op === "setText" || operation.op === "deletePage") {
        if (operation.pageIndex >= pageCount)
          issues.push({
            operationIndex,
            path: "/pageIndex",
            code: "unknown-target",
            message: `No page ${operation.pageIndex}`,
          });
        else if (operation.op === "deletePage") pageCount -= 1;
      } else if (operation.op === "insertPage") pageCount += 1;
    });
    return issues;
  }

  async apply(batch: EngineBatch, signal: AbortSignal): Promise<EngineChange> {
    this.calls.push("apply");
    const changed = new Set<number>();
    const createdIds: string[] = [];
    for (const raw of batch.operations) {
      const operation = raw as FakeOperation;
      switch (operation.op) {
        case "setText":
          this.pages[operation.pageIndex] = operation.text;
          changed.add(operation.pageIndex);
          break;
        case "insertPage":
          this.pages.splice(operation.index, 0, operation.text);
          createdIds.push(`page:${operation.text}`);
          for (
            let index = operation.index;
            index < this.pages.length;
            index += 1
          )
            changed.add(index);
          break;
        case "deletePage":
          this.pages.splice(operation.pageIndex, 1);
          for (
            let index = operation.pageIndex;
            index < this.pages.length;
            index += 1
          )
            changed.add(index);
          break;
        case "stamp": {
          // Payloads arrive as asset references once the core interned them.
          const bytes =
            typeof operation.data === "string"
              ? this.assets.get(operation.data)
              : operation.data;
          if (!bytes)
            throw new Error(`unknown asset ${String(operation.data)}`);
          this.pages[operation.pageIndex] += `+${bytes.byteLength}`;
          changed.add(operation.pageIndex);
          break;
        }
        case "fail":
          // Half-applied on purpose: the core must roll this back.
          this.pages[0] = "CORRUPT";
          throw new Error("engine exploded");
        case "hang":
          await new Promise<never>((_, reject) =>
            signal.addEventListener("abort", () => reject(abortError()), {
              once: true,
            }),
          );
      }
    }
    return {
      createdIds,
      removedIds: [],
      changedPages: [...changed].sort((a, b) => a - b),
      pageCount: this.pages.length,
      warnings:
        batch.operations.length > 1
          ? [{ code: "unsupported-feature", message: "batch" }]
          : [],
    };
  }

  async materialize(
    purpose: "show" | "save" = "show",
    options: Readonly<Record<string, unknown>> = {},
  ): Promise<Uint8Array> {
    this.calls.push(
      purpose === "save" && options.mode !== undefined
        ? `materialize:${String(options.mode)}`
        : "materialize",
    );
    if (this.options.failMaterialize) throw new Error("disk full");
    return encodePages(this.pages);
  }

  /** Bases the last restore started from, decoded; undefined means the original. */
  readonly restoreBases: (string[] | undefined)[] = [];
  readonly assets = new Map<string, Uint8Array>();

  async restore(target: RestoreTarget): Promise<void> {
    this.calls.push(`restore:${target.batches.length}`);
    if (this.options.failRestore) throw new Error("restore failed");
    this.restoreBases.push(target.base ? decodePages(target.base) : undefined);
    this.pages = target.base ? decodePages(target.base) : [...this.original];
    for (const batch of target.batches)
      await this.apply(batch, new AbortController().signal);
  }

  async putAsset(id: string, data: Uint8Array): Promise<void> {
    this.calls.push(`putAsset:${id.slice(0, 12)}`);
    this.assets.set(id, data);
  }

  async getElements(query: ElementQuery): Promise<readonly EditElement[]> {
    this.calls.push("getElements");
    return this.pages.flatMap((text, pageIndex) =>
      query.pageIndex !== undefined && query.pageIndex !== pageIndex
        ? []
        : text.split(" ").map((word, index) => ({
            id: `p${pageIndex}w${index}`,
            kind: "word",
            pageIndex,
            bounds: { x: index * 10, y: 0, width: 10, height: 10 },
            text: word,
            operations: ["setText"],
          })),
    );
  }

  async elementsAt(
    pageIndex: number,
    point: PagePoint,
  ): Promise<readonly EditElement[]> {
    const elements = await this.getElements({ pageIndex });
    return elements.filter(
      (element) =>
        point.x >= element.bounds.x &&
        point.x < element.bounds.x + element.bounds.width,
    );
  }

  async findText(
    query: string,
    options: EditFindOptions,
  ): Promise<readonly TextTarget[]> {
    const targets: TextTarget[] = [];
    this.pages.forEach((text, pageIndex) => {
      const haystack = options.caseSensitive ? text : text.toLowerCase();
      const needle = options.caseSensitive ? query : query.toLowerCase();
      if (haystack.includes(needle))
        targets.push({
          pageIndex,
          text: query,
          rects: [{ x: 0, y: 0, width: 10, height: 10 }],
          elementIds: [`p${pageIndex}w0`],
        });
    });
    return targets;
  }

  async dispose(): Promise<void> {
    this.disposed = true;
  }
}

export function fakeProvider(
  options: FakeEngineOptions = {},
  onLoad?: (engine: FakeEditEngine) => void,
): EditEngineProvider & { readonly loads: number } {
  const provider = {
    loads: 0,
    formats: ["pdf"] as const satisfies readonly EditableFormat[],
    async load(original: Uint8Array, _context: EditEngineContext) {
      provider.loads += 1;
      const engine = new FakeEditEngine(original, options);
      onLoad?.(engine);
      return engine;
    },
    // The fake format borrows the PDF session type; tests cast it back.
    createSession: (core: EditSessionCore) => core as unknown as EditSession,
  };
  return provider;
}

export interface FakeHostOptions {
  readonly limits?: Partial<ResourceLimits>;
  /** Throw from `prepareDocument` while true. */
  failReplace?: boolean;
  /** Throw from `emit` while true, like a host listener with a bug. */
  failEmit?: boolean;
  /** Runs inside `prepareDocument`, before it resolves. */
  onPrepare?: () => void;
  /** Runs inside `commitDocument`, before it returns. */
  onCommit?: () => void;
  /** Page count the host reports instead of the real one. */
  reportPageCount?: number;
}

export class FakeHost implements EditSessionHost {
  readonly format: EditableFormat = "pdf";
  readonly limits: ResourceLimits;
  readonly options: FakeHostOptions;
  readonly events: { readonly type: string; readonly event: unknown }[] = [];
  /** Every document the viewer was asked to show, decoded. */
  readonly shown: string[][] = [];
  /** Preparations that were opened but never shown. */
  readonly discarded: string[][] = [];

  constructor(options: FakeHostOptions = {}) {
    this.options = options;
    this.limits = resolveLimits(defaultResourceLimits, options.limits);
  }

  async prepareDocument(
    bytes: Uint8Array,
  ): Promise<PreparedDocument & { readonly pages: string[] }> {
    if (this.options.failReplace) throw new Error("renderer rejected the file");
    const pages = decodePages(bytes);
    this.options.onPrepare?.();
    return { pageCount: this.options.reportPageCount ?? pages.length, pages };
  }

  commitDocument(prepared: PreparedDocument): number {
    const { pages } = prepared as PreparedDocument & { pages: string[] };
    this.shown.push(pages);
    this.options.onCommit?.();
    return prepared.pageCount;
  }

  discardDocument(prepared: PreparedDocument): void {
    this.discarded.push(
      (prepared as PreparedDocument & { pages: string[] }).pages,
    );
  }

  emit<K extends "editstatechange" | "documentchange">(
    type: K,
    event: ViewerEventMap[K],
  ): void {
    this.events.push({ type, event });
    if (this.options.failEmit) throw new Error("listener exploded");
  }

  get eventTypes(): string[] {
    return this.events.map((entry) => entry.type);
  }

  get current(): string[] | undefined {
    return this.shown.at(-1);
  }
}

/*
 * A viewer adapter over the same JSON page format, so the viewer integration
 * can be tested end to end without a real document format.
 */

export interface FakeHandle {
  readonly id: number;
  readonly pages: readonly string[];
}

export interface FakeAdapterOptions {
  readonly edit?: EditEngineProvider;
  /** Implement `reopen` instead of opening from scratch. */
  readonly reopen?: boolean;
  /** Throw from `open`/`reopen` while true. */
  failOpen?: boolean;
}

export function fakeEditableAdapter(options: FakeAdapterOptions = {}) {
  let nextId = 1;
  const closed: number[] = [];
  const reopened: number[] = [];
  const adapter: DocumentAdapter<FakeHandle> & {
    readonly closed: number[];
    readonly reopened: number[];
  } = {
    id: "fake-editable",
    formats: ["pdf"],
    closed,
    reopened,
    async open(data) {
      if (options.failOpen) throw new Error("open refused");
      return { id: nextId++, pages: decodePages(data) };
    },
    async getInfo(handle) {
      return {
        format: "pdf",
        unit: "page",
        pageCount: handle.pages.length,
        pageSizes: handle.pages.map(() => ({ width: 100, height: 200 })),
      };
    },
    async render() {},
    async getTextMap(handle, pageIndex) {
      const words = handle.pages[pageIndex]!.split(" ");
      let offset = 0;
      return words.map((word, index) => {
        const text = index === words.length - 1 ? word : `${word} `;
        const run: TextRun = {
          text,
          x: index * 10,
          y: 0,
          width: 10,
          height: 10,
          logicalStart: offset,
          logicalEnd: offset + text.length,
        };
        offset += text.length;
        return run;
      });
    },
    close(handle) {
      closed.push(handle.id);
    },
    ...(options.edit ? { edit: options.edit } : {}),
    ...(options.reopen
      ? {
          async reopen(previous: FakeHandle, data: Uint8Array) {
            if (options.failOpen) throw new Error("reopen refused");
            reopened.push(previous.id);
            return { id: nextId++, pages: decodePages(data) };
          },
        }
      : {}),
  };
  return adapter;
}
