import type { ResourceLimits, ViewerWarning } from "../../contracts.js";
import { ViewerError } from "../../errors.js";
import { AssetStore } from "../assets.js";
import type {
  EditEngine,
  EngineBatch,
  EngineChange,
  MaterializedDocument,
  MaterializeOptions,
  RestoreTarget,
} from "../engine.js";
import { OoxmlPackage } from "../ooxml/package.js";
import { invalidOperationError, parseReference } from "../operations.js";
import type {
  EditFindOptions,
  EditOperation,
  ElementQuery,
  OperationIssue,
  PagePoint,
  TextTarget,
} from "../types.js";
import { NO_PAGE, toElement } from "./elements.js";
import { docxHandlers } from "./handlers.js";
import { DocxModel, type AnyRecord } from "./model.js";
import { issueCollector, type DocxOperationContext } from "./operations.js";
import { docxOperationSchemas } from "./schemas.js";
import type { DocxElement, DocxOperation } from "./types.js";

/*
 * The DOCX edit engine: the package layer under a block index of the body
 * story. It runs inside the OOXML edit worker in the browser and directly
 * in Node tests. It never lays out: elements carry no geometry, and the
 * session joins the renderer's runs on the main thread.
 */

export class DocxEditEngine implements EditEngine {
  readonly schemas = docxOperationSchemas;
  readonly #original: Uint8Array;
  readonly #limits: ResourceLimits;
  readonly #assets = new AssetStore();
  #pkg: OoxmlPackage;
  #model: Promise<DocxModel> | undefined;
  #disposed = false;
  #stateId = 0;

  private constructor(
    original: Uint8Array,
    pkg: OoxmlPackage,
    limits: ResourceLimits,
  ) {
    this.#original = original;
    this.#pkg = pkg;
    this.#limits = limits;
  }

  /** Opens the package and reads the block index, so a broken document fails here. */
  static async open(
    bytes: Uint8Array,
    limits: ResourceLimits,
    signal?: AbortSignal,
  ): Promise<DocxEditEngine> {
    const pkg = await OoxmlPackage.open(bytes, {
      limits,
      ...(signal ? { signal } : {}),
    });
    const engine = new DocxEditEngine(bytes, pkg, limits);
    await engine.model(signal);
    return engine;
  }

  /** The package behind the engine, for tests and the operations. */
  get package(): OoxmlPackage {
    return this.#pkg;
  }

  /** The engine has no pages of its own; the renderer counts them. */
  get pageCount(): number {
    return 0;
  }

  /** The block index at the current revision. */
  model(signal?: AbortSignal): Promise<DocxModel> {
    this.#assertAlive();
    const revision = this.#pkg.revision;
    if (!this.#model) return this.#startModel(revision, signal);
    return this.#model.then((current) =>
      current.revision === revision
        ? current
        : this.#startModel(revision, signal),
    );
  }

  #startModel(revision: number, signal?: AbortSignal): Promise<DocxModel> {
    void revision;
    const pending = DocxModel.load(this.#pkg, signal);
    this.#model = pending;
    pending.catch(() => {
      if (this.#model === pending) this.#model = undefined;
    });
    return pending;
  }

  async validate(
    operations: readonly EditOperation[],
    signal: AbortSignal,
  ): Promise<readonly OperationIssue[]> {
    const issues: OperationIssue[] = [];
    const base = await this.#context(0, signal);
    for (const [index, operation] of operations.entries()) {
      const context = { ...base, operationIndex: index };
      const handler = docxHandlers.get(operation.op);
      if (!handler) {
        issues.push({
          operationIndex: index,
          path: "/op",
          code: "unknown-operation",
          message: `Unknown operation ${operation.op}`,
        });
        continue;
      }
      // A same-batch reference names an element that does not exist yet:
      // its target is checked when the batch is applied, the rest now.
      const referenced = referenceFields(operation);
      const forward = referenced.find((entry) => entry.reference >= index);
      if (forward) {
        issues.push({
          operationIndex: index,
          path: `/${forward.field}`,
          code: "unknown-target",
          message: `"${forward.value}" must refer to an earlier operation`,
        });
        continue;
      }
      const collect = issueCollector(index, issues);
      const skipped = new Set(referenced.map((entry) => `/${entry.field}`));
      await handler.validate(
        operation as DocxOperation,
        context,
        referenced.length === 0
          ? collect
          : (path, code, message) => {
              if (
                ![...skipped].some(
                  (prefix) => path === prefix || path.startsWith(`${prefix}/`),
                ) &&
                !path.startsWith("/range")
              )
                collect(path, code, message);
            },
      );
    }
    return issues;
  }

  /** Plain operation arrays, as the unit tests pass them, become the next batch. */
  async apply(
    input: EngineBatch | readonly EditOperation[],
    signal: AbortSignal,
  ): Promise<EngineChange> {
    const batch: EngineBatch = Array.isArray(input)
      ? { stateId: this.#nextStateId(), operations: input }
      : (input as EngineBatch);
    this.#stateId = Math.max(this.#stateId, batch.stateId);
    const snapshot = this.#pkg.snapshot();
    const createdIds: string[] = [];
    const removedIds: string[] = [];
    const warnings: ViewerWarning[] = [];
    const createdByOperation: string[][] = [];
    try {
      for (const [index, raw] of batch.operations.entries()) {
        throwIfAborted(signal);
        const operation = resolveReferences(raw, createdByOperation, index);
        const handler = docxHandlers.get(operation.op);
        if (!handler)
          throw new ViewerError(
            "invalid-operation",
            `Unknown operation ${operation.op}`,
          );
        const context = await this.#context(batch.stateId, signal, index);
        const issues: OperationIssue[] = [];
        await handler.validate(
          operation,
          context,
          issueCollector(index, issues),
        );
        if (issues.length > 0) throw invalidOperationError(issues);
        const result = await handler.apply(operation, context);
        this.#model = undefined;
        createdByOperation.push([...result.createdIds]);
        createdIds.push(...result.createdIds);
        removedIds.push(...(result.removedIds ?? []));
        warnings.push(...result.warnings);
      }
    } catch (error) {
      // Everything the batch did is undone.
      this.#pkg.restore(snapshot);
      this.#model = undefined;
      throw error;
    }
    this.#pkg.release(snapshot);
    await this.model(signal);
    return {
      createdIds,
      removedIds,
      // A flow document reflows from the first changed element onwards;
      // naming its pages needs the renderer's layout, which the operations
      // of the next tasks report through the session.
      changedPages: [],
      warnings,
    };
  }

  #nextStateId(): number {
    this.#stateId += 1;
    return this.#stateId;
  }

  async #context(
    stateId: number,
    signal: AbortSignal | undefined,
    operationIndex = 0,
  ): Promise<DocxOperationContext> {
    const model = await this.model(signal);
    return {
      pkg: this.#pkg,
      model,
      limits: this.#limits,
      assets: this.#assets,
      stateId,
      operationIndex,
    };
  }

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
    _options: MaterializeOptions = {},
    signal: AbortSignal = new AbortController().signal,
  ): Promise<MaterializedDocument> {
    this.#assertAlive();
    const own =
      purposeOrSignal instanceof AbortSignal ? purposeOrSignal : signal;
    return { bytes: await this.#pkg.save({}, own), warnings: [] };
  }

  async restore(
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
    this.#assertAlive();
    this.#pkg = await OoxmlPackage.open(target.base ?? this.#original, {
      limits: this.#limits,
      signal,
    });
    this.#model = undefined;
    await this.model(signal);
    for (const batch of target.batches) await this.apply(batch, signal);
  }

  async putAsset(
    id: string,
    data: Uint8Array,
    _signal: AbortSignal,
  ): Promise<void> {
    this.#assets.set(id, data);
  }

  /**
   * Every element of the body story, in document order, without geometry.
   * `pageIndex` and `intersects` are the session's to apply once the runs
   * are joined; `kinds` filters here.
   */
  async getElements(
    query: ElementQuery,
    signal: AbortSignal,
  ): Promise<readonly DocxElement[]> {
    const model = await this.model(signal);
    const out: DocxElement[] = [];
    for (const record of model.records) {
      if (query.kinds && !query.kinds.includes(record.kind)) continue;
      out.push(toElement(model, record));
    }
    return out;
  }

  async getElement(
    id: string,
    signal: AbortSignal,
  ): Promise<DocxElement | undefined> {
    const record = await this.locate(id, signal);
    return record ? toElement(await this.model(signal), record) : undefined;
  }

  /** The record an element id names, at the current revision. */
  async locate(
    id: string,
    signal?: AbortSignal,
  ): Promise<AnyRecord | undefined> {
    return (await this.model(signal)).byId.get(id);
  }

  /** Hit-testing needs the renderer's geometry; the session answers it. */
  async elementsAt(
    _pageIndex: number,
    _point: PagePoint,
    _signal: AbortSignal,
  ): Promise<readonly DocxElement[]> {
    return [];
  }

  /**
   * Matches in paragraph and table text, in document order, without
   * rectangles or pages: the session adds those from the renderer's runs.
   */
  async findText(
    query: string,
    options: EditFindOptions,
    signal: AbortSignal,
  ): Promise<readonly TextTarget[]> {
    if (query.length === 0) return [];
    const model = await this.model(signal);
    const fold = (value: string): string =>
      options.caseSensitive ? value : value.toLocaleLowerCase();
    const needle = fold(query);
    const limit = options.maxResults ?? Number.POSITIVE_INFINITY;
    const targets: TextTarget[] = [];
    for (const record of model.records) {
      if (record.kind !== "paragraph") continue;
      const text = record.text.text;
      if (!text) continue;
      const haystack = fold(text);
      let from = 0;
      while (targets.length < limit) {
        const at = haystack.indexOf(needle, from);
        if (at < 0) break;
        const end = at + query.length;
        targets.push({
          pageIndex: NO_PAGE,
          text: text.slice(at, end),
          rects: [],
          elementIds: [record.elementId],
          ranges: [
            {
              start: { elementId: record.elementId, offset: at },
              end: { elementId: record.elementId, offset: end },
            },
          ],
        });
        from = end;
      }
      if (targets.length >= limit) break;
    }
    return targets;
  }

  async dispose(): Promise<void> {
    this.#disposed = true;
    this.#model = undefined;
  }

  #assertAlive(): void {
    if (this.#disposed)
      throw new ViewerError("lifecycle-error", "The DOCX engine was disposed");
  }
}

const REFERENCE_FIELDS = ["target", "before", "after"] as const;

function referenceFields(
  operation: EditOperation,
): { field: string; value: string; reference: number }[] {
  const out: { field: string; value: string; reference: number }[] = [];
  for (const field of REFERENCE_FIELDS) {
    const value = (operation as unknown as Record<string, unknown>)[field];
    if (typeof value !== "string") continue;
    const reference = parseReference(value);
    if (reference !== undefined) out.push({ field, value, reference });
  }
  return out;
}

/** Replaces `"$<n>"` references with the first id operation `n` created. */
function resolveReferences(
  operation: EditOperation,
  created: readonly (readonly string[])[],
  index: number,
): DocxOperation {
  let resolved = operation as unknown as Record<string, unknown>;
  for (const { field, value, reference } of referenceFields(operation)) {
    const id = reference < index ? created[reference]?.[0] : undefined;
    if (!id)
      throw new ViewerError(
        "invalid-operation",
        `Operation ${index} refers to "${value}", which created nothing`,
        { details: { operationIndex: index, target: value } },
      );
    resolved = { ...resolved, [field]: id };
  }
  return resolved as unknown as DocxOperation;
}

function throwIfAborted(signal: AbortSignal | undefined): void {
  if (signal?.aborted)
    throw new ViewerError("aborted", "The operation was aborted");
}
