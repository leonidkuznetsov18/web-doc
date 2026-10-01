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
import { invalidOperationError, parseReference } from "../operations.js";
import { OoxmlPackage } from "../ooxml/package.js";
import type {
  EditFindOptions,
  EditOperation,
  ElementQuery,
  OperationIssue,
  PagePoint,
  TextTarget,
} from "../types.js";
import {
  readPlaceholders,
  readSlideElements,
  type InheritanceSources,
  type PlaceholderTable,
  type ShapeRecord,
  type SlideElements,
} from "./elements.js";
import { frameContains, rectsIntersect } from "./geometry.js";
import { DeckModel, type SlideRecord } from "./model.js";
import {
  issueCollector,
  pptxHandlers,
  type PptxOperationContext,
} from "./operations.js";
import { pptxOperationSchemas } from "./schemas.js";
import type {
  PptxElement,
  PptxLayoutInfo,
  PptxOperation,
  PptxSlideInfo,
} from "./types.js";

/*
 * The PPTX engine: an OOXML package plus the deck index and the per-slide
 * element cache, both rebuilt when the package revision changes. Thread
 * agnostic: the worker handler and the Node tests drive it directly.
 */

/** The reads behind `getSlides` and `getLayouts`, beyond the core engine interface. */
export interface PptxEngineReads {
  slides(signal: AbortSignal): Promise<readonly PptxSlideInfo[]>;
  layouts(signal: AbortSignal): Promise<readonly PptxLayoutInfo[]>;
}

interface Inspection {
  readonly revision: number;
  readonly model: DeckModel;
  readonly placeholders: Map<string, Promise<PlaceholderTable>>;
  readonly slides: Map<string, Promise<SlideElements>>;
}

export class PptxEditEngine implements EditEngine, PptxEngineReads {
  readonly schemas = pptxOperationSchemas;
  readonly #original: Uint8Array;
  readonly #limits: ResourceLimits;
  readonly #assets = new AssetStore();
  #pkg: OoxmlPackage;
  #inspection: Promise<Inspection> | undefined;
  #disposed = false;

  private constructor(
    original: Uint8Array,
    pkg: OoxmlPackage,
    limits: ResourceLimits,
  ) {
    this.#original = original;
    this.#pkg = pkg;
    this.#limits = limits;
  }

  /** Opens the package and reads the deck index, so a broken deck fails here. */
  static async open(
    bytes: Uint8Array,
    limits: ResourceLimits,
    signal?: AbortSignal,
  ): Promise<PptxEditEngine> {
    const pkg = await OoxmlPackage.open(bytes, {
      limits,
      ...(signal ? { signal } : {}),
    });
    const engine = new PptxEditEngine(bytes, pkg, limits);
    await engine.model(signal);
    return engine;
  }

  /** The package behind the engine, for tests and the operations. */
  get package(): OoxmlPackage {
    return this.#pkg;
  }

  get pageCount(): number {
    return this.#modelSync?.pageCount ?? 0;
  }

  #modelSync: DeckModel | undefined;

  /** The deck index at the current revision. */
  async model(signal?: AbortSignal): Promise<DeckModel> {
    return (await this.#inspect(signal)).model;
  }

  #inspect(signal?: AbortSignal): Promise<Inspection> {
    this.#assertAlive();
    const revision = this.#pkg.revision;
    if (!this.#inspection) {
      this.#inspection = this.#buildInspection(revision, signal);
      return this.#inspection;
    }
    return this.#inspection.then((current) => {
      if (current.revision === revision) return current;
      this.#inspection = this.#buildInspection(revision, signal);
      return this.#inspection;
    });
  }

  async #buildInspection(
    revision: number,
    signal?: AbortSignal,
  ): Promise<Inspection> {
    const model = await DeckModel.load(this.#pkg, signal);
    this.#modelSync = model;
    return { revision, model, placeholders: new Map(), slides: new Map() };
  }

  async validate(
    operations: readonly EditOperation[],
    signal: AbortSignal,
  ): Promise<readonly OperationIssue[]> {
    const issues: OperationIssue[] = [];
    const context = await this.#context(0, signal);
    for (const [index, operation] of operations.entries()) {
      const handler = pptxHandlers.get(operation.op);
      if (!handler) {
        issues.push({
          operationIndex: index,
          path: "/op",
          code: "unknown-operation",
          message: `Unknown operation ${operation.op}`,
        });
        continue;
      }
      await handler.validate(
        operation as PptxOperation,
        { ...context, operationIndex: index },
        issueCollector(index, issues),
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
    const changedPages = new Set<number>();
    const warnings: ViewerWarning[] = [];
    const createdByOperation: string[][] = [];
    try {
      for (const [index, raw] of batch.operations.entries()) {
        throwIfAborted(signal);
        const operation = resolveReferences(raw, createdByOperation, index);
        const handler = pptxHandlers.get(operation.op);
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
        createdByOperation.push([...result.createdIds]);
        createdIds.push(...result.createdIds);
        removedIds.push(...(result.removedIds ?? []));
        for (const page of result.changedPages) changedPages.add(page);
        warnings.push(...result.warnings);
      }
    } catch (error) {
      this.#pkg.restore(snapshot);
      this.#inspection = undefined;
      throw error;
    }
    const model = await this.model(signal);
    return {
      createdIds,
      removedIds,
      changedPages: [...changedPages].sort((a, b) => a - b),
      pageCount: model.pageCount,
      warnings,
    };
  }

  #stateId = 0;

  #nextStateId(): number {
    this.#stateId += 1;
    return this.#stateId;
  }

  async #context(
    stateId: number,
    signal: AbortSignal | undefined,
    operationIndex = 0,
  ): Promise<PptxOperationContext> {
    const model = await this.model(signal);
    return {
      pkg: this.#pkg,
      model,
      limits: this.#limits,
      assets: this.#assets,
      stateId,
      operationIndex,
      elements: (pageIndex) => this.#slideElements(pageIndex, signal),
      locate: (id) => this.#locate(id, signal),
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
    this.#inspection = undefined;
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

  async getElements(
    query: ElementQuery,
    signal: AbortSignal,
  ): Promise<readonly PptxElement[]> {
    const model = await this.model(signal);
    const pages =
      query.pageIndex === undefined
        ? model.slides.map((_, index) => index)
        : [query.pageIndex];
    const out: PptxElement[] = [];
    for (const pageIndex of pages) {
      if (pageIndex < 0 || pageIndex >= model.pageCount) continue;
      const elements = await this.#slideElements(pageIndex, signal);
      for (const record of elements.records) {
        const element = record.element;
        if (query.kinds && !query.kinds.includes(element.kind)) continue;
        if (
          query.intersects &&
          !rectsIntersect(element.bounds, query.intersects)
        )
          continue;
        out.push(element);
      }
    }
    return out;
  }

  async getElement(
    id: string,
    signal: AbortSignal,
  ): Promise<PptxElement | undefined> {
    return (await this.#locate(id, signal))?.element;
  }

  async elementsAt(
    pageIndex: number,
    point: PagePoint,
    signal: AbortSignal,
  ): Promise<readonly PptxElement[]> {
    const model = await this.model(signal);
    if (pageIndex < 0 || pageIndex >= model.pageCount) return [];
    const elements = await this.#slideElements(pageIndex, signal);
    const hits: PptxElement[] = [];
    for (let index = elements.records.length - 1; index >= 0; index -= 1) {
      const record = elements.records[index]!;
      if (record.placed && frameContains(record.placed, point))
        hits.push(record.element);
    }
    return hits;
  }

  async findText(
    query: string,
    options: EditFindOptions,
    signal: AbortSignal,
  ): Promise<readonly TextTarget[]> {
    if (query.length === 0) return [];
    const model = await this.model(signal);
    const [first, last] = options.pageRange ?? [0, model.pageCount - 1];
    const fold = (value: string): string =>
      options.caseSensitive ? value : value.toLocaleLowerCase();
    const needle = fold(query);
    const targets: TextTarget[] = [];
    const limit = options.maxResults ?? Number.POSITIVE_INFINITY;
    for (
      let pageIndex = Math.max(0, first);
      pageIndex <= Math.min(last, model.pageCount - 1);
      pageIndex += 1
    ) {
      const elements = await this.#slideElements(pageIndex, signal);
      for (const record of elements.records) {
        const text = record.element.text;
        if (!text) continue;
        const haystack = fold(text);
        let from = 0;
        while (targets.length < limit) {
          const at = haystack.indexOf(needle, from);
          if (at < 0) break;
          const end = at + query.length;
          targets.push({
            pageIndex,
            text: text.slice(at, end),
            rects: [record.element.bounds],
            elementIds: [record.element.id],
            ranges: [
              {
                start: { elementId: record.element.id, offset: at },
                end: { elementId: record.element.id, offset: end },
              },
            ],
          });
          from = end;
        }
        if (targets.length >= limit) return targets;
      }
    }
    return targets;
  }

  async slides(signal: AbortSignal): Promise<readonly PptxSlideInfo[]> {
    const model = await this.model(signal);
    const out: PptxSlideInfo[] = [];
    for (const [pageIndex, slide] of model.slides.entries())
      out.push({
        pageIndex,
        key: slide.key,
        layout: slide.layout?.id ?? "",
        hidden: await model.hidden(slide, signal),
      });
    return out;
  }

  async layouts(signal: AbortSignal): Promise<readonly PptxLayoutInfo[]> {
    const model = await this.model(signal);
    return model.layouts.map((layout) => ({
      id: layout.id,
      name: layout.name,
      ...(layout.type ? { type: layout.type } : {}),
      master: layout.master.id,
    }));
  }

  async dispose(): Promise<void> {
    this.#disposed = true;
    this.#inspection = undefined;
  }

  async #slideElements(
    pageIndex: number,
    signal?: AbortSignal,
  ): Promise<SlideElements> {
    const inspection = await this.#inspect(signal);
    const slide = inspection.model.slideAt(pageIndex);
    let pending = inspection.slides.get(slide.key);
    if (!pending) {
      pending = this.#readSlide(inspection, slide, pageIndex, signal);
      inspection.slides.set(slide.key, pending);
    }
    return pending;
  }

  async #readSlide(
    inspection: Inspection,
    slide: SlideRecord,
    pageIndex: number,
    signal?: AbortSignal,
  ): Promise<SlideElements> {
    const table = (
      part: string | undefined,
    ): Promise<PlaceholderTable> | undefined => {
      if (!part) return undefined;
      let pending = inspection.placeholders.get(part);
      if (!pending) {
        pending = readPlaceholders(this.#pkg, part, signal);
        inspection.placeholders.set(part, pending);
      }
      return pending;
    };
    const layout = await table(slide.layout?.part);
    const master = await table(slide.layout?.master.part);
    const presentation = inspection.model.presentation;
    const defaultsNode = presentation.find("defaultTextStyle");
    const sources: InheritanceSources = {
      ...(layout ? { layout } : {}),
      ...(master ? { master } : {}),
      ...(defaultsNode
        ? { defaults: { part: presentation, node: defaultsNode } }
        : {}),
    };
    return readSlideElements(
      this.#pkg,
      inspection.model,
      slide,
      pageIndex,
      sources,
      signal,
    );
  }

  async #locate(
    id: string,
    signal?: AbortSignal,
  ): Promise<ShapeRecord | undefined> {
    const key = id.split(":")[0];
    if (!key) return undefined;
    const model = await this.model(signal);
    const pageIndex = model.slides.findIndex((slide) => slide.key === key);
    if (pageIndex < 0) return undefined;
    return (await this.#slideElements(pageIndex, signal)).byId(id);
  }

  #assertAlive(): void {
    if (this.#disposed)
      throw new ViewerError("lifecycle-error", "The PPTX engine was disposed");
  }
}

/** Replaces `"$<n>"` targets with the first id operation `n` created. */
function resolveReferences(
  operation: EditOperation,
  created: readonly (readonly string[])[],
  index: number,
): PptxOperation {
  const target = (operation as { readonly target?: unknown }).target;
  if (typeof target !== "string") return operation as PptxOperation;
  const reference = parseReference(target);
  if (reference === undefined) return operation as PptxOperation;
  const id = reference < index ? created[reference]?.[0] : undefined;
  if (!id)
    throw new ViewerError(
      "invalid-operation",
      `Operation ${index} refers to "${target}", which created nothing`,
      { details: { operationIndex: index, target } },
    );
  return { ...operation, target: id } as PptxOperation;
}

function throwIfAborted(signal: AbortSignal | undefined): void {
  if (signal?.aborted)
    throw new ViewerError("aborted", "The operation was aborted");
}
