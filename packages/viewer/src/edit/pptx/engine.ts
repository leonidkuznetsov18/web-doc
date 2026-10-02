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
import { pptxHandlers } from "./handlers.js";
import {
  issueCollector,
  type PptxOperationContext,
  type PptxOperationResult,
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
  /** The package revision the index describes; moved forward when a commit touched slides only. */
  revision: number;
  readonly model: DeckModel;
  readonly placeholders: Map<string, Promise<PlaceholderTable>>;
  readonly slides: Map<string, Promise<SlideElements>>;
}

const SLIDE_PART = /^\/ppt\/slides\/slide(\d+)\.xml$/i;
const SLIDE_RELS = /^\/ppt\/slides\/_rels\/slide(\d+)\.xml\.rels$/i;

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
    if (!this.#inspection) return this.#startInspection(revision, signal);
    return this.#inspection.then((current) =>
      current.revision === revision
        ? current
        : this.#startInspection(revision, signal),
    );
  }

  /** Builds the index; a failed or aborted build is not kept. */
  #startInspection(
    revision: number,
    signal?: AbortSignal,
  ): Promise<Inspection> {
    const pending = this.#buildInspection(revision, signal);
    this.#inspection = pending;
    pending.catch(() => {
      if (this.#inspection === pending) this.#inspection = undefined;
    });
    return pending;
  }

  /**
   * After a commit: a change confined to slide parts keeps the deck index
   * (the presentation, every layout and master, every slide's
   * relationships); anything else — the presentation, parts added or
   * removed other than media — rebuilds it on the next read.
   */
  async #afterCommit(parts: PptxOperationResult["parts"]): Promise<void> {
    const current = this.#inspection
      ? await this.#inspection.catch(() => undefined)
      : undefined;
    if (!parts || !current) {
      this.#inspection = undefined;
      return;
    }
    for (const name of parts.changed)
      if (!SLIDE_PART.test(name) && !SLIDE_RELS.test(name)) {
        this.#inspection = undefined;
        return;
      }
    if (parts.added.length > 0 || parts.removed.length > 0) {
      const structural = [...parts.added, ...parts.removed].some(
        (name) => !/^\/ppt\/media\//i.test(name),
      );
      if (structural) {
        this.#inspection = undefined;
        return;
      }
    }
    // Every cached slide was scanned at the old revision; a patch on such a
    // scan is refused, so the slides are read again on demand while the
    // deck index (presentation, layouts, masters) is kept.
    current.slides.clear();
    current.revision = this.#pkg.revision;
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
    const base = await this.#context(0, signal);
    // Slide operations change the count the later operations see.
    let pageCount = base.pageCount;
    for (const [index, operation] of operations.entries()) {
      const context = { ...base, pageCount, operationIndex: index };
      switch (operation.op) {
        case "insertSlide":
        case "duplicateSlide":
          pageCount += 1;
          break;
        case "deleteSlide":
          pageCount = Math.max(1, pageCount - 1);
          break;
        default:
          break;
      }
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
      // A same-batch reference names an element that does not exist yet:
      // its target is checked when the batch is applied, the rest now.
      const target = (operation as { readonly target?: unknown }).target;
      const reference =
        typeof target === "string" ? parseReference(target) : undefined;
      if (reference !== undefined && reference >= index) {
        issues.push({
          operationIndex: index,
          path: "/target",
          code: "unknown-target",
          message: `"${target}" must refer to an earlier operation`,
        });
        continue;
      }
      const collect = issueCollector(index, issues);
      await handler.validate(
        operation as PptxOperation,
        context,
        reference === undefined
          ? collect
          : (path, code, message) => {
              if (!path.startsWith("/target") && !path.startsWith("/range"))
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
    const issuedIds = new Map(this.#issuedIds);
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
        await this.#afterCommit(result.parts);
        createdByOperation.push([...result.createdIds]);
        createdIds.push(...result.createdIds);
        removedIds.push(...(result.removedIds ?? []));
        // A removed id is never issued again on its slide.
        for (const id of result.removedIds ?? []) {
          const match = /^(sld\d+):(\d+)/.exec(id);
          if (!match) continue;
          const key = match[1]!;
          this.#issuedIds.set(
            key,
            Math.max(this.#issuedIds.get(key) ?? 0, Number(match[2])),
          );
        }
        for (const page of result.changedPages) changedPages.add(page);
        warnings.push(...result.warnings);
      }
    } catch (error) {
      // Everything the batch did, including the ids it issued, is undone.
      this.#pkg.restore(snapshot);
      this.#issuedIds = issuedIds;
      this.#inspection = undefined;
      this.#modelSync = undefined;
      throw error;
    }
    this.#pkg.release(snapshot);
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
  /** Highest `p:cNvPr` id issued per slide key in this session; replays rebuild it. */
  #issuedIds = new Map<string, number>();

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
      pageCount: model.pageCount,
      stateId,
      operationIndex,
      elements: (pageIndex) => this.#slideElements(pageIndex, signal),
      locate: (id) => this.#locate(id, signal),
      allocateShapeId: (elements) => this.#allocateShapeId(elements),
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
    this.#modelSync = undefined;
    this.#issuedIds = new Map();
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
      const started = this.#readSlide(inspection, slide, pageIndex, signal);
      pending = started;
      inspection.slides.set(slide.key, started);
      // An aborted or failed read must not stand in for the slide.
      started.catch(() => {
        if (inspection.slides.get(slide.key) === started)
          inspection.slides.delete(slide.key);
      });
    }
    return pending;
  }

  /** One above every id on the slide and every id issued there this session. */
  #allocateShapeId(elements: SlideElements): number {
    let max = 1;
    for (const record of elements.records) max = Math.max(max, record.cNvPrId);
    const tree = elements.part.find("spTree");
    if (tree)
      for (const node of elements.part.findAll("cNvPr", tree)) {
        const id = Number(elements.part.attribute(node, "id") ?? NaN);
        if (Number.isInteger(id)) max = Math.max(max, id);
      }
    const issued = this.#issuedIds.get(elements.slide.key) ?? 0;
    const next = Math.max(max, issued) + 1;
    this.#issuedIds.set(elements.slide.key, next);
    return next;
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
        const started = readPlaceholders(this.#pkg, part, signal);
        pending = started;
        inspection.placeholders.set(part, started);
        started.catch(() => {
          if (inspection.placeholders.get(part) === started)
            inspection.placeholders.delete(part);
        });
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
