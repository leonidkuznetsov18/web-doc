import type { TextRun } from "../../contracts.js";
import { readDescription, readOutline } from "../ai/outline.js";
import { resolveTargets } from "../ai/targets.js";
import { buildToolSet, callTool as runTool } from "../ai/tools.js";
import type {
  DescribeOptions,
  DocumentDescription,
  EditCheckpoint,
  OutlineOptions,
  OutlineResult,
  TargetCandidate,
  TargetQuery,
  ToolCall,
  ToolCallOptions,
  ToolResult,
  ToolSet,
} from "../ai/types.js";
import { ViewerError } from "../../errors.js";
import type {
  EditEngine,
  EditSessionAccess,
  EditSessionCore,
} from "../engine.js";
import type { DocxEngineReads } from "./engine.js";
import {
  canvasMeasurer,
  caretAt,
  layParagraph,
  unionRects,
  type LaidLine,
  type TextMeasurer,
} from "./layout.js";
import {
  checkOperations,
  freezeOperations,
  invalidOperationError,
} from "../operations.js";
import { docxOperationSchemas } from "./schemas.js";
import { spanOnTarget } from "../range-style.js";
import type {
  ApplyOptions,
  AssetOptions,
  EditFindOptions,
  EditOperation,
  EditReceipt,
  EditState,
  ElementFragment,
  ElementQuery,
  HistoryOptions,
  OperationSchemaSet,
  PagePoint,
  PageRect,
  ReadItem,
  ReadOptions,
  ReadResult,
  SavedDocument,
  TextLayout,
  TextPosition,
  TextRange,
  TextTarget,
} from "../types.js";
import type {
  DocxDeleteElementOperation,
  DocxEditSession,
  DocxElement,
  DocxFields,
  DocxInsertImageOperation,
  DocxInsertParagraphOperation,
  DocxInsertTableOperation,
  DocxMoveElementOperation,
  DocxOperation,
  DocxReplaceTextOperation,
  DocxRevision,
  DocxSaveOptions,
  DocxSetParagraphStyleOperation,
  DocxSetTableCellOperation,
  DocxSetTextStyleOperation,
  DocxTextStyle,
  DocxTextPreview,
  DocxTextPreviewOptions,
} from "./types.js";

/*
 * The DOCX session: the core session narrowed to DOCX operations and
 * elements, plus the geometry join. The engine never lays out, so the
 * session places elements with the renderer's text runs: every run carries
 * the id of its paragraph, a paragraph's fragments are the unions of its
 * runs' boxes per page, a table's are its cell paragraphs' unions, an
 * inline object takes its paragraph's. Pages the viewer has not laid out
 * are not read for a query without `pageIndex`.
 */

/** Per paragraph id (eight hex digits), the union of its runs per page. */
type Placement = ReadonlyMap<string, ReadonlyMap<number, PageRect>>;

export class DocxSession implements DocxEditSession {
  readonly format = "docx" as const;
  readonly #core: EditSessionCore;
  #tools: ToolSet | undefined;
  readonly #access: EditSessionAccess;
  /** Measures the runs' fonts; null where no canvas exists. */
  #measurer: TextMeasurer | null | undefined;

  constructor(core: EditSessionCore, access: EditSessionAccess) {
    this.#core = core;
    this.#access = access;
  }

  get sessionId(): string {
    return this.#core.sessionId;
  }

  get state(): EditState {
    return this.#core.state;
  }

  get schemas(): OperationSchemaSet {
    return this.#core.schemas;
  }

  apply(
    operations: readonly DocxOperation[],
    options?: ApplyOptions,
  ): Promise<EditReceipt> {
    return this.#core.apply(operations, options);
  }

  applyJson(
    operations: readonly EditOperation[],
    options?: ApplyOptions,
  ): Promise<EditReceipt> {
    return this.#core.applyJson(operations, options);
  }

  undo(options?: HistoryOptions): Promise<EditReceipt> {
    return this.#core.undo(options);
  }

  redo(options?: HistoryOptions): Promise<EditReceipt> {
    return this.#core.redo(options);
  }

  reset(options?: HistoryOptions): Promise<EditReceipt> {
    return this.#core.reset(options);
  }

  save(options?: DocxSaveOptions): Promise<SavedDocument> {
    return this.#core.save(options);
  }

  markSaved(stateToken: string): void {
    this.#core.markSaved(stateToken);
  }

  addAsset(data: Uint8Array, options?: AssetOptions): Promise<string> {
    return this.#core.addAsset(data, options);
  }

  getElements(
    query: ElementQuery = {},
    options?: ReadOptions,
  ): Promise<ReadResult<DocxElement>> {
    return this.#core.readItems(options, async (engine, signal) => {
      const elements = (await engine.getElements(
        {},
        signal,
      )) as readonly DocxElement[];
      const pages =
        query.pageIndex === undefined
          ? this.#access.cachedPages()
          : [query.pageIndex];
      const placement = await this.#placement(pages, signal);
      return placeElements(elements, placement).filter((element) =>
        matchesQuery(element, query),
      );
    });
  }

  getElement(
    id: string,
    options?: ReadOptions,
  ): Promise<ReadItem<DocxElement>> {
    return this.#core.readItem(options, async (engine, signal) => {
      // A table is placed through its cell paragraphs, so every element is
      // read and the one asked for picked out of the joined list.
      const elements = (await engine.getElements(
        {},
        signal,
      )) as readonly DocxElement[];
      if (!elements.some((element) => element.id === id)) return undefined;
      const placement = await this.#placement(
        this.#access.cachedPages(),
        signal,
      );
      return placeElements(elements, placement).find(
        (element) => element.id === id,
      );
    });
  }

  elementsAt(
    pageIndex: number,
    point: PagePoint,
    options?: ReadOptions,
  ): Promise<ReadResult<DocxElement>> {
    return this.#core.readItems(options, async (engine, signal) => {
      const runs = await this.#access.getTextRuns(pageIndex, signal);
      const hit = new Set<string>();
      for (const run of runs)
        if (run.paragraphId && rectContains(runRect(run), point))
          hit.add(`p:${run.paragraphId}`);
      if (hit.size === 0) return [];
      const elements = (await engine.getElements(
        {},
        signal,
      )) as readonly DocxElement[];
      const placed = placeElements(
        elements,
        await this.#placement([pageIndex], signal),
      );
      const paragraphs = placed.filter((element) => hit.has(element.id));
      const tables = new Set(
        paragraphs
          .map((element) => element.parentId)
          .filter((id): id is string => id !== undefined),
      );
      return [
        ...paragraphs,
        ...placed.filter((element) => tables.has(element.id)),
      ];
    });
  }

  async findText(
    query: string,
    options: EditFindOptions = {},
  ): Promise<ReadResult<TextTarget>> {
    // A page range bounds the matches by where they are placed, which
    // only the geometry knows: the engine then searches without a limit
    // and the limit applies to the matches on those pages.
    const { pageRange } = options;
    const engineOptions = { ...options };
    if (pageRange) delete engineOptions.maxResults;
    const found = await this.#core.findText(query, engineOptions);
    const placed = await this.#placeTargets(found.items, options);
    const items = pageRange
      ? placed
          .filter(
            (target) =>
              target.pageIndex >= pageRange[0] &&
              target.pageIndex <= pageRange[1],
          )
          .slice(0, options.maxResults ?? placed.length)
      : placed;
    return Object.freeze({ ...found, items });
  }

  getOutline(options?: OutlineOptions): Promise<OutlineResult> {
    return readOutline(this, this.#core.limits, options);
  }

  describe(options?: DescribeOptions): Promise<ReadItem<DocumentDescription>> {
    return readDescription(this, this.#core.limits, options);
  }

  resolveTargets(
    query: TargetQuery,
    options?: ReadOptions,
  ): Promise<ReadResult<TargetCandidate>> {
    return resolveTargets(this, query, options);
  }

  createCheckpoint(label?: string): Promise<EditCheckpoint> {
    return this.#core.createCheckpoint(label);
  }

  listCheckpoints(): readonly EditCheckpoint[] {
    return this.#core.listCheckpoints();
  }

  restoreCheckpoint(
    id: string,
    options?: HistoryOptions,
  ): Promise<EditReceipt> {
    return this.#core.restoreCheckpoint(id, options);
  }

  dropCheckpoint(id: string): void {
    this.#core.dropCheckpoint(id);
  }

  get tools(): ToolSet {
    return (this.#tools ??= buildToolSet(this.format, this.schemas));
  }

  callTool(call: ToolCall, options?: ToolCallOptions): Promise<ToolResult> {
    return runTool(this, this.tools, call, options);
  }

  getRevisions(
    elementId: string,
    options?: ReadOptions,
  ): Promise<ReadResult<DocxRevision>> {
    return this.#core.readItems(options, (engine, signal) =>
      docxReads(engine).revisions(elementId, signal),
    );
  }

  async getTextStyle(
    fields: { readonly target: string; readonly range?: TextRange },
    options?: ReadOptions,
  ): Promise<ReadItem<Partial<DocxTextStyle>>> {
    const span = spanOnTarget(fields.target, fields.range);
    return this.#core.readItem(options, (engine, signal) =>
      docxReads(engine).textStyle(fields.target, span, signal),
    );
  }

  getTextLayout(
    elementId: string,
    options?: ReadOptions,
  ): Promise<ReadItem<TextLayout>> {
    return this.#core.readItem(options, async (engine, signal) => {
      const paragraph = await paragraphOf(engine, elementId, signal);
      // A paragraph without visible text draws no run on any page: no
      // page needs reading to say so.
      if (!paragraph || INVISIBLE.test(paragraph.text ?? "")) return undefined;
      const runs = this.#runLoader(signal);
      const holds = async (pageIndex: number): Promise<boolean> =>
        (await runs(pageIndex)).some(
          (run) => run.paragraphId === paragraphIdOf(elementId),
        );
      const pageCount = this.#core.state.pageCount;
      const candidates = [
        ...[...this.#access.cachedPages()].sort((a, b) => a - b),
        ...range(0, pageCount - 1, pageCount),
      ];
      const pageIndex = await firstPageHolding(candidates, holds);
      if (pageIndex === undefined) return undefined;
      const lines = this.#lay(paragraph, [
        { pageIndex, runs: await runs(pageIndex) },
      ]);
      if (lines.length === 0) return undefined;
      const colors = await docxReads(engine).textColors(
        elementId,
        lines.map((line) => ({
          start: line.range.start.offset,
          end: line.range.end.offset,
        })),
        signal,
      );
      return {
        elementId,
        pageIndex,
        frame: unionRects(lines.map((line) => line.box)),
        lines: lines.map(
          ({ pageIndex: _page, box: _box, ends: _ends, ...line }, index) => ({
            ...line,
            color: colors?.[index] ?? "#000000",
          }),
        ),
      };
    });
  }

  positionAt(
    pageIndex: number,
    point: PagePoint,
    options?: ReadOptions,
  ): Promise<ReadItem<TextPosition>> {
    return this.#core.readItem(options, async (engine, signal) => {
      const runs = this.#runLoader(signal);
      const ids = new Set(
        (await runs(pageIndex))
          .map((run) => run.paragraphId)
          .filter((id): id is string => id !== undefined),
      );
      if (ids.size === 0) return undefined;
      const paragraphs = (await engine.getElements(
        { kinds: ["paragraph"] },
        signal,
      )) as readonly DocxElement[];
      const lines: LaidLine[] = [];
      for (const paragraph of paragraphs) {
        const paragraphId = paragraphIdOf(paragraph.id);
        if (!ids.has(paragraphId)) continue;
        // Offsets continue from the pages the paragraph starts on.
        const pages = [pageIndex];
        while (
          pages[0]! > 0 &&
          (await runs(pages[0]! - 1)).some(
            (run) => run.paragraphId === paragraphId,
          )
        )
          pages.unshift(pages[0]! - 1);
        const laid = this.#lay(
          paragraph,
          await Promise.all(
            pages.map(async (page) => ({
              pageIndex: page,
              runs: await runs(page),
            })),
          ),
        );
        lines.push(...laid.filter((line) => line.pageIndex === pageIndex));
      }
      return caretAt(lines, point);
    });
  }

  /** Text runs per page, each page read once per call. */
  #runLoader(
    signal: AbortSignal,
  ): (pageIndex: number) => Promise<readonly TextRun[]> {
    const pages = new Map<number, Promise<readonly TextRun[]>>();
    return (pageIndex) => {
      let runs = pages.get(pageIndex);
      if (!runs) {
        runs = this.#access.getTextRuns(pageIndex, signal);
        pages.set(pageIndex, runs);
      }
      return runs;
    };
  }

  /** The lines a paragraph's runs on the given pages draw. */
  #lay(
    paragraph: DocxElement,
    pages: readonly {
      readonly pageIndex: number;
      readonly runs: readonly TextRun[];
    }[],
  ): LaidLine[] {
    const paragraphId = paragraphIdOf(paragraph.id);
    const placed = pages.flatMap(({ pageIndex, runs }) =>
      runs
        .filter((run) => run.paragraphId === paragraphId)
        .map((run) => ({ run, pageIndex })),
    );
    this.#measurer ??= canvasMeasurer() ?? null;
    return layParagraph(
      paragraph.id,
      paragraph.text ?? "",
      placed,
      {
        fontFamily: paragraph.textStyle?.fontFamily ?? "",
        fontSize: paragraph.textStyle?.fontSize ?? DEFAULT_FONT_SIZE,
      },
      this.#measurer ?? undefined,
    );
  }

  previewText(
    fields: DocxFields<DocxReplaceTextOperation>,
    options?: ReadOptions,
  ): Promise<ReadItem<Uint8Array>> {
    const draft = frozenPreviewFields(fields);
    return this.#core.readItem(options, (engine, signal) =>
      docxReads(engine).previewText(draft, signal),
    );
  }

  previewTextPages(
    fields: DocxFields<DocxReplaceTextOperation>,
    render: DocxTextPreviewOptions,
    options?: ReadOptions,
  ): Promise<ReadItem<DocxTextPreview>> {
    const draft = frozenPreviewFields(fields);
    const targets = {
      ...render,
      pages: render.pages.map((page) => ({ ...page })),
    };
    return this.#core.readItem(options, async (engine, signal) => {
      const previewDocument = this.#access.previewDocument;
      if (!previewDocument)
        throw new ViewerError(
          "edit-unsupported",
          "Draft rendering is unavailable",
        );
      const document = await docxReads(engine).previewDraft(draft, signal);
      const pages = await previewDocument(document.bytes, targets, signal);
      const paragraph = document.paragraph;
      const page = [...pages.pages]
        .sort((left, right) => left.pageIndex - right.pageIndex)
        .find((page) =>
          page.runs.some(
            (run) => run.paragraphId === paragraphIdOf(draft.target),
          ),
        );
      if (!paragraph || paragraph.kind !== "paragraph" || !page) return pages;
      const lines = this.#lay(paragraph, [page]);
      if (lines.length === 0) return pages;
      const color =
        typeof paragraph.textStyle?.color === "string" &&
        paragraph.textStyle.color.startsWith("#")
          ? paragraph.textStyle.color
          : "#000000";
      return {
        ...pages,
        layout: {
          elementId: draft.target,
          pageIndex: page.pageIndex,
          frame: unionRects(lines.map((line) => line.box)),
          lines: lines.map(
            ({ pageIndex: _page, box: _box, ends: _ends, ...line }) => ({
              ...line,
              color,
            }),
          ),
        },
      };
    });
  }

  replaceText(
    fields: DocxFields<DocxReplaceTextOperation>,
    options?: ApplyOptions,
  ): Promise<EditReceipt> {
    return this.apply([{ op: "replaceText", ...fields }], options);
  }

  setTextStyle(
    fields: DocxFields<DocxSetTextStyleOperation>,
    options?: ApplyOptions,
  ): Promise<EditReceipt> {
    return this.apply([{ op: "setTextStyle", ...fields }], options);
  }

  setParagraphStyle(
    fields: DocxFields<DocxSetParagraphStyleOperation>,
    options?: ApplyOptions,
  ): Promise<EditReceipt> {
    return this.apply([{ op: "setParagraphStyle", ...fields }], options);
  }

  insertParagraph(
    fields: DocxFields<DocxInsertParagraphOperation>,
    options?: ApplyOptions,
  ): Promise<EditReceipt> {
    return this.apply([{ op: "insertParagraph", ...fields }], options);
  }

  deleteElement(
    fields: DocxFields<DocxDeleteElementOperation>,
    options?: ApplyOptions,
  ): Promise<EditReceipt> {
    return this.apply([{ op: "deleteElement", ...fields }], options);
  }

  moveElement(
    fields: DocxFields<DocxMoveElementOperation>,
    options?: ApplyOptions,
  ): Promise<EditReceipt> {
    return this.apply([{ op: "moveElement", ...fields }], options);
  }

  insertTable(
    fields: DocxFields<DocxInsertTableOperation>,
    options?: ApplyOptions,
  ): Promise<EditReceipt> {
    return this.apply([{ op: "insertTable", ...fields }], options);
  }

  setTableCell(
    fields: DocxFields<DocxSetTableCellOperation>,
    options?: ApplyOptions,
  ): Promise<EditReceipt> {
    return this.apply([{ op: "setTableCell", ...fields }], options);
  }

  insertImage(
    fields: DocxFields<DocxInsertImageOperation>,
    options?: ApplyOptions,
  ): Promise<EditReceipt> {
    return this.apply([{ op: "insertImage", ...fields }], options);
  }

  /** The union of every paragraph's runs on the given pages. */
  async #placement(
    pages: readonly number[],
    signal: AbortSignal,
  ): Promise<Placement> {
    const placement = new Map<string, Map<number, PageRect>>();
    for (const pageIndex of [...new Set(pages)].sort((a, b) => a - b)) {
      const runs = await this.#access.getTextRuns(pageIndex, signal);
      for (const run of runs) {
        if (!run.paragraphId) continue;
        let pages = placement.get(run.paragraphId);
        if (!pages) {
          pages = new Map();
          placement.set(run.paragraphId, pages);
        }
        const current = pages.get(pageIndex);
        pages.set(
          pageIndex,
          current ? union(current, runRect(run)) : runRect(run),
        );
      }
    }
    return placement;
  }

  /**
   * Rectangles and pages for text matches: the runs of the match's
   * paragraph on the page that holds it, cut to the matched characters
   * when the run text can be aligned with the paragraph text, else the
   * paragraph's runs on its first page. Pages are read from the viewer's
   * cache first, then from the requested range, then in order.
   */
  async #placeTargets(
    targets: readonly TextTarget[],
    options: EditFindOptions,
  ): Promise<readonly TextTarget[]> {
    if (targets.length === 0) return targets;
    const signal = options.signal;
    const runsByPage = new Map<number, readonly TextRun[]>();
    const load = async (pageIndex: number): Promise<readonly TextRun[]> => {
      let runs = runsByPage.get(pageIndex);
      if (!runs) {
        runs = await this.#access.getTextRuns(pageIndex, signal);
        runsByPage.set(pageIndex, runs);
      }
      return runs;
    };
    const pageCount = this.#core.state.pageCount;
    const cached = this.#access.cachedPages();
    const scan = options.pageRange
      ? range(options.pageRange[0], options.pageRange[1], pageCount)
      : range(0, pageCount - 1, pageCount);
    /** Pages that hold runs of a paragraph, found so far. */
    const pagesOf = new Map<string, number[]>();
    const holds = async (
      pageIndex: number,
      paragraphId: string,
    ): Promise<boolean> =>
      (await load(pageIndex)).some((run) => run.paragraphId === paragraphId);
    // A paragraph's pages are contiguous: the first page found to hold
    // it, among the cached pages and then the scanned ones, is grown
    // in both directions until a page without it.
    const locate = async (paragraphId: string): Promise<readonly number[]> => {
      const known = pagesOf.get(paragraphId);
      if (known) return known;
      let seed: number | undefined;
      for (const pageIndex of [...cached, ...scan]) {
        if (await holds(pageIndex, paragraphId)) {
          seed = pageIndex;
          break;
        }
      }
      const pages: number[] = [];
      if (seed !== undefined) {
        pages.push(seed);
        for (
          let pageIndex = seed - 1;
          pageIndex >= 0 && (await holds(pageIndex, paragraphId));
          pageIndex -= 1
        )
          pages.unshift(pageIndex);
        for (
          let pageIndex = seed + 1;
          pageIndex < pageCount && (await holds(pageIndex, paragraphId));
          pageIndex += 1
        )
          pages.push(pageIndex);
      }
      pagesOf.set(paragraphId, pages);
      return pages;
    };
    const fold = (value: string): string =>
      options.caseSensitive ? value : value.toLocaleLowerCase();
    const occurrences = new Map<string, number>();
    const out: TextTarget[] = [];
    for (const target of targets) {
      const elementId = target.elementIds[0];
      const paragraphId = elementId?.startsWith("p:")
        ? elementId.slice(2)
        : undefined;
      const pages = paragraphId ? await locate(paragraphId) : [];
      if (!paragraphId || pages.length === 0) {
        out.push(target);
        continue;
      }
      const runs: { run: TextRun; pageIndex: number }[] = [];
      for (const pageIndex of pages)
        for (const run of await load(pageIndex))
          if (run.paragraphId === paragraphId) runs.push({ run, pageIndex });
      const key = `${paragraphId}\u0000${fold(target.text)}`;
      const occurrence = occurrences.get(key) ?? 0;
      occurrences.set(key, occurrence + 1);
      const rects = matchRects(runs, fold(target.text), occurrence, fold);
      const placed =
        rects.length > 0
          ? rects
          : runs
              .filter((entry) => entry.pageIndex === pages[0])
              .map((entry) => ({
                pageIndex: entry.pageIndex,
                rect: runRect(entry.run),
              }));
      out.push({
        ...target,
        pageIndex: placed[0]?.pageIndex ?? target.pageIndex,
        rects: placed.map((entry) => entry.rect),
      });
    }
    return out;
  }
}

/** Word's default run size, in points, for a paragraph whose style says none. */
const DEFAULT_FONT_SIZE = 10;

/** The id a paragraph's runs carry: its element id without the `p:` prefix. */
function paragraphIdOf(elementId: string): string {
  return elementId.slice(2);
}

/** The paragraph an element id names, when it is one. */
async function paragraphOf(
  engine: EditEngine,
  elementId: string,
  signal: AbortSignal,
): Promise<DocxElement | undefined> {
  if (!elementId.startsWith("p:")) return undefined;
  const elements = (await engine.getElements(
    { kinds: ["paragraph"] },
    signal,
  )) as readonly DocxElement[];
  return elements.find((element) => element.id === elementId);
}

/**
 * The first page that holds a paragraph: the first candidate found to hold
 * it, then the pages before it as long as they hold it too.
 */
async function firstPageHolding(
  candidates: readonly number[],
  holds: (pageIndex: number) => Promise<boolean>,
): Promise<number | undefined> {
  for (const candidate of new Set(candidates)) {
    if (!(await holds(candidate))) continue;
    let first = candidate;
    while (first > 0 && (await holds(first - 1))) first -= 1;
    return first;
  }
  return undefined;
}

function range(first: number, last: number, pageCount: number): number[] {
  const out: number[] = [];
  for (
    let index = Math.max(0, first);
    index <= Math.min(last, pageCount - 1);
    index += 1
  )
    out.push(index);
  return out;
}

/** The `occurrence`-th match of `needle` in the runs' text, cut to run boxes. */
function matchRects(
  runs: readonly { run: TextRun; pageIndex: number }[],
  needle: string,
  occurrence: number,
  fold: (value: string) => string,
): { pageIndex: number; rect: PageRect }[] {
  if (needle.length === 0) return [];
  const text = fold(runs.map((entry) => entry.run.text).join(""));
  let at = -1;
  for (let found = 0; found <= occurrence; found += 1) {
    at = text.indexOf(needle, at + 1);
    if (at < 0) return [];
  }
  const end = at + needle.length;
  const out: { pageIndex: number; rect: PageRect }[] = [];
  let offset = 0;
  for (const { run, pageIndex } of runs) {
    const runStart = offset;
    const runEnd = offset + run.text.length;
    offset = runEnd;
    if (runEnd <= at || runStart >= end || run.text.length === 0) continue;
    const from = Math.max(at, runStart) - runStart;
    const to = Math.min(end, runEnd) - runStart;
    const box = runRect(run);
    const scale = box.width / run.text.length;
    out.push({
      pageIndex,
      rect: {
        x: box.x + from * scale,
        y: box.y,
        width: (to - from) * scale,
        height: box.height,
      },
    });
  }
  return out;
}

function runRect(run: TextRun): PageRect {
  return { x: run.x, y: run.y, width: run.width, height: run.height };
}

function union(a: PageRect, b: PageRect): PageRect {
  const x = Math.min(a.x, b.x);
  const y = Math.min(a.y, b.y);
  return {
    x,
    y,
    width: Math.max(a.x + a.width, b.x + b.width) - x,
    height: Math.max(a.y + a.height, b.y + b.height) - y,
  };
}

function rectContains(rect: PageRect, point: PagePoint): boolean {
  return (
    point.x >= rect.x &&
    point.x <= rect.x + rect.width &&
    point.y >= rect.y &&
    point.y <= rect.y + rect.height
  );
}

function rectsIntersect(a: PageRect, b: PageRect): boolean {
  return (
    a.x < b.x + b.width &&
    b.x < a.x + a.width &&
    a.y < b.y + b.height &&
    b.y < a.y + a.height
  );
}

function fragmentsOf(
  pages: ReadonlyMap<number, PageRect> | undefined,
): ElementFragment[] {
  if (!pages) return [];
  return [...pages.entries()]
    .sort((a, b) => a[0] - b[0])
    .map(([pageIndex, bounds]) => ({ pageIndex, bounds }));
}

function withFragments(
  element: DocxElement,
  fragments: readonly ElementFragment[],
): DocxElement {
  const first = fragments[0];
  return {
    ...element,
    pageIndex: first ? first.pageIndex : -1,
    bounds: first ? first.bounds : { x: 0, y: 0, width: 0, height: 0 },
    fragments,
  };
}

/** Characters a paragraph may hold without the renderer drawing a text run for it. */
const INVISIBLE = /^[\ufffc\s]*$/;

/**
 * A paragraph that draws no run (empty, or holding only pictures) is
 * placed by estimate right after the placed paragraph before it in the
 * same body or cell, else right before the one after it: a picture
 * paragraph takes its largest picture's extent, an empty one a line of
 * its neighbour's height. Nothing is estimated for a paragraph with text,
 * whose runs simply lie on a page not read yet.
 */
function estimateRunless(
  elements: readonly DocxElement[],
  placed: Map<string, DocxElement>,
): void {
  const paragraphs = elements.filter((element) => element.kind === "paragraph");
  const sizeOf = (
    paragraph: DocxElement,
  ): { width: number; height: number } | undefined => {
    const pictures = elements.filter(
      (candidate) =>
        candidate.kind === "image" &&
        candidate.parentId === paragraph.id &&
        candidate.imageSize,
    );
    if (pictures.length === 0) return undefined;
    return {
      width: Math.max(...pictures.map((picture) => picture.imageSize!.width)),
      height: Math.max(...pictures.map((picture) => picture.imageSize!.height)),
    };
  };
  for (const [index, paragraph] of paragraphs.entries()) {
    const current = placed.get(paragraph.id)!;
    if (current.pageIndex >= 0 || !INVISIBLE.test(paragraph.text ?? ""))
      continue;
    const sibling = (step: number): ElementFragment | undefined => {
      for (
        let at = index + step;
        at >= 0 && at < paragraphs.length;
        at += step
      ) {
        const other = paragraphs[at]!;
        if (other.parentId !== paragraph.parentId) continue;
        const fragments = placed.get(other.id)?.fragments ?? [];
        const fragment = step < 0 ? fragments.at(-1) : fragments[0];
        if (fragment) return fragment;
        if (!INVISIBLE.test(other.text ?? "")) return undefined;
      }
      return undefined;
    };
    const before = sibling(-1);
    const after = before ? undefined : sibling(1);
    const anchor = before ?? after;
    if (!anchor) continue;
    const size = sizeOf(paragraph) ?? {
      width: anchor.bounds.width,
      height: anchor.bounds.height,
    };
    const bounds = before
      ? {
          x: before.bounds.x,
          y: before.bounds.y + before.bounds.height,
          ...size,
        }
      : {
          x: after!.bounds.x,
          y: Math.max(0, after!.bounds.y - size.height),
          ...size,
        };
    placed.set(
      paragraph.id,
      withFragments(paragraph, [{ pageIndex: anchor.pageIndex, bounds }]),
    );
  }
}

/** Joins the engine's elements with the runs' placement. */
export function placeElements(
  elements: readonly DocxElement[],
  placement: Placement,
): DocxElement[] {
  const placed = new Map<string, DocxElement>();
  // Paragraphs first; tables and inline objects derive from them.
  for (const element of elements)
    if (element.kind === "paragraph")
      placed.set(
        element.id,
        withFragments(element, fragmentsOf(placement.get(element.id.slice(2)))),
      );
  estimateRunless(elements, placed);
  for (const element of elements) {
    if (element.kind === "table") {
      const pages = new Map<number, PageRect>();
      for (const candidate of elements) {
        if (candidate.kind !== "paragraph" || candidate.parentId !== element.id)
          continue;
        for (const fragment of placed.get(candidate.id)?.fragments ?? []) {
          const current = pages.get(fragment.pageIndex);
          pages.set(
            fragment.pageIndex,
            current ? union(current, fragment.bounds) : fragment.bounds,
          );
        }
      }
      placed.set(element.id, withFragments(element, fragmentsOf(pages)));
    } else if (element.kind !== "paragraph") {
      const parent = element.parentId
        ? placed.get(element.parentId)
        : undefined;
      placed.set(element.id, withFragments(element, parent?.fragments ?? []));
    }
  }
  return elements.map((element) => placed.get(element.id) ?? element);
}

function matchesQuery(element: DocxElement, query: ElementQuery): boolean {
  if (query.kinds && !query.kinds.includes(element.kind)) return false;
  if (query.pageIndex === undefined) return true;
  const fragment = element.fragments?.find(
    (candidate) => candidate.pageIndex === query.pageIndex,
  );
  if (!fragment) return false;
  return !query.intersects || rectsIntersect(fragment.bounds, query.intersects);
}

function docxReads(engine: EditEngine): DocxEngineReads {
  const candidate = engine as Partial<DocxEngineReads>;
  if (
    typeof candidate.revisions !== "function" ||
    typeof candidate.textStyle !== "function" ||
    typeof candidate.textColors !== "function"
  )
    throw new ViewerError(
      "edit-unsupported",
      "The engine does not provide the DOCX reads",
      { details: { format: "docx", reason: "no-reads" } },
    );
  return candidate as DocxEngineReads;
}

/** Validate at the public boundary before copying: a cyclic value must not recurse in the copier. */
function frozenPreviewFields(
  fields: DocxFields<DocxReplaceTextOperation>,
): DocxFields<DocxReplaceTextOperation> {
  const issues = checkOperations(
    [{ ...fields, op: "replaceText" }],
    docxOperationSchemas,
  );
  if (issues.length > 0) throw invalidOperationError(issues);
  const draft = freezeOperations([fields])[0];
  if (!draft)
    throw new ViewerError("invalid-operation", "Missing draft fields");
  return draft;
}
