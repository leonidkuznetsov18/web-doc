import { ViewerError } from "../../errors.js";
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
import type { TextRun } from "../../contracts.js";
import type {
  EditEngine,
  EditSessionAccess,
  EditSessionCore,
} from "../engine.js";
import type {
  ApplyOptions,
  AssetOptions,
  EditFindOptions,
  EditOperation,
  EditReceipt,
  EditState,
  ElementQuery,
  HistoryOptions,
  OperationSchemaSet,
  PagePoint,
  ReadItem,
  ReadOptions,
  ReadResult,
  SavedDocument,
  TextRange,
  TextTarget,
} from "../types.js";
import type { PptxEngineReads } from "./engine.js";
import { spanOnTarget } from "../range-style.js";
import { isDrawn } from "./visibility.js";
import type {
  PptxDeleteElementOperation,
  PptxDeleteSlideOperation,
  PptxDuplicateSlideOperation,
  PptxEditSession,
  PptxElement,
  PptxFields,
  PptxInsertImageOperation,
  PptxInsertSlideOperation,
  PptxInsertTableOperation,
  PptxInsertTextBoxOperation,
  PptxLayoutInfo,
  PptxMoveElementOperation,
  PptxMoveSlideOperation,
  PptxOperation,
  PptxReplaceTextOperation,
  PptxResizeElementOperation,
  PptxSaveOptions,
  PptxSetShapeStyleOperation,
  PptxSetTableCellOperation,
  PptxSetTextStyleOperation,
  PptxSlideInfo,
  PptxTextStyle,
} from "./types.js";

/**
 * The PPTX session: the core session narrowed to PPTX operations and
 * elements, plus the slide and layout reads. Each typed method is `apply()`
 * with a single operation.
 */
export class PptxSession implements PptxEditSession {
  readonly format = "pptx" as const;
  readonly #core: EditSessionCore;
  #tools: ToolSet | undefined;

  /** The shown deck's text runs; absent in a headless session, which hits frames only. */
  readonly #access: EditSessionAccess | undefined;

  constructor(core: EditSessionCore, access?: EditSessionAccess) {
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
    operations: readonly PptxOperation[],
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

  save(options?: PptxSaveOptions): Promise<SavedDocument> {
    return this.#core.save(options);
  }

  markSaved(stateToken: string): void {
    this.#core.markSaved(stateToken);
  }

  addAsset(data: Uint8Array, options?: AssetOptions): Promise<string> {
    return this.#core.addAsset(data, options);
  }

  getElements(
    query?: ElementQuery,
    options?: ReadOptions,
  ): Promise<ReadResult<PptxElement>> {
    return this.#core.getElements(query, options) as Promise<
      ReadResult<PptxElement>
    >;
  }

  getElement(
    id: string,
    options?: ReadOptions,
  ): Promise<ReadItem<PptxElement>> {
    return this.#core.getElement(id, options) as Promise<ReadItem<PptxElement>>;
  }

  /**
   * The drawn elements under a point, top-most first: those whose frame
   * holds it, and a shape whose text is painted there past its frame, as
   * text wrapped below a short box is. Hidden shapes, and shapes in hidden
   * groups, are never hit; `getElements` and `getElement` still list them.
   */
  async elementsAt(
    pageIndex: number,
    point: PagePoint,
    options?: ReadOptions,
  ): Promise<ReadResult<PptxElement>> {
    const framed = (await this.#core.elementsAt(
      pageIndex,
      point,
      options,
    )) as ReadResult<PptxElement>;
    const access = this.#access;
    if (!access) return framed;
    const runs = (await access.getTextRuns(pageIndex, options?.signal)).filter(
      (run) => runHolds(run, point),
    );
    if (runs.length === 0) return framed;
    const all = (await this.getElements({ pageIndex }, options)).items;
    const byId = new Map(all.map((element) => [element.id, element]));
    const visibleText = all.filter(
      (element) => element.text !== undefined && isDrawn(element, byId),
    );
    // The edit ID is sld<part-number>:<cNvPr-id>[#duplicate-occurrence].
    // Index every element, including hidden ones: a malformed duplicate ID
    // cannot safely be bound to whichever visible occurrence comes first.
    const owners = new Map<string, PptxElement | undefined>();
    for (const element of all) {
      const sourceId = /:(\d+)(?:#\d+)?$/.exec(element.id)?.[1];
      if (sourceId !== undefined)
        owners.set(sourceId, owners.has(sourceId) ? undefined : element);
    }
    const paintedIds = new Set<string>();
    for (const run of runs) {
      if (run.shapeId !== undefined || run.shapeSource !== undefined) {
        // Layout/master IDs may equal a slide ID. Missing, stale or ambiguous
        // native ownership must not fall back to a coincident origin.
        if (
          run.shapeSource !== "slide" ||
          run.shapeId === undefined ||
          !/^\d+$/.test(run.shapeId) ||
          !Number.isSafeInteger(Number(run.shapeId))
        )
          continue;
        const owner = owners.get(String(Number(run.shapeId)));
        if (owner && visibleText.includes(owner)) paintedIds.add(owner.id);
        continue;
      }
      // Legacy providers report only geometry. A shared origin cannot prove
      // which shape painted the text, so keep ordinary frame hits in that case.
      const origin = run.shapeOrigin;
      const candidates = origin
        ? visibleText.filter((element) => originOf(element, origin))
        : [];
      const candidate = candidates[0];
      if (candidate && candidates.length === 1) paintedIds.add(candidate.id);
    }
    const painted = visibleText.filter(
      (element) =>
        paintedIds.has(element.id) &&
        !framed.items.some((hit) => hit.id === element.id),
    );
    if (painted.length === 0) return framed;
    // Elements are listed back to front; whatever is drawn after a shape
    // covers the text it paints past its frame.
    const order = new Map(all.map((element, index) => [element.id, index]));
    return Object.freeze({
      ...framed,
      items: Object.freeze(
        [...painted, ...framed.items].sort(
          (a, b) => (order.get(b.id) ?? -1) - (order.get(a.id) ?? -1),
        ),
      ),
    });
  }

  findText(
    query: string,
    options?: EditFindOptions,
  ): Promise<ReadResult<TextTarget>> {
    return this.#core.findText(query, options);
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

  getSlides(options?: ReadOptions): Promise<ReadResult<PptxSlideInfo>> {
    return this.#core.readItems(options, (engine, signal) =>
      pptxReads(engine).slides(signal),
    );
  }

  getLayouts(options?: ReadOptions): Promise<ReadResult<PptxLayoutInfo>> {
    return this.#core.readItems(options, (engine, signal) =>
      pptxReads(engine).layouts(signal),
    );
  }

  async getTextStyle(
    fields: { readonly target: string; readonly range?: TextRange },
    options?: ReadOptions,
  ): Promise<ReadItem<Partial<PptxTextStyle>>> {
    const span = spanOnTarget(fields.target, fields.range);
    return this.#core.readItem(options, (engine, signal) =>
      pptxReads(engine).textStyle(fields.target, span, signal),
    );
  }

  replaceText(
    fields: PptxFields<PptxReplaceTextOperation>,
    options?: ApplyOptions,
  ): Promise<EditReceipt> {
    return this.apply([{ op: "replaceText", ...fields }], options);
  }

  setTextStyle(
    fields: PptxFields<PptxSetTextStyleOperation>,
    options?: ApplyOptions,
  ): Promise<EditReceipt> {
    return this.apply([{ op: "setTextStyle", ...fields }], options);
  }

  setShapeStyle(
    fields: PptxFields<PptxSetShapeStyleOperation>,
    options?: ApplyOptions,
  ): Promise<EditReceipt> {
    return this.apply([{ op: "setShapeStyle", ...fields }], options);
  }

  moveElement(
    fields: PptxFields<PptxMoveElementOperation>,
    options?: ApplyOptions,
  ): Promise<EditReceipt> {
    return this.apply([{ op: "moveElement", ...fields }], options);
  }

  resizeElement(
    fields: PptxFields<PptxResizeElementOperation>,
    options?: ApplyOptions,
  ): Promise<EditReceipt> {
    return this.apply([{ op: "resizeElement", ...fields }], options);
  }

  deleteElement(
    fields: PptxFields<PptxDeleteElementOperation>,
    options?: ApplyOptions,
  ): Promise<EditReceipt> {
    return this.apply([{ op: "deleteElement", ...fields }], options);
  }

  insertTextBox(
    fields: PptxFields<PptxInsertTextBoxOperation>,
    options?: ApplyOptions,
  ): Promise<EditReceipt> {
    return this.apply([{ op: "insertTextBox", ...fields }], options);
  }

  insertImage(
    fields: PptxFields<PptxInsertImageOperation>,
    options?: ApplyOptions,
  ): Promise<EditReceipt> {
    return this.apply([{ op: "insertImage", ...fields }], options);
  }

  insertTable(
    fields: PptxFields<PptxInsertTableOperation>,
    options?: ApplyOptions,
  ): Promise<EditReceipt> {
    return this.apply([{ op: "insertTable", ...fields }], options);
  }

  setTableCell(
    fields: PptxFields<PptxSetTableCellOperation>,
    options?: ApplyOptions,
  ): Promise<EditReceipt> {
    return this.apply([{ op: "setTableCell", ...fields }], options);
  }

  insertSlide(
    fields: PptxFields<PptxInsertSlideOperation>,
    options?: ApplyOptions,
  ): Promise<EditReceipt> {
    return this.apply([{ op: "insertSlide", ...fields }], options);
  }

  duplicateSlide(
    fields: PptxFields<PptxDuplicateSlideOperation>,
    options?: ApplyOptions,
  ): Promise<EditReceipt> {
    return this.apply([{ op: "duplicateSlide", ...fields }], options);
  }

  deleteSlide(
    fields: PptxFields<PptxDeleteSlideOperation>,
    options?: ApplyOptions,
  ): Promise<EditReceipt> {
    return this.apply([{ op: "deleteSlide", ...fields }], options);
  }

  moveSlide(
    fields: PptxFields<PptxMoveSlideOperation>,
    options?: ApplyOptions,
  ): Promise<EditReceipt> {
    return this.apply([{ op: "moveSlide", ...fields }], options);
  }
}

/** The engine's PPTX reads; an engine without them cannot serve this session. */
function pptxReads(engine: EditEngine): PptxEngineReads {
  const candidate = engine as Partial<PptxEngineReads>;
  if (
    typeof candidate.slides !== "function" ||
    typeof candidate.layouts !== "function" ||
    typeof candidate.textStyle !== "function"
  )
    throw new ViewerError(
      "edit-unsupported",
      "The engine does not provide the PPTX reads",
      { details: { format: "pptx", reason: "no-reads" } },
    );
  return candidate as PptxEngineReads;
}

/** Whether a run's box holds a point. */
function runHolds(run: TextRun, point: PagePoint): boolean {
  return (
    point.x >= run.x &&
    point.x <= run.x + run.width &&
    point.y >= run.y &&
    point.y <= run.y + run.height
  );
}

/** Whether a shape's frame starts where a run says its shape does, to half a pixel. */
function originOf(
  element: PptxElement,
  origin: { readonly x: number; readonly y: number },
): boolean {
  const frame = element.frame ?? element.bounds;
  return (
    Math.abs(frame.x - origin.x) <= 0.5 && Math.abs(frame.y - origin.y) <= 0.5
  );
}
