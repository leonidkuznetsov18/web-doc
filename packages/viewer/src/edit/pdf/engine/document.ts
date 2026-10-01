import { ViewerError } from "../../../errors.js";
import type { EngineBatch, EngineChange, RestoreTarget } from "../../engine.js";
import { invalidOperationError, parseReference } from "../../operations.js";
import type {
  EditFindOptions,
  EditOperation,
  ElementQuery,
  OperationIssue,
  PagePoint,
  TextTarget,
} from "../../types.js";
import type { PdfElement, PdfOperation } from "../types.js";
import {
  readMark,
  scanPage,
  type MarkParams,
  type ObjectRecord,
} from "./elements.js";
import { FontLibrary, TextMeasurer } from "./fonts.js";
import { fontRequestsOf } from "./text-box.js";
import {
  issueCollector,
  type ElementLocation,
  type OperationContext,
  type OperationHandler,
} from "./operations.js";
import { insertTextBox } from "./text-box.js";
import { replaceText, setTextStyle } from "./existing-text.js";
import { deleteElement, moveElement, resizeElement } from "./transform.js";
import { deletePage, insertPage, movePage, rotatePage } from "./pages.js";
import { insertShape, setShapeStyle } from "./shapes.js";
import { ImageCache, insertImage } from "./images.js";
import { insertTable, setTableCell, tableSpecOf } from "./tables.js";
import { defaultResourceLimits } from "../../../limits.js";
import type { ResourceLimits } from "../../../contracts.js";
import {
  displayedSize,
  rectContains,
  rectsIntersect,
  roundRect,
  unionRects,
  userRectToPage,
  type PageGeometry,
} from "./geometry.js";
import type { Pdfium, PdfiumDocument } from "./pdfium.js";

/** One page of the working copy and what the model knows about it. */
interface PageRecord {
  /** Stable key: the page's position in the original, or the operation that created it. */
  readonly key: string;
  geometry?: PageGeometry;
  /** Objects in drawing order, assigned once per page and kept across edits. */
  objects?: ObjectRecord[];
  /** Elements derived from `objects`; dropped whenever the page changes. */
  elements?: readonly PdfElement[];
}

/**
 * The working copy of one PDF: a PDFium document rebuilt from the original
 * bytes plus the applied batches, with a stable id for every page and object.
 * Everything a query answers comes from PDFium; the model only remembers
 * identities and caches.
 */
/** Batches arrive as plain JSON; unknown operations are reported, not typed away. */
type PdfOrUnknownOperation = PdfOperation | EditOperation;

/** Operations whose first created id a `"$<n>"` reference can name. */
const CREATING_OPERATIONS = new Set<string>([
  "insertTextBox",
  "insertShape",
  "insertImage",
  "insertTable",
]);

function referenceOf(operation: { readonly op: string }): number | undefined {
  const target = (operation as { readonly target?: unknown }).target;
  return typeof target === "string" ? parseReference(target) : undefined;
}

const handlers: Readonly<Record<PdfOperation["op"], OperationHandler>> = {
  insertTextBox: insertTextBox as OperationHandler,
  replaceText: replaceText as OperationHandler,
  setTextStyle: setTextStyle as OperationHandler,
  resizeElement: resizeElement as OperationHandler,
  moveElement: moveElement as OperationHandler,
  deleteElement: deleteElement as OperationHandler,
  insertPage: insertPage as OperationHandler,
  deletePage: deletePage as OperationHandler,
  movePage: movePage as OperationHandler,
  rotatePage: rotatePage as OperationHandler,
  insertShape: insertShape as OperationHandler,
  setShapeStyle: setShapeStyle as OperationHandler,
  insertImage: insertImage as OperationHandler,
  insertTable: insertTable as OperationHandler,
  setTableCell: setTableCell as OperationHandler,
};

export class PdfEditDocument {
  readonly #pdfium: Pdfium;
  readonly #original: Uint8Array;
  readonly #fonts: FontLibrary;
  readonly images = new ImageCache();
  readonly #limits: ResourceLimits;
  /** Signature fields in the original; an edit leaves them uncovering the new revision. */
  readonly #signatures: number;
  /** What the working copy was opened from: the original, or a checkpoint. */
  #base: Uint8Array;
  #document: PdfiumDocument;
  #measurer: TextMeasurer;
  #pages: PageRecord[];
  #batches = 0;

  constructor(
    pdfium: Pdfium,
    original: Uint8Array,
    fonts: FontLibrary = new FontLibrary(async () => {
      throw new Error("No font source is configured");
    }),
    limits: ResourceLimits = defaultResourceLimits,
  ) {
    this.#pdfium = pdfium;
    this.#original = original;
    this.#base = original;
    this.#fonts = fonts;
    this.#limits = limits;
    this.#document = pdfium.openDocument(original);
    this.#measurer = new TextMeasurer(pdfium, this.#document.handle);
    this.#pages = this.#originalPages();
    this.#signatures = pdfium.lib.FPDF_GetSignatureCount(this.#document.handle);
  }

  /** Signature fields in the document. */
  get signatureCount(): number {
    return this.#signatures;
  }

  /**
   * The families and texts a batch will draw, so the fonts can be fetched
   * before the synchronous validation and application run.
   */
  fontRequests(
    operations: readonly PdfOrUnknownOperation[],
  ): { readonly family: string; readonly text: string }[] {
    return fontRequestsOf(
      operations,
      (id) => this.#locate(id)?.record.mark,
      (id) => this.getElement(id),
    );
  }

  get pageCount(): number {
    return this.#pages.length;
  }

  get pdfium(): Pdfium {
    return this.#pdfium;
  }

  /**
   * Checks a batch against the current document. Operations are checked one
   * after another against the state before the batch, which is exact for
   * everything but a page count another operation of the batch changes.
   */
  validate(operations: readonly PdfOrUnknownOperation[]): OperationIssue[] {
    const issues: OperationIssue[] = [];
    const context = this.#context(0, 0);
    operations.forEach((operation, operationIndex) => {
      const handler = handlers[operation.op as PdfOperation["op"]];
      if (!handler) {
        issues.push({
          operationIndex,
          path: "/op",
          code: "unknown-operation",
          message: `Unknown pdf operation ${operation.op}`,
        });
        return;
      }
      const issue = issueCollector(operationIndex, issues);
      const reference = referenceOf(operation);
      if (reference !== undefined) {
        // The element does not exist yet; the rest is checked when applying.
        const creator = operations[reference];
        if (
          reference >= operationIndex ||
          !creator ||
          !CREATING_OPERATIONS.has(creator.op)
        )
          issue(
            "/target",
            "unknown-target",
            `"$${reference}" must name an earlier operation that creates an element`,
          );
        return;
      }
      handler.validate(operation as PdfOperation, context, issue);
    });
    return issues;
  }

  /**
   * Applies a batch. Plain operation arrays, as the unit tests pass them, get
   * the next batch number as their state id.
   */
  apply(input: readonly PdfOrUnknownOperation[] | EngineBatch): EngineChange {
    const batch: EngineBatch = Array.isArray(input)
      ? { stateId: this.#batches + 1, operations: input }
      : (input as EngineBatch);
    const { stateId, operations } = batch;
    const createdIds: string[] = [];
    const createdByOperation: string[][] = [];
    const removedIds: string[] = [];
    const changedPages = new Set<number>();
    const warnings: EngineChange["warnings"][number][] = [];
    // The first change of a signed file is the point where the signatures
    // stop covering what is shown; the incremental save keeps them valid for
    // the original revision.
    if (this.#batches === 0 && this.#signatures > 0)
      warnings.push({
        code: "fidelity-degraded",
        message: `The document carries ${this.#signatures} digital signature${this.#signatures === 1 ? "" : "s"} that will not cover the edited revision`,
        details: { signatures: this.#signatures },
      });
    operations.forEach((raw, operationIndex) => {
      const handler = handlers[raw.op as PdfOperation["op"]];
      if (!handler)
        throw new ViewerError("internal", `Unknown pdf operation ${raw.op}`);
      const context = this.#context(stateId, operationIndex);
      const operation = this.#resolveReference(
        raw as PdfOperation,
        operationIndex,
        createdByOperation,
        handler,
        context,
      );
      const result = handler.apply(operation, context);
      createdIds.push(...result.createdIds);
      createdByOperation[operationIndex] = [...result.createdIds];
      removedIds.push(...(result.removedIds ?? []));
      for (const pageIndex of result.changedPages) changedPages.add(pageIndex);
      warnings.push(...result.warnings);
    });
    this.#batches += 1;
    return {
      createdIds,
      removedIds,
      changedPages: [...changedPages].sort((a, b) => a - b),
      pageCount: this.pageCount,
      warnings,
    };
  }

  /**
   * Turns a `"$<n>"` target into the id operation `n` created, then runs the
   * handler's own checks on the resolved operation: a reference that lands
   * on an element the operation cannot act on fails the batch the same way
   * validation would have.
   */
  #resolveReference(
    operation: PdfOperation,
    operationIndex: number,
    createdByOperation: readonly (readonly string[])[],
    handler: OperationHandler,
    context: OperationContext,
  ): PdfOperation {
    const reference = referenceOf(operation);
    if (reference === undefined) return operation;
    const issues: OperationIssue[] = [];
    const issue = issueCollector(operationIndex, issues);
    const id = createdByOperation[reference]?.[0];
    if (id === undefined) {
      issue(
        "/target",
        "unknown-target",
        `Operation ${reference} created no element for "$${reference}"`,
      );
      throw invalidOperationError(issues);
    }
    const resolved = { ...operation, target: id } as PdfOperation;
    handler.validate(resolved, context, issue);
    if (issues.length > 0) throw invalidOperationError(issues);
    return resolved;
  }

  /** The base bytes while nothing changed, else an incremental update. */
  materialize(): Uint8Array {
    return this.#batches === 0
      ? this.#base.slice()
      : this.#document.save("incremental");
  }

  /**
   * Rebuilds a state from its base (the original, or a checkpoint) and the
   * batches after it. Plain arrays of batches, as the unit tests pass them,
   * get state ids 1, 2, 3…
   */
  restore(
    input: readonly (readonly PdfOrUnknownOperation[])[] | RestoreTarget,
  ): void {
    const target: RestoreTarget = Array.isArray(input)
      ? {
          batches: (input as readonly (readonly PdfOrUnknownOperation[])[]).map(
            (operations, index) => ({ stateId: index + 1, operations }),
          ),
        }
      : (input as RestoreTarget);
    this.#fonts.release(this.#pdfium, this.#document.handle);
    this.#document.close();
    this.#base = target.base ?? this.#original;
    this.#document = this.#pdfium.openDocument(this.#base);
    this.#measurer = new TextMeasurer(this.#pdfium, this.#document.handle);
    this.#pages = this.#originalPages();
    this.#batches = 0;
    for (const batch of target.batches) this.apply(batch);
  }

  getElements(query: ElementQuery): PdfElement[] {
    const pages =
      query.pageIndex === undefined
        ? this.#pages.map((_, index) => index)
        : [query.pageIndex];
    const result: PdfElement[] = [];
    for (const pageIndex of pages) {
      if (pageIndex < 0 || pageIndex >= this.#pages.length) continue;
      for (const element of this.#elementsOf(pageIndex)) {
        if (query.kinds && !query.kinds.includes(element.kind)) continue;
        if (
          query.intersects &&
          !rectsIntersect(element.bounds, query.intersects)
        )
          continue;
        result.push(element);
      }
    }
    return result;
  }

  getElement(id: string): PdfElement | undefined {
    const pageIndex = this.#pageIndexOf(id);
    if (pageIndex === undefined) return undefined;
    return this.#elementsOf(pageIndex).find((element) => element.id === id);
  }

  /** Elements under a point, top-most (drawn last) first. */
  elementsAt(pageIndex: number, point: PagePoint): PdfElement[] {
    if (pageIndex < 0 || pageIndex >= this.#pages.length) return [];
    return this.#elementsOf(pageIndex)
      .filter((element) => rectContains(element.bounds, point))
      .reverse();
  }

  findText(query: string, options: EditFindOptions): TextTarget[] {
    if (!query) return [];
    const [first, last] = options.pageRange ?? [0, this.#pages.length - 1];
    const limit = options.maxResults ?? Number.POSITIVE_INFINITY;
    const targets: TextTarget[] = [];
    for (
      let pageIndex = Math.max(0, first);
      pageIndex <= Math.min(last, this.#pages.length - 1) &&
      targets.length < limit;
      pageIndex += 1
    )
      this.#withPage(pageIndex, (page, textPage, geometry) => {
        const records = this.#objectsOf(pageIndex, page);
        const { byObject } = scanPage(
          this.#pdfium,
          page,
          textPage,
          pageIndex,
          geometry,
          records,
        );
        for (const match of this.#matches(textPage, query, options)) {
          if (targets.length >= limit) break;
          targets.push(
            this.#target(textPage, pageIndex, geometry, match, byObject),
          );
        }
      });
    return targets;
  }

  dispose(): void {
    this.#fonts.release(this.#pdfium, this.#document.handle);
    this.#document.close();
  }

  #context(stateId: number, operationIndex: number): OperationContext {
    let created = 0;
    return {
      pdfium: this.#pdfium,
      document: this.#document.handle,
      measurer: this.#measurer,
      fonts: this.#fonts,
      images: this.images,
      limits: this.#limits,
      pageCount: this.pageCount,
      geometry: (pageIndex) => {
        const record = this.#pages[pageIndex];
        if (!record) throw new ViewerError("internal", `No page ${pageIndex}`);
        return (
          record.geometry ??
          this.#withPage(pageIndex, (page) => this.#geometryOf(pageIndex, page))
        );
      },
      // Ids name the state the batch leads to, the operation and the item,
      // so replaying the history reproduces them and an undone id is never
      // handed out again.
      newId: (pageIndex, suffix = "") =>
        `${this.#pages[pageIndex]!.key}:n${stateId}.${operationIndex}.${created++}${suffix}`,
      withPage: (pageIndex, use) => this.#writePage(pageIndex, use),
      appendObjects: (pageIndex, records) => {
        const page = this.#pages[pageIndex]!;
        page.objects = [...(page.objects ?? []), ...records];
        delete page.elements;
      },
      locate: (id) => this.#locate(id),
      element: (id) => this.getElement(id),
      pageElementIds: (pageIndex) =>
        this.#elementsOf(pageIndex).map((element) => element.id),
      pageSize: (pageIndex) =>
        displayedSize(
          this.#withPage(pageIndex, (page) =>
            this.#geometryOf(pageIndex, page),
          ),
        ),
      insertPageRecord: (index) => {
        const key = `q${stateId}.${operationIndex}`;
        this.#pages.splice(index, 0, { key });
        this.#forgetElements();
        return key;
      },
      removePageRecord: (index) => {
        this.#pages.splice(index, 1);
        this.#forgetElements();
      },
      movePageRecord: (from, to) => {
        const [record] = this.#pages.splice(from, 1);
        this.#pages.splice(to, 0, record!);
        this.#forgetElements();
      },
      invalidatePage: (index) => {
        const page = this.#pages[index];
        if (!page) return;
        delete page.geometry;
        delete page.elements;
      },
      spliceObjects: (pageIndex, start, count, records) => {
        const page = this.#pages[pageIndex]!;
        const objects = [...(page.objects ?? [])];
        objects.splice(start, count, ...records);
        page.objects = objects;
        delete page.elements;
      },
    };
  }

  /** Cached elements carry page indexes, which a structure change makes stale. */
  #forgetElements(): void {
    for (const page of this.#pages) delete page.elements;
  }

  #locate(id: string): ElementLocation | undefined {
    const pageIndex = this.#pageIndexOf(id);
    if (pageIndex === undefined) return undefined;
    const records = this.#withPage(pageIndex, (page) =>
      this.#objectsOf(pageIndex, page),
    );
    const indexes: number[] = [];
    records.forEach((record, index) => {
      if (record.id === id) indexes.push(index);
    });
    const record = records[indexes[0] ?? -1];
    return record ? { pageIndex, indexes, record } : undefined;
  }

  /** Loads a page, lets `use` change it, regenerates its content stream. */
  #writePage<T>(pageIndex: number, use: (page: number) => T): T {
    const { lib } = this.#pdfium;
    const page = lib.FPDF_LoadPage(this.#document.handle, pageIndex);
    if (!page)
      throw new ViewerError("render-failed", "PDFium could not load the page", {
        details: { pageIndex },
      });
    try {
      // Objects are assigned before the change so appended ones follow them.
      this.#objectsOf(pageIndex, page);
      const result = use(page);
      if (!lib.FPDFPage_GenerateContent(page))
        throw new ViewerError(
          "edit-failed",
          "PDFium could not rewrite the page",
          {
            details: { stage: "apply", pageIndex },
          },
        );
      delete this.#pages[pageIndex]!.elements;
      return result;
    } finally {
      lib.FPDF_ClosePage(page);
    }
  }

  #originalPages(): PageRecord[] {
    const count = this.#pdfium.lib.FPDF_GetPageCount(this.#document.handle);
    return Array.from({ length: count }, (_, index) => ({ key: `p${index}` }));
  }

  #pageIndexOf(id: string): number | undefined {
    const key = id.split(":")[0];
    const index = this.#pages.findIndex((page) => page.key === key);
    return index < 0 ? undefined : index;
  }

  #elementsOf(pageIndex: number): readonly PdfElement[] {
    const record = this.#pages[pageIndex]!;
    if (record.elements) return record.elements;
    return this.#withPage(pageIndex, (page, textPage, geometry) => {
      const { elements } = scanPage(
        this.#pdfium,
        page,
        textPage,
        pageIndex,
        geometry,
        this.#objectsOf(pageIndex, page),
      );
      record.elements = Object.freeze(elements);
      return record.elements;
    });
  }

  /** Object records of a page, assigned on first sight; `page` must be loaded. */
  #objectsOf(pageIndex: number, page: number): ObjectRecord[] {
    const record = this.#pages[pageIndex]!;
    if (record.objects) return record.objects;
    const { lib } = this.#pdfium;
    const count = lib.FPDFPage_CountObjects(page);
    const objects: ObjectRecord[] = [];
    for (let index = 0; index < count; index += 1) {
      const object = lib.FPDFPage_GetObject(page, index);
      const mark = readMark(this.#pdfium, object);
      objects.push({
        id: mark ? mark.id : `${record.key}:o${index}`,
        type: lib.FPDFPageObj_GetType(object),
        ...(mark ? { mark: this.#ownMark(mark, record.key) } : {}),
      });
    }
    // A table's inputs sit on its first object; members without that head
    // are plain objects again.
    const heads = new Map<string, boolean>();
    for (const object of objects)
      if (object.mark && !heads.has(object.id))
        heads.set(object.id, tableSpecOf(object.mark) !== undefined);
    record.objects = objects.map((object, index) =>
      object.mark?.kind === "table" && !heads.get(object.id)
        ? { id: `${record.key}:o${index}`, type: object.type }
        : object,
    );
    return record.objects;
  }

  /** Marks from other sessions keep their id only if it cannot collide with ours. */
  #ownMark(mark: MarkParams, pageKey: string): MarkParams {
    return mark.id.startsWith(`${pageKey}:`)
      ? mark
      : { ...mark, id: `${pageKey}:${mark.id}` };
  }

  #geometryOf(pageIndex: number, page: number): PageGeometry {
    const record = this.#pages[pageIndex]!;
    if (record.geometry) return record.geometry;
    const { lib } = this.#pdfium;
    const box = this.#pdfium.readNumbers(4, "float", ([pointer]) =>
      lib.FPDF_GetPageBoundingBox(page, pointer!),
    );
    const [left, top, right, bottom] = box ?? [
      0,
      lib.FPDF_GetPageHeightF(page),
      lib.FPDF_GetPageWidthF(page),
      0,
    ];
    record.geometry = {
      box: { left: left!, bottom: bottom!, right: right!, top: top! },
      rotation: lib.FPDFPage_GetRotation(page),
    };
    return record.geometry;
  }

  #withPage<T>(
    pageIndex: number,
    use: (page: number, textPage: number, geometry: PageGeometry) => T,
  ): T {
    const { lib } = this.#pdfium;
    const page = lib.FPDF_LoadPage(this.#document.handle, pageIndex);
    if (!page)
      throw new ViewerError("render-failed", "PDFium could not load the page", {
        details: { pageIndex },
      });
    try {
      const textPage = lib.FPDFText_LoadPage(page);
      try {
        return use(page, textPage, this.#geometryOf(pageIndex, page));
      } finally {
        lib.FPDFText_ClosePage(textPage);
      }
    } finally {
      lib.FPDF_ClosePage(page);
    }
  }

  *#matches(
    textPage: number,
    query: string,
    options: EditFindOptions,
  ): Generator<{ readonly start: number; readonly end: number }> {
    const { lib } = this.#pdfium;
    const count = lib.FPDFText_CountChars(textPage);
    if (count === 0) return;
    const bytes = (count + 1) * 2;
    const buffer = this.#pdfium.malloc(bytes);
    let text: string;
    try {
      lib.FPDFText_GetText(textPage, 0, count, buffer);
      text = this.#pdfium.readWideStringAt(buffer, bytes);
    } finally {
      this.#pdfium.free(buffer);
    }
    const haystack = options.caseSensitive ? text : text.toLowerCase();
    const needle = options.caseSensitive ? query : query.toLowerCase();
    let from = 0;
    while (from <= haystack.length - needle.length) {
      const at = haystack.indexOf(needle, from);
      if (at < 0) return;
      // Text indexes count UTF-16 units; PDFium's char indexes do not.
      const start = lib.FPDFText_GetCharIndexFromTextIndex(textPage, at);
      const end = lib.FPDFText_GetCharIndexFromTextIndex(
        textPage,
        at + needle.length - 1,
      );
      if (start >= 0 && end >= start) yield { start, end: end + 1 };
      from = at + Math.max(1, needle.length);
    }
  }

  #target(
    textPage: number,
    pageIndex: number,
    geometry: PageGeometry,
    match: { readonly start: number; readonly end: number },
    byObject: ReadonlyMap<number, string>,
  ): TextTarget {
    const { lib } = this.#pdfium;
    const rects: {
      left: number;
      bottom: number;
      right: number;
      top: number;
    }[] = [];
    const ids: string[] = [];
    let text = "";
    for (let index = match.start; index < match.end; index += 1) {
      const box = this.#pdfium.readNumbers(4, "double", ([l, r, b, t]) =>
        lib.FPDFText_GetCharBox(textPage, index, l!, r!, b!, t!),
      );
      const unit = lib.FPDFText_GetUnicode(textPage, index);
      text += String.fromCodePoint(unit);
      const id = byObject.get(lib.FPDFText_GetTextObject(textPage, index));
      if (id && !ids.includes(id)) ids.push(id);
      if (!box) continue;
      const [left, right, bottom, top] = box as [
        number,
        number,
        number,
        number,
      ];
      if (left === right || bottom === top) continue;
      const current = rects.at(-1);
      // Chars on one line share their vertical extent; a new line starts a new box.
      if (current && bottom < current.top && top > current.bottom) {
        current.left = Math.min(current.left, left);
        current.right = Math.max(current.right, right);
        current.bottom = Math.min(current.bottom, bottom);
        current.top = Math.max(current.top, top);
      } else rects.push({ left, bottom, right, top });
    }
    const pageRects = rects.map((rect) =>
      roundRect(
        userRectToPage(geometry, rect.left, rect.bottom, rect.right, rect.top),
      ),
    );
    return {
      pageIndex,
      text,
      rects:
        pageRects.length > 0
          ? pageRects
          : [roundRect(unionRects([{ x: 0, y: 0, width: 0, height: 0 }]))],
      elementIds: ids,
    };
  }
}
