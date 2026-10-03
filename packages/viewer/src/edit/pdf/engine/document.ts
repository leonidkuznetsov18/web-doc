import { ViewerError } from "../../../errors.js";
import type {
  EngineBatch,
  EngineChange,
  MaterializedDocument,
  RestoreTarget,
} from "../../engine.js";
import { invalidOperationError, parseReference } from "../../operations.js";
import type {
  EditFindOptions,
  EditOperation,
  ElementQuery,
  OperationIssue,
  PagePoint,
  PageRect,
  TextPosition,
  TextRange,
  TextTarget,
  TextAnchorMigration,
} from "../../types.js";
import type {
  PageLayout,
  PdfElement,
  PdfOperation,
  PdfTextParagraph,
  TextFont,
  TextLayout,
} from "../types.js";
import type { EditWorkerBitmap } from "../../../worker-protocol.js";
import {
  layoutOf,
  layoutsOf,
  positionIn,
  rectsOf,
  type TextPageScan,
} from "./layout.js";
import {
  markIsFresh,
  OBJECT_TEXT,
  readMark,
  scanPage,
  type MarkParams,
  type ObjectRecord,
} from "./elements.js";
import { FontLibrary, TextMeasurer } from "./fonts.js";
import { textFaceOf, type TextFace } from "./text-font.js";
import { fontRequestsOf } from "./text-box.js";
import {
  issueCollector,
  type ElementLocation,
  type OperationContext,
  type OperationHandler,
} from "./operations.js";
import { insertTextBox } from "./text-box.js";
import {
  discoverParagraphs,
  paragraphElement,
  type ParagraphTarget,
} from "./paragraph.js";
import { replaceParagraphText } from "./paragraph-edit.js";
import { replaceText, setTextStyle } from "./existing-text.js";
import { deleteElement, moveElement, resizeElement } from "./transform.js";
import { deletePage, insertPage, movePage, rotatePage } from "./pages.js";
import { insertShape, setShapeStyle } from "./shapes.js";
import { compactPdf, PdfCompactionError } from "./compact.js";
import { ImageCache, insertImage } from "./images.js";
import { AssetStore, type AssetSource } from "../../assets.js";
import { insertTable, setTableCell, tableSpecOf } from "./tables.js";
import { defaultResourceLimits } from "../../../limits.js";
import type { ResourceLimits } from "../../../contracts.js";
import {
  displayedSize,
  rectContains,
  rectsIntersect,
  round,
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
  paragraphs?: ReadonlyMap<string, ParagraphTarget>;
}

/**
 * The working copy of one PDF: a PDFium document rebuilt from the original
 * bytes plus the applied batches, with a stable id for every page and object.
 * Everything a query answers comes from PDFium; the model only remembers
 * identities and caches.
 */
/** Batches arrive as plain JSON; unknown operations are reported, not typed away. */
type PdfOrUnknownOperation = PdfOperation | EditOperation;

/** Properties of the original that a change invalidates or leaves behind. */
export type DocumentFeature = "docmdp" | "tagged" | "pdfa";

/**
 * Looks for a DocMDP certification, a tagged structure and a PDF/A claim.
 * Signature dictionaries and XMP metadata are stored uncompressed, so the
 * raw bytes answer for those; tagging also shows as a structure tree.
 */
/** FPDFBitmap_CreateEx pixel format with alpha. */
const BITMAP_BGRA = 4;
/** FPDF_RenderPageBitmap flag: draw annotations, as PDF.js does. */
const RENDER_ANNOTATIONS = 0x01;

function detectFeatures(
  pdfium: Pdfium,
  document: number,
  original: Uint8Array,
): DocumentFeature[] {
  const features: DocumentFeature[] = [];
  if (containsAscii(original, "/DocMDP")) features.push("docmdp");
  const { lib } = pdfium;
  let tagged = containsAscii(original, "/Marked true");
  if (!tagged && lib.FPDF_GetPageCount(document) > 0) {
    const page = lib.FPDF_LoadPage(document, 0);
    if (page) {
      const tree = lib.FPDF_StructTree_GetForPage(page);
      if (tree) {
        tagged = lib.FPDF_StructTree_CountChildren(tree) > 0;
        lib.FPDF_StructTree_Close(tree);
      }
      lib.FPDF_ClosePage(page);
    }
  }
  if (tagged) features.push("tagged");
  if (containsAscii(original, "pdfaid:part")) features.push("pdfa");
  return features;
}

function containsAscii(bytes: Uint8Array, needle: string): boolean {
  const first = needle.charCodeAt(0);
  const limit = bytes.length - needle.length;
  for (
    let at = bytes.indexOf(first);
    at >= 0 && at <= limit;
    at = bytes.indexOf(first, at + 1)
  ) {
    let index = 1;
    while (
      index < needle.length &&
      bytes[at + index] === needle.charCodeAt(index)
    )
      index += 1;
    if (index === needle.length) return true;
  }
  return false;
}

const FEATURE_NOTES: Readonly<Record<DocumentFeature, string>> = {
  docmdp: "a DocMDP certification that any change invalidates",
  tagged: "a tagged structure that inserted content does not join",
  pdfa: "a PDF/A claim that inserted standard fonts do not meet",
};

function firstChangeMessage(
  signatures: number,
  features: readonly DocumentFeature[],
): string {
  const notes = [
    ...(signatures > 0
      ? [
          `${signatures} digital signature${signatures === 1 ? "" : "s"} that will not cover the edited revision`,
        ]
      : []),
    ...features.map((feature) => FEATURE_NOTES[feature]),
  ];
  return `The document carries ${notes.join("; ")}`;
}

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
  replaceParagraphText: replaceParagraphText as OperationHandler,
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
  readonly #compact: (bytes: Uint8Array) => Uint8Array;
  readonly #original: Uint8Array;
  readonly #fonts: FontLibrary;
  readonly images = new ImageCache();
  readonly #limits: ResourceLimits;
  readonly #assets: AssetSource;
  /** Signature fields in the original; an edit leaves them uncovering the new revision. */
  readonly #signatures: number;
  /** What the first change will break or leave behind: certification, tagging, PDF/A. */
  readonly #features: readonly DocumentFeature[];
  /** What the working copy was opened from: the original, or a checkpoint. */
  #base: Uint8Array;
  #document: PdfiumDocument;
  #measurer: TextMeasurer;
  #pages: PageRecord[];
  #batches = 0;
  /** Browser faces of the document's fonts, by font program; see `textFont`. */
  readonly #faces = new Map<string, TextFace>();

  constructor(
    pdfium: Pdfium,
    original: Uint8Array,
    fonts: FontLibrary = new FontLibrary(async () => {
      throw new Error("No font source is configured");
    }),
    limits: ResourceLimits = defaultResourceLimits,
    assets: AssetSource = new AssetStore(),
    compact: (bytes: Uint8Array) => Uint8Array = compactPdf,
  ) {
    this.#compact = compact;
    this.#pdfium = pdfium;
    this.#original = original;
    this.#base = original;
    this.#fonts = fonts;
    this.#limits = limits;
    this.#assets = assets;
    this.#document = pdfium.openDocument(original);
    this.#measurer = new TextMeasurer(pdfium, this.#document.handle);
    this.#pages = this.#originalPages();
    this.#signatures = pdfium.lib.FPDF_GetSignatureCount(this.#document.handle);
    this.#features = detectFeatures(pdfium, this.#document.handle, original);
  }

  /** Features of the original a change affects, see `DocumentFeature`. */
  get features(): readonly DocumentFeature[] {
    return this.#features;
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
    const textAnchorMigrations: TextAnchorMigration[] = [];
    const changedPages = new Set<number>();
    const warnings: EngineChange["warnings"][number][] = [];
    // The first change of a signed, certified, tagged or PDF/A file is the
    // point where those properties stop holding for what is shown; the
    // warning names them once.
    if (this.#batches === 0 && (this.#signatures > 0 || this.#features.length))
      warnings.push({
        code: "fidelity-degraded",
        message: firstChangeMessage(this.#signatures, this.#features),
        details: { signatures: this.#signatures, features: this.#features },
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
      for (const migration of result.textAnchorMigrations ?? [])
        textAnchorMigrations.push({ ...migration, operationIndex });
      for (const pageIndex of result.changedPages) changedPages.add(pageIndex);
      warnings.push(...result.warnings);
    });
    this.#batches += 1;
    return {
      createdIds,
      removedIds,
      ...(textAnchorMigrations.length ? { textAnchorMigrations } : {}),
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

  /**
   * Bytes of the current state. The viewer reopens the incremental form,
   * which is cheap to produce and read. A save defaults to a full rewrite —
   * compacted, so deleted content is gone and the bytes do not depend on
   * which pages were read — unless the file is signed, where the incremental
   * form keeps the signed revision intact. Without changes either form is
   * the bytes the document was opened from.
   */
  materialize(
    purpose: "show" | "save" = "show",
    mode?: "incremental" | "full",
  ): Uint8Array {
    return this.materializeDocument(purpose, mode).bytes;
  }

  /**
   * The bytes of the current state. A full save goes through the compaction
   * pass; when the pass cannot read PDFium's output, the uncompacted full
   * save stands in — silently for the display copy, with a
   * `privacy-not-guaranteed` warning for a save, since deleted content may
   * then remain recoverable in the file (decided 2026-10-02). Signed files
   * take the incremental form for both purposes so the signed revision
   * stays intact in every base the session may restore from.
   */
  materializeDocument(
    purpose: "show" | "save" = "show",
    mode?: "incremental" | "full",
  ): MaterializedDocument {
    const chosen =
      purpose === "show"
        ? this.#signatures > 0
          ? "incremental"
          : "full"
        : (mode ?? (this.#signatures > 0 ? "incremental" : "full"));
    if (this.#batches === 0 && this.#base === this.#original)
      return { bytes: this.#original.slice(), warnings: [] };
    if (chosen === "incremental")
      return {
        bytes:
          this.#batches === 0
            ? this.#base.slice()
            : this.#document.save("incremental"),
        warnings: [],
      };
    const full = this.#document.save("full");
    try {
      return { bytes: this.#compact(full), warnings: [] };
    } catch (error) {
      if (!(error instanceof PdfCompactionError)) throw error;
      if (purpose === "show") {
        console.warn(
          `web-doc: ${error.message}; the uncompacted full save is shown`,
        );
        return { bytes: full, warnings: [] };
      }
      return {
        bytes: full,
        warnings: [
          {
            code: "privacy-not-guaranteed",
            message:
              "The full save could not be compacted; deleted or replaced content may remain recoverable in the file",
            details: { reason: "pdf-compaction", cause: error.message },
          },
        ],
      };
    }
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
    const element = this.#elementsOf(pageIndex).find(
      (element) => element.id === id,
    );
    if (element) return element;
    const target = this.#pages[pageIndex]?.paragraphs?.get(id);
    return target?.paragraph.id === id
      ? paragraphElement(target.paragraph)
      : undefined;
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
      this.#scanText(pageIndex, ({ textPage, geometry, byObject, offsets }) => {
        for (const match of this.#matches(textPage, query, options)) {
          if (targets.length >= limit) break;
          targets.push(
            this.#target(
              textPage,
              pageIndex,
              geometry,
              match,
              byObject,
              offsets,
            ),
          );
        }
      });
    return targets;
  }

  /** Resolves the canonical imported paragraph, when its rows can be grouped safely. */
  textParagraph(elementId: string): PdfTextParagraph | undefined {
    return this.#paragraphOf(elementId)?.paragraph;
  }

  #paragraphOf(id: string): ParagraphTarget | undefined {
    const pageIndex = this.#pageIndexOf(id);
    if (pageIndex === undefined) return undefined;
    this.#elementsOf(pageIndex);
    return this.#pages[pageIndex]?.paragraphs?.get(id);
  }

  /** Lines, glyph boxes and styles of a text, text box, paragraph or table element. */
  textLayout(elementId: string): TextLayout | undefined {
    const pageIndex = this.#pageIndexOf(elementId);
    if (pageIndex === undefined) return undefined;
    const paragraph = this.#paragraphOf(elementId);
    return this.#scanText(pageIndex, (raw) => {
      const scan =
        paragraph?.paragraph.id === elementId
          ? this.#paragraphScan(raw, paragraph)
          : raw;
      const element = scan.elements.find((entry) => entry.id === elementId);
      const layout = element
        ? layoutOf(this.#pdfium, scan, element)
        : undefined;
      return layout && paragraph?.paragraph.id === elementId
        ? {
            ...layout,
            lines: [...layout.lines]
              .sort((a, b) => a.range.start.offset - b.range.start.offset)
              .map((line) => ({
                ...line,
                fontSize: paragraph.paragraph.textStyle.fontSize,
              })),
          }
        : layout;
    });
  }

  /** The browser face of the font a text or text box element is drawn in, see `TextFont`. */
  textFont(elementId: string): TextFont | undefined {
    const location = this.#locate(elementId);
    const element = location && this.getElement(elementId);
    if (!element || (element.kind !== "text" && element.kind !== "textBox"))
      return undefined;
    const { lib } = this.#pdfium;
    return this.#withPage(location.pageIndex, (page) => {
      // A text box's lines share their font; its first line answers.
      const object = location.indexes
        .map((index) => lib.FPDFPage_GetObject(page, index))
        .find((entry) => lib.FPDFPageObj_GetType(entry) === OBJECT_TEXT);
      return object === undefined
        ? undefined
        : { elementId, ...textFaceOf(this.#pdfium, object, this.#faces) };
    });
  }

  /** The layouts of every text element on a page, with the page's displayed size. */
  pageLayout(pageIndex: number): PageLayout | undefined {
    if (pageIndex < 0 || pageIndex >= this.#pages.length) return undefined;
    return this.#scanText(pageIndex, (scan) => {
      const size = displayedSize(scan.geometry);
      return {
        pageIndex,
        width: round(size.width),
        height: round(size.height),
        layouts: layoutsOf(this.#pdfium, scan),
      };
    });
  }

  /** The caret position nearest to a page-space point; none on a page without text. */
  positionAt(pageIndex: number, point: PagePoint): TextPosition | undefined {
    if (pageIndex < 0 || pageIndex >= this.#pages.length) return undefined;
    return this.#scanText(pageIndex, (scan) =>
      positionIn(this.#pdfium, scan, point),
    );
  }

  /**
   * Renders a page with the listed elements inactive, as RGBA over white at
   * `scale` device pixels per point. The objects are reactivated before the
   * call returns, so nothing about the document changes.
   */
  renderPageWithout(
    pageIndex: number,
    elementIds: readonly string[],
    scale: number,
  ): EditWorkerBitmap {
    if (pageIndex < 0 || pageIndex >= this.#pages.length)
      throw new ViewerError("invalid-operation", "No such page", {
        details: { pageIndex },
      });
    if (!Number.isFinite(scale) || scale <= 0)
      throw new ViewerError("invalid-operation", "The scale must be positive", {
        details: { scale },
      });
    const { lib } = this.#pdfium;
    return this.#withPage(pageIndex, (page, _textPage, geometry) => {
      const size = displayedSize(geometry);
      const width = Math.max(1, Math.ceil(size.width * scale));
      const height = Math.max(1, Math.ceil(size.height * scale));
      const pixels = width * height;
      if (
        !Number.isSafeInteger(pixels) ||
        pixels > this.#limits.maxDecodedPixels
      )
        throw new ViewerError(
          "resource-limit",
          "The rendered page exceeds maxDecodedPixels",
          { details: { actual: pixels, limit: this.#limits.maxDecodedPixels } },
        );
      const wanted = new Set(
        elementIds.flatMap((id) => {
          const paragraph = this.#paragraphOf(id)?.paragraph;
          return paragraph?.id === id ? paragraph.memberIds : [id];
        }),
      );
      const records = this.#objectsOf(pageIndex, page);
      const suppressed: {
        readonly object: number;
        readonly active: boolean;
      }[] = [];
      records.forEach((record, index) => {
        if (!wanted.has(record.id)) return;
        const object = lib.FPDFPage_GetObject(page, index);
        const active =
          this.#pdfium.readNumbers(1, "i32", ([pointer]) =>
            lib.FPDFPageObj_GetIsActive(object, pointer!),
          )?.[0] !== 0;
        suppressed.push({ object, active });
        lib.FPDFPageObj_SetIsActive(object, false);
      });
      try {
        return {
          pageIndex,
          scale,
          width,
          height,
          data: this.#render(page, width, height),
        };
      } finally {
        for (const { object, active } of suppressed)
          lib.FPDFPageObj_SetIsActive(object, active);
      }
    });
  }

  /** RGBA pixels of a loaded page over white, through a BGRA bitmap in WASM memory. */
  #render(page: number, width: number, height: number): ArrayBuffer {
    const { lib } = this.#pdfium;
    const stride = width * 4;
    const buffer = this.#pdfium.malloc(stride * height);
    try {
      const bitmap = lib.FPDFBitmap_CreateEx(
        width,
        height,
        BITMAP_BGRA,
        buffer,
        stride,
      );
      if (!bitmap)
        throw new ViewerError(
          "render-failed",
          "PDFium could not create the bitmap",
          {
            details: { width, height },
          },
        );
      try {
        lib.FPDFBitmap_FillRect(bitmap, 0, 0, width, height, 0xffffffff);
        lib.FPDF_RenderPageBitmap(
          bitmap,
          page,
          0,
          0,
          width,
          height,
          0,
          RENDER_ANNOTATIONS,
        );
      } finally {
        lib.FPDFBitmap_Destroy(bitmap);
      }
      const bgra = this.#pdfium.readBytes(buffer, stride * height);
      const rgba = new Uint8Array(bgra.byteLength);
      for (let offset = 0; offset < bgra.byteLength; offset += 4) {
        rgba[offset] = bgra[offset + 2]!;
        rgba[offset + 1] = bgra[offset + 1]!;
        rgba[offset + 2] = bgra[offset]!;
        rgba[offset + 3] = bgra[offset + 3]!;
      }
      return rgba.buffer;
    } finally {
      this.#pdfium.free(buffer);
    }
  }

  /** The rectangles a range covers, one per line fragment, in reading order. */
  rangeRects(range: TextRange): PageRect[] {
    const pageIndex = this.#pageIndexOf(range.start.elementId);
    if (
      pageIndex === undefined ||
      pageIndex !== this.#pageIndexOf(range.end.elementId)
    )
      return [];
    const paragraph =
      range.start.elementId === range.end.elementId
        ? this.#paragraphOf(range.start.elementId)
        : undefined;
    return this.#scanText(pageIndex, (scan) =>
      rectsOf(
        this.#pdfium,
        paragraph?.paragraph.id === range.start.elementId
          ? this.#paragraphScan(scan, paragraph)
          : scan,
        range,
      ),
    );
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
      assets: this.#assets,
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
      // handed out again. A file saved by an earlier session carries the ids
      // that session numbered from 1 as well, so one already on the page
      // gets the first free `~n` instead (ACTION-886); the saved objects are
      // part of the base, so a replay meets them and picks the same id.
      newId: (pageIndex, suffix = "") =>
        `${this.#unusedId(pageIndex, `${this.#pages[pageIndex]!.key}:n${stateId}.${operationIndex}.${created++}`)}${suffix}`,
      withPage: (pageIndex, use) => this.#writePage(pageIndex, use),
      readPage: (pageIndex, use) => this.#withPage(pageIndex, use),
      paragraph: (id) => this.#paragraphOf(id),
      pageElements: (pageIndex) => this.#elementsOf(pageIndex),
      appendObjects: (pageIndex, records) => {
        const page = this.#pages[pageIndex]!;
        page.objects = [...(page.objects ?? []), ...records];
        delete page.elements;
        delete page.paragraphs;
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
        delete page.paragraphs;
      },
      spliceObjects: (pageIndex, start, count, records) => {
        const page = this.#pages[pageIndex]!;
        const objects = [...(page.objects ?? [])];
        objects.splice(start, count, ...records);
        page.objects = objects;
        delete page.elements;
        delete page.paragraphs;
      },
    };
  }

  /** Cached elements carry page indexes, which a structure change makes stale. */
  #forgetElements(): void {
    for (const page of this.#pages) {
      delete page.elements;
      delete page.paragraphs;
    }
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
    if (record) return { pageIndex, indexes, record };
    const paragraph = this.#paragraphOf(id);
    const first = paragraph && records[paragraph.indexes[0] ?? -1];
    return paragraph?.paragraph.id === id && first
      ? {
          pageIndex,
          indexes: paragraph.indexes,
          record: { ...first, id, mark: paragraph.spec },
        }
      : undefined;
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
      delete this.#pages[pageIndex]!.paragraphs;
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
      record.paragraphs = discoverParagraphs(
        this.#pdfium,
        page,
        pageIndex,
        geometry,
        this.#objectsOf(pageIndex, page),
        elements,
        textPage,
        this.#measurer,
      );
      record.elements = Object.freeze(
        elements.map((element) => {
          const paragraph = record.paragraphs?.get(element.id)?.paragraph;
          return paragraph && paragraph.id !== element.id
            ? { ...element, textEditingTarget: paragraph.id }
            : element;
        }),
      );
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
    // are plain objects again. So is a group whose objects no longer match
    // the inputs its mark stores.
    const heads = new Map<string, MarkParams | undefined>();
    const members = new Map<string, number[]>();
    objects.forEach((object, index) => {
      if (!object.mark) return;
      if (!heads.has(object.id))
        heads.set(
          object.id,
          object.mark.kind === "table"
            ? tableSpecOf(object.mark) && object.mark
            : object.mark,
        );
      members.set(object.id, [...(members.get(object.id) ?? []), index]);
    });
    const stale = new Set<string>();
    if (members.size > 0) {
      const geometry = this.#geometryOf(pageIndex, page);
      const textPage = lib.FPDFText_LoadPage(page);
      try {
        for (const [id, indexes] of members) {
          const head = heads.get(id);
          if (
            !head ||
            !markIsFresh(this.#pdfium, page, textPage, geometry, head, indexes)
          )
            stale.add(id);
        }
      } finally {
        lib.FPDFText_ClosePage(textPage);
      }
    }
    record.objects = objects.map((object, index) =>
      object.mark && stale.has(object.id)
        ? {
            id: `${record.key}:o${index}`,
            type: object.type,
            staleMarkId: object.id,
          }
        : object,
    );
    return record.objects;
  }

  /**
   * `id`, or its first free `~n` when an object saved by an earlier session
   * has it already. Only ids a session numbered (`<key>:n…`) can collide, so
   * only those are gathered; the objects are read without the page's text.
   */
  #unusedId(pageIndex: number, id: string): string {
    const record = this.#pages[pageIndex]!;
    const objects = record.objects ?? this.#loadObjects(pageIndex);
    const numbered = `${record.key}:n`;
    const taken = new Set<string>();
    for (const object of objects) {
      if (object.id.startsWith(numbered)) taken.add(object.id);
      if (object.staleMarkId?.startsWith(numbered))
        taken.add(object.staleMarkId);
    }
    if (!taken.has(id)) return id;
    let suffix = 1;
    while (taken.has(`${id}~${suffix}`)) suffix += 1;
    return `${id}~${suffix}`;
  }

  /** A page's objects, loading the page alone when they are not known yet. */
  #loadObjects(pageIndex: number): ObjectRecord[] {
    const { lib } = this.#pdfium;
    const page = lib.FPDF_LoadPage(this.#document.handle, pageIndex);
    if (!page)
      throw new ViewerError("render-failed", "PDFium could not load the page", {
        details: { pageIndex },
      });
    try {
      return this.#objectsOf(pageIndex, page);
    } finally {
      lib.FPDF_ClosePage(page);
    }
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

  #paragraphScan(scan: TextPageScan, target: ParagraphTarget): TextPageScan {
    const paragraph = target.paragraph;
    if (paragraph.memberIds.includes(paragraph.id)) return scan;
    const members = new Map(
      paragraph.members.map((member) => [member.elementId, member]),
    );
    return {
      ...scan,
      elements: [
        ...scan.elements.filter((element) => !members.has(element.id)),
        paragraphElement(paragraph),
      ],
      offsets: scan.offsets.map((position) => {
        const member = position && members.get(position.elementId);
        return member && position
          ? {
              elementId: paragraph.id,
              offset: Math.min(member.end, member.start + position.offset),
            }
          : position;
      }),
    };
  }

  /** Loads a page with its text page and the character-to-element mapping. */
  #scanText<T>(pageIndex: number, use: (scan: TextPageScan) => T): T {
    return this.#withPage(pageIndex, (page, textPage, geometry) => {
      const records = this.#objectsOf(pageIndex, page);
      const { byObject, elements } = scanPage(
        this.#pdfium,
        page,
        textPage,
        pageIndex,
        geometry,
        records,
      );
      const offsets = this.#charOffsets(
        page,
        textPage,
        byObject,
        new Map(elements.map((element) => [element.id, element.text ?? ""])),
      );
      return use({ page, textPage, geometry, byObject, elements, offsets });
    });
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

  /**
   * For every character of the text page: the element it belongs to and its
   * offset in that element's text, so matches can be reported as ranges.
   * Lines of a text box or cells of a table are separate objects; their
   * offsets are found by locating each object's text inside the element's.
   */
  #charOffsets(
    page: number,
    textPage: number,
    byObject: ReadonlyMap<number, string>,
    elementTexts: ReadonlyMap<string, string>,
  ): readonly (TextPosition | undefined)[] {
    const { lib } = this.#pdfium;
    const count = lib.FPDFText_CountChars(textPage);
    const positions: (TextPosition | undefined)[] = [];
    const seen = new Map<number, number>();
    const cursors = new Map<string, number>();
    const bases = new Map<number, number>();
    for (let index = 0; index < count; index += 1) {
      const object = lib.FPDFText_GetTextObject(textPage, index);
      const elementId = byObject.get(object);
      if (!object || elementId === undefined) {
        positions.push(undefined);
        continue;
      }
      if (!bases.has(object)) {
        const own = this.#pdfium.readWideString((buffer, bytes) =>
          lib.FPDFTextObj_GetText(object, textPage, buffer, bytes),
        );
        const whole = elementTexts.get(elementId) ?? "";
        const from = cursors.get(elementId) ?? 0;
        const at = whole.indexOf(own, from);
        const base = at >= 0 ? at : whole.indexOf(own.trim(), from);
        bases.set(object, base >= 0 ? base : from);
        cursors.set(elementId, (base >= 0 ? base : from) + own.length);
      }
      const within = seen.get(object) ?? 0;
      seen.set(object, within + 1);
      positions.push({ elementId, offset: bases.get(object)! + within });
    }
    void page;
    return positions;
  }

  #target(
    textPage: number,
    pageIndex: number,
    geometry: PageGeometry,
    match: { readonly start: number; readonly end: number },
    byObject: ReadonlyMap<number, string>,
    offsets: readonly (TextPosition | undefined)[],
  ): TextTarget {
    const { lib } = this.#pdfium;
    const ranges: TextRange[] = [];
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
      const position = offsets[index];
      if (position) {
        const last = ranges.at(-1);
        if (
          last &&
          last.end.elementId === position.elementId &&
          last.end.offset === position.offset
        )
          ranges[ranges.length - 1] = {
            start: last.start,
            end: { elementId: position.elementId, offset: position.offset + 1 },
          };
        else
          ranges.push({
            start: position,
            end: { elementId: position.elementId, offset: position.offset + 1 },
          });
      }
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
      ranges,
    };
  }
}
