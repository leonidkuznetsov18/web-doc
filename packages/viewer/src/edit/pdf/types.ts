import type { TextSelection } from "../../contracts.js";
import type {
  ApplyOptions,
  BinaryData,
  EditElement,
  EditReceipt,
  EditSessionBase,
  PagePoint,
  PageRect,
  ReadItem,
  ReadOptions,
  ReadResult,
  SavedDocument,
  SaveOptions,
  TextPosition,
  TextRange,
} from "../types.js";

export type PdfElementKind =
  | "text" // one text object as stored in the file (often a word or a line)
  | "image"
  | "shape" // a path object
  | "textBox" // a text box created by insertTextBox
  | "paragraph" // confidently grouped imported text, retained after editing
  | "table" // a table created by insertTable
  | "other"; // shadings, form XObjects and anything else

export type PdfTextAlign = "left" | "center" | "right";

export interface PdfTextStyle {
  readonly fontFamily: string;
  /** Points. */
  readonly fontSize: number;
  readonly bold: boolean;
  readonly italic: boolean;
  /** `#RRGGBB`. */
  readonly color: string;
  /** Text boxes only. */
  readonly align?: PdfTextAlign;
  /** Text boxes and paragraphs; a multiple of the font size. */
  readonly lineHeight?: number;
}

export interface PdfShapeStyle {
  readonly stroke?: { readonly color: string; readonly width: number };
  readonly fill?: { readonly color: string };
}

export interface PdfElement extends EditElement {
  readonly kind: PdfElementKind;
  /** Canonical native paragraph target, when this imported row belongs to one. */
  readonly textEditingTarget?: string;
  /** Present for `text`, `textBox` and `paragraph`. */
  readonly textStyle?: PdfTextStyle;
  /** Present for `shape` and `table`. */
  readonly shapeStyle?: PdfShapeStyle;
  /** Present for `table`. */
  readonly table?: { readonly rows: readonly (readonly string[])[] };
}

export interface PdfTextBoxStyle {
  /** "Helvetica" (default), "Times" or "Courier"; registered families follow. */
  readonly fontFamily?: string;
  /** 1–500 points, default 12. */
  readonly fontSize?: number;
  readonly bold?: boolean;
  readonly italic?: boolean;
  /** `#RRGGBB`, default "#000000". */
  readonly color?: string;
  readonly align?: PdfTextAlign;
  /** Multiple of the font size, 0.5–5, default 1.2. */
  readonly lineHeight?: number;
}

export interface InsertTextBoxOperation {
  readonly op: "insertTextBox";
  readonly pageIndex: number;
  readonly rect: PageRect;
  readonly text: string;
  readonly style?: PdfTextBoxStyle;
}

export interface ReplaceTextOperation {
  readonly op: "replaceText";
  /** A `textBox` or `text` element. */
  readonly target: string;
  readonly text: string;
  /**
   * The part of the target's text to replace, both ends on the target;
   * absent, the whole text. A text object is split around the range only
   * when its font cannot draw the new text; the parts keep their font, size,
   * colour and baseline and the first part keeps the id.
   */
  readonly range?: TextRange;
}

/** A confidently resolved homogeneous paragraph of imported horizontal text. */
export interface PdfTextParagraph {
  readonly id: string;
  readonly pageIndex: number;
  readonly text: string;
  readonly bounds: PageRect;
  readonly textStyle: PdfTextStyle & { readonly lineHeight: number };
  readonly memberIds: readonly string[];
  /** Half-open UTF-16 spans in the paragraph's logical text. */
  readonly members: readonly {
    readonly elementId: string;
    readonly start: number;
    readonly end: number;
  }[];
}

export interface ReplaceParagraphTextOperation {
  readonly op: "replaceParagraphText";
  readonly target: string;
  readonly text: string;
  /** Both endpoints name the paragraph id, not an individual visual row. */
  readonly range?: TextRange;
}

export interface SetTextStyleOperation {
  readonly op: "setTextStyle";
  /** A `textBox` or `text` element. */
  readonly target: string;
  readonly style: PdfTextBoxStyle;
}

export interface ResizeElementOperation {
  readonly op: "resizeElement";
  readonly target: string;
  /** New bounds in page space; a text box is laid out again inside them. */
  readonly rect: PageRect;
}

export interface MoveElementOperation {
  readonly op: "moveElement";
  readonly target: string;
  /** New top-left corner of the element's bounds, in page space. */
  readonly to?: PagePoint;
  /** Offset in page space. Exactly one of `to` and `by` is given. */
  readonly by?: { readonly dx: number; readonly dy: number };
}

export interface DeleteElementOperation {
  readonly op: "deleteElement";
  readonly target: string;
}

export interface InsertPageOperation {
  readonly op: "insertPage";
  /** Position of the new page, 0 to the current page count. */
  readonly index: number;
  /** Points; defaults to the size of the page before, else after, the position. */
  readonly size?: { readonly width: number; readonly height: number };
}

export interface DeletePageOperation {
  readonly op: "deletePage";
  readonly pageIndex: number;
}

export interface MovePageOperation {
  readonly op: "movePage";
  readonly from: number;
  /** The page's index after the move. */
  readonly to: number;
}

export interface PdfStroke {
  readonly color: string;
  /** Points; 0 draws the thinnest line the device can. */
  readonly width: number;
}

export interface PdfFill {
  readonly color: string;
}

export type InsertShapeOperation =
  | {
      readonly op: "insertShape";
      readonly pageIndex: number;
      readonly shape: "rectangle" | "ellipse";
      readonly rect: PageRect;
      readonly stroke?: PdfStroke;
      readonly fill?: PdfFill;
    }
  | {
      readonly op: "insertShape";
      readonly pageIndex: number;
      readonly shape: "line";
      readonly from: PagePoint;
      readonly to: PagePoint;
      readonly stroke?: PdfStroke;
    };

export interface InsertImageOperation {
  readonly op: "insertImage";
  readonly pageIndex: number;
  readonly rect: PageRect;
  /** PNG or JPEG bytes; base64 over pure-JSON transports. */
  readonly data: BinaryData;
  readonly mimeType: "image/png" | "image/jpeg";
}

export interface PdfTableStyle {
  /** "Helvetica" (default), "Times", "Courier" or a registered family. */
  readonly fontFamily?: string;
  /** 1–500 points, default 10. */
  readonly fontSize?: number;
  /** Text colour, `#RRGGBB`, default "#000000". */
  readonly color?: string;
  /** Grid colour, default "#000000". */
  readonly borderColor?: string;
  /** Grid stroke in points, 0–20, default 0.75. */
  readonly borderWidth?: number;
  /** Space between a cell's edges and its text, 0–100, default 4. */
  readonly cellPadding?: number;
  /** Fill of the first row; none by default. */
  readonly headerFill?: string;
}

export interface InsertTableOperation {
  readonly op: "insertTable";
  readonly pageIndex: number;
  /** Top-left corner in page space. */
  readonly at: PagePoint;
  /** Total width in points; the height follows the wrapped cell text. */
  readonly width: number;
  /** Cell text by row; every row has the same number of cells. 1–100 rows, 1–20 columns. */
  readonly rows: readonly (readonly string[])[];
  /** Relative column weights, one per column; equal when omitted. */
  readonly columnWidths?: readonly number[];
  readonly style?: PdfTableStyle;
}

export interface SetTableCellOperation {
  readonly op: "setTableCell";
  /** A `table` element. */
  readonly target: string;
  readonly row: number;
  readonly column: number;
  /** The cell's new text; empty clears the cell. */
  readonly text: string;
}

export interface SetShapeStyleOperation {
  readonly op: "setShapeStyle";
  readonly target: string;
  /** A new stroke, `null` to remove it, or absent to keep it. */
  readonly stroke?: PdfStroke | null;
  /** A new fill, `null` to remove it, or absent to keep it. */
  readonly fill?: PdfFill | null;
}

/**
 * Sets a page's rotation, or turns it from where it stands: a host offering
 * "rotate" does not need to know the angle a file was saved with, and a
 * replay of the history turns from the same angle again.
 */
export type RotatePageOperation =
  | {
      readonly op: "rotatePage";
      readonly pageIndex: number;
      /** Absolute clockwise rotation in degrees. */
      readonly rotation: 0 | 90 | 180 | 270;
      readonly by?: never;
    }
  | {
      readonly op: "rotatePage";
      readonly pageIndex: number;
      /** Clockwise turn in degrees from the page's current rotation. */
      readonly by: 90 | 180 | 270;
      readonly rotation?: never;
    };

/** The PDF operation union; operations are added as they ship. */
export type PdfOperation =
  | InsertTextBoxOperation
  | ReplaceTextOperation
  | ReplaceParagraphTextOperation
  | SetTextStyleOperation
  | ResizeElementOperation
  | MoveElementOperation
  | DeleteElementOperation
  | InsertPageOperation
  | DeletePageOperation
  | MovePageOperation
  | RotatePageOperation
  | InsertShapeOperation
  | SetShapeStyleOperation
  | InsertImageOperation
  | InsertTableOperation
  | SetTableCellOperation;

export interface PdfSaveOptions extends SaveOptions {
  /**
   * `full` rewrites the file, so deleted content is gone and the bytes do not
   * depend on which pages were read; `incremental` appends to the original
   * bytes and keeps earlier signed revisions intact. The default is `full`
   * for a document without signature fields and `incremental` for a signed one.
   */
  readonly mode?: "full" | "incremental";
}

/** One drawn character of a layout line. */
export interface TextLayoutGlyph {
  /** Offset of the character in `EditElement.text`. */
  readonly offset: number;
  /** Tight box of the glyph in page space; a space takes its advance box. */
  readonly box: PageRect;
  /** Advance width along the baseline, in points. */
  readonly advance: number;
}

/** One line of a text element: one PDFium text object, as the file stores it. */
export interface TextLayoutLine {
  /** The part of the element's text the line draws, half-open. */
  readonly range: TextRange;
  readonly text: string;
  /** Union of the glyph boxes, in page space. */
  readonly bounds: PageRect;
  /** Start of the baseline, in page space. */
  readonly baseline: PagePoint;
  readonly glyphs: readonly TextLayoutGlyph[];
  readonly fontFamily: string;
  readonly fontSize: number;
  readonly color: string;
}

/** The drawn geometry of a `text`, `textBox` or `table` element. */
export interface TextLayout {
  readonly elementId: string;
  readonly pageIndex: number;
  /** Lines in reading order: a text box's lines, a table's cells. */
  readonly lines: readonly TextLayoutLine[];
}

/** Every text element of a page with its layout, plus the page's displayed size in points. */
export interface PageLayout {
  readonly pageIndex: number;
  readonly width: number;
  readonly height: number;
  /** In reading order. */
  readonly layouts: readonly TextLayout[];
}

/** A page rendered by PDFium in the worker: RGBA pixels, row-major, unpremultiplied over white. */
export interface PageBitmap {
  readonly pageIndex: number;
  /** Device pixels per point the page was rendered at. */
  readonly scale: number;
  readonly width: number;
  readonly height: number;
  readonly data: Uint8Array;
}

export interface RenderOptions extends ReadOptions {
  /** Device pixels per point, default 1; bounded by `maxDecodedPixels`. */
  readonly scale?: number;
}

/**
 * An operation's fields without its `op`, as the typed methods take them.
 * Distributes over unions so `insertShape` keeps its per-shape fields.
 */
export type Fields<T extends PdfOperation> = T extends unknown
  ? Omit<T, "op">
  : never;

export interface PdfEditSession extends EditSessionBase<
  PdfOperation,
  PdfElement
> {
  readonly format: "pdf";
  save(options?: PdfSaveOptions): Promise<SavedDocument>;
  /** Resolves an imported row or a paragraph id; undefined when grouping is unsafe. */
  getTextParagraph(
    elementId: string,
    options?: ReadOptions,
  ): Promise<ReadItem<PdfTextParagraph>>;
  /** Reflows the complete paragraph as one atomic history change. */
  replaceParagraphText(
    fields: Fields<ReplaceParagraphTextOperation>,
    options?: ApplyOptions,
  ): Promise<EditReceipt>;
  /** Lines, glyph boxes and styles of a `text`, `textBox` or `table` element. */
  getTextLayout(
    elementId: string,
    options?: ReadOptions,
  ): Promise<ReadItem<TextLayout>>;
  /** The text position nearest to a page-space point; none on a page without text. */
  positionAt(
    pageIndex: number,
    point: PagePoint,
    options?: ReadOptions,
  ): Promise<ReadItem<TextPosition>>;
  /** The rectangles a range covers, one per line fragment, in reading order. */
  rangeRects(
    range: TextRange,
    options?: ReadOptions,
  ): Promise<ReadResult<PageRect>>;
  /**
   * The page rendered by PDFium with the listed elements left out, so a host's
   * input surface can stand in for them on screen. Nothing is reopened and the
   * session's bytes stay as they are; unknown ids are ignored.
   */
  renderPageWithout(
    pageIndex: number,
    elementIds: readonly string[],
    options?: RenderOptions,
  ): Promise<ReadItem<PageBitmap>>;
  /** The layouts of every text element on a page, in one read. */
  getPageLayout(
    pageIndex: number,
    options?: ReadOptions,
  ): Promise<ReadItem<PageLayout>>;
  /**
   * Maps the viewer's text selection (PDF.js runs in page space) to elements
   * and ranges: a line whose box the run covers by half, else the one line
   * that contains the run, else a line whose folded text contains the run's.
   */
  elementsForSelection(
    selection: TextSelection,
    options?: ReadOptions,
  ): Promise<ReadResult<TextRange>>;
  /**
   * Elements under a point from the session's main-thread geometry cache,
   * without waiting: the last `getElements({ pageIndex })` result of the
   * page, refreshed after every committed change. A page never read answers
   * no items; `cachedPages` says which pages answer.
   */
  elementsAtSync(pageIndex: number, point: PagePoint): ReadResult<PdfElement>;
  /** Pages whose geometry the cache holds. */
  readonly cachedPages: readonly number[];
  /**
   * Where a range taken at `fromRevision` is now, after the batches, undos,
   * redos and resets of this session since then; none when its element is
   * gone or the revision is older than the session remembers.
   */
  mapRange(
    range: TextRange,
    fromRevision: number,
    options?: ReadOptions,
  ): Promise<ReadItem<TextRange>>;
  /** Lays `text` out inside `rect` as new text objects; the box's id is in `createdIds`. */
  insertTextBox(
    fields: Fields<InsertTextBoxOperation>,
    options?: ApplyOptions,
  ): Promise<EditReceipt>;
  /** Replaces the text of a text box (laid out again) or of a text object. */
  replaceText(
    fields: Fields<ReplaceTextOperation>,
    options?: ApplyOptions,
  ): Promise<EditReceipt>;
  /** Changes style fields; unspecified fields keep their value. */
  setTextStyle(
    fields: Fields<SetTextStyleOperation>,
    options?: ApplyOptions,
  ): Promise<EditReceipt>;
  /** Gives an element new bounds; text boxes reflow, other elements stretch. */
  resizeElement(
    fields: Fields<ResizeElementOperation>,
    options?: ApplyOptions,
  ): Promise<EditReceipt>;
  /** Moves an element to a point or by an offset, in page space. */
  moveElement(
    fields: Fields<MoveElementOperation>,
    options?: ApplyOptions,
  ): Promise<EditReceipt>;
  /** Removes an element and, for a group, every object in it. */
  deleteElement(
    fields: Fields<DeleteElementOperation>,
    options?: ApplyOptions,
  ): Promise<EditReceipt>;
  /** Adds a blank page. */
  insertPage(
    fields: Fields<InsertPageOperation>,
    options?: ApplyOptions,
  ): Promise<EditReceipt>;
  /** Removes a page; the last page cannot be removed. */
  deletePage(
    fields: Fields<DeletePageOperation>,
    options?: ApplyOptions,
  ): Promise<EditReceipt>;
  /** Reorders pages. */
  movePage(
    fields: Fields<MovePageOperation>,
    options?: ApplyOptions,
  ): Promise<EditReceipt>;
  /** Sets a page's rotation, or turns it from the one it has. */
  rotatePage(
    fields: Fields<RotatePageOperation>,
    options?: ApplyOptions,
  ): Promise<EditReceipt>;
  /** Draws a rectangle, ellipse or line as a path object. */
  insertShape(
    fields: Fields<InsertShapeOperation>,
    options?: ApplyOptions,
  ): Promise<EditReceipt>;
  /** Changes or removes a shape's stroke and fill. */
  setShapeStyle(
    fields: Fields<SetShapeStyleOperation>,
    options?: ApplyOptions,
  ): Promise<EditReceipt>;
  /** Places a PNG or JPEG inside `rect`; JPEG data is embedded as it is. */
  insertImage(
    fields: Fields<InsertImageOperation>,
    options?: ApplyOptions,
  ): Promise<EditReceipt>;
  /** Draws a grid with wrapped cell text; rows grow to fit their tallest cell. */
  insertTable(
    fields: Fields<InsertTableOperation>,
    options?: ApplyOptions,
  ): Promise<EditReceipt>;
  /** Changes one cell's text and lays the table out again. */
  setTableCell(
    fields: Fields<SetTableCellOperation>,
    options?: ApplyOptions,
  ): Promise<EditReceipt>;
}
