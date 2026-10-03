import type {
  ApplyOptions,
  BinaryData,
  EditColor,
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
  TextRange,
} from "../types.js";

/*
 * The PPTX editing contract: elements of a slide, slides and layouts, and
 * the operations that change them. Geometry is slide space (CSS pixels at
 * 96 dpi, EMU / 9525); font sizes and line widths are points.
 */

export type PptxElementKind =
  | "shape" // p:sp — a text box, a placeholder, an auto shape, WordArt
  | "image" // p:pic
  | "table" // p:graphicFrame holding a:tbl
  | "connector" // p:cxnSp
  | "group" // p:grpSp; its children carry parentId
  | "other"; // charts, diagrams, OLE objects, media frames

export type PptxTextAlign = "left" | "center" | "right" | "justify";

/** The style of the first run with text, resolved through the placeholder chain and the theme. */
export interface PptxTextStyle {
  /** Theme fonts resolved: "+mn-lt" becomes the minor Latin face. */
  readonly fontFamily: string;
  /** Points. */
  readonly fontSize: number;
  readonly bold: boolean;
  readonly italic: boolean;
  readonly underline: boolean;
  /** `#RRGGBB`, `{ theme, mods }` when the file uses a scheme colour, "auto" when unresolvable. */
  readonly color: EditColor;
  readonly align: PptxTextAlign;
}

export interface PptxLineStyle {
  readonly color: EditColor | "none";
  /** Points. */
  readonly width: number;
}

export interface PptxShapeStyle {
  /** "none" for a:noFill; absent when inherited from the shape style or the theme. */
  readonly fill?: EditColor | "none";
  readonly line?: PptxLineStyle;
}

export interface PptxPlaceholder {
  /** `p:ph/@type`, "body" when the file leaves it out. */
  readonly type: string;
  readonly idx?: number;
}

export interface PptxElement extends EditElement {
  readonly kind: PptxElementKind;
  /** `p:cNvPr/@name`, as PowerPoint shows it in the selection pane. */
  readonly name: string;
  readonly placeholder?: PptxPlaceholder;
  /** Present for a shape with a text body. */
  readonly textStyle?: PptxTextStyle;
  /** Present for shapes and connectors. */
  readonly shapeStyle?: PptxShapeStyle;
  /** Present for a table. */
  readonly table?: { readonly rows: readonly (readonly string[])[] };
  /** `p:cNvPr/@hidden`: listed and editable, but not drawn by the renderer. */
  readonly hidden?: boolean;
}

export interface PptxSlideInfo {
  readonly pageIndex: number;
  /** Stable for the session: the slide part's number, as in element ids ("sld3"). */
  readonly key: string;
  /** A `PptxLayoutInfo` id. */
  readonly layout: string;
  readonly hidden: boolean;
}

export interface PptxLayoutInfo {
  /** "layout2": from the layout part's number. */
  readonly id: string;
  /** `p:cSld/@name`, for example "Title and Content". */
  readonly name: string;
  /** `p:sldLayout/@type` when present. */
  readonly type?: string;
  /** "master1": from the master part's number. */
  readonly master: string;
}

export interface PptxTextStyleChange {
  /** Written as a:latin/@typeface; theme names ("+mn-lt") pass through. */
  readonly fontFamily?: string;
  /** 1–400 points. */
  readonly fontSize?: number;
  readonly bold?: boolean;
  readonly italic?: boolean;
  readonly underline?: boolean;
  /** `#RRGGBB` writes a:srgbClr; `{ theme, mods }` writes a:schemeClr with modifiers. */
  readonly color?: EditColor;
  /** a:pPr/@algn on every paragraph the range touches. */
  readonly align?: PptxTextAlign;
}

export interface PptxTableStyle {
  /** Header-row flag of the table; default true. */
  readonly firstRow?: boolean;
  /** Banded-rows flag of the table; default true. */
  readonly bandRow?: boolean;
}

export interface PptxLineChange {
  readonly color: EditColor | "none";
  /** Points; kept when absent. */
  readonly width?: number;
}

export interface PptxReplaceTextOperation {
  readonly op: "replaceText";
  /** A `shape` element with a text body. */
  readonly target: string;
  readonly text: string;
  /** Both ends on the target; absent, the whole text; collapsed, an insertion. */
  readonly range?: TextRange;
}

export interface PptxSetTextStyleOperation {
  readonly op: "setTextStyle";
  readonly target: string;
  readonly range?: TextRange;
  readonly style: PptxTextStyleChange;
}

export interface PptxSetShapeStyleOperation {
  readonly op: "setShapeStyle";
  /** A `shape` or `connector` element. */
  readonly target: string;
  /** A colour, "none" for no fill, `null` to drop the explicit fill, absent to keep it. */
  readonly fill?: EditColor | "none" | null;
  /** A line, "none" for no line, `null` to drop the explicit line, absent to keep it. */
  readonly line?: PptxLineChange | "none" | null;
}

export interface PptxMoveElementOperation {
  readonly op: "moveElement";
  readonly target: string;
  /** New top-left corner of the element's bounds, in slide space. */
  readonly to?: PagePoint;
  /** Offset in slide space. Exactly one of `to` and `by` is given. */
  readonly by?: { readonly dx: number; readonly dy: number };
}

export interface PptxResizeElementOperation {
  readonly op: "resizeElement";
  readonly target: string;
  /** New bounds in slide space; the rotation is kept. */
  readonly rect: PageRect;
}

export interface PptxDeleteElementOperation {
  readonly op: "deleteElement";
  readonly target: string;
}

export interface PptxInsertTextBoxOperation {
  readonly op: "insertTextBox";
  readonly pageIndex: number;
  readonly rect: PageRect;
  readonly text: string;
  readonly style?: PptxTextStyleChange;
}

export interface PptxInsertImageOperation {
  readonly op: "insertImage";
  readonly pageIndex: number;
  readonly rect: PageRect;
  /** PNG or JPEG bytes; base64 over pure-JSON transports. */
  readonly data: BinaryData;
  readonly mimeType: "image/png" | "image/jpeg";
}

export interface PptxInsertTableOperation {
  readonly op: "insertTable";
  readonly pageIndex: number;
  readonly rect: PageRect;
  /** Cell text by row; every row has the same number of cells. 1–100 rows, 1–20 columns. */
  readonly rows: readonly (readonly string[])[];
  /** Relative column weights, one per column; equal when omitted. */
  readonly columnWidths?: readonly number[];
  readonly style?: PptxTableStyle;
}

export interface PptxSetTableCellOperation {
  readonly op: "setTableCell";
  /** A `table` element. */
  readonly target: string;
  readonly row: number;
  readonly column: number;
  /** The cell's new text; empty clears the cell. */
  readonly text: string;
}

export interface PptxInsertSlideOperation {
  readonly op: "insertSlide";
  /** Position of the new slide, 0 to the current slide count. */
  readonly index: number;
  /** A layout id; default: the layout of the slide before the position, else the first layout. */
  readonly layout?: string;
}

export interface PptxDuplicateSlideOperation {
  readonly op: "duplicateSlide";
  readonly pageIndex: number;
  /** Position of the copy; default: right after the source. */
  readonly index?: number;
}

export interface PptxDeleteSlideOperation {
  readonly op: "deleteSlide";
  readonly pageIndex: number;
}

export interface PptxMoveSlideOperation {
  readonly op: "moveSlide";
  readonly from: number;
  /** The slide's index after the move. */
  readonly to: number;
}

/** The PPTX operation union; operations are added as they ship. */
export type PptxOperation =
  | PptxReplaceTextOperation
  | PptxSetTextStyleOperation
  | PptxSetShapeStyleOperation
  | PptxMoveElementOperation
  | PptxResizeElementOperation
  | PptxDeleteElementOperation
  | PptxInsertTextBoxOperation
  | PptxInsertImageOperation
  | PptxInsertTableOperation
  | PptxSetTableCellOperation
  | PptxInsertSlideOperation
  | PptxDuplicateSlideOperation
  | PptxDeleteSlideOperation
  | PptxMoveSlideOperation;

/** PPTX saves have no fields of their own yet. */
export type PptxSaveOptions = SaveOptions;

/** An operation's fields without its `op`, as the typed methods take them. */
export type PptxFields<T extends PptxOperation> = T extends unknown
  ? Omit<T, "op">
  : never;

export interface PptxEditSession extends EditSessionBase<
  PptxOperation,
  PptxElement
> {
  readonly format: "pptx";
  save(options?: PptxSaveOptions): Promise<SavedDocument>;
  /** The slides in presentation order, with keys that survive reordering. */
  getSlides(options?: ReadOptions): Promise<ReadResult<PptxSlideInfo>>;
  /** Every layout of every master, for `insertSlide`. */
  getLayouts(options?: ReadOptions): Promise<ReadResult<PptxLayoutInfo>>;
  /**
   * The text style a range of a shape's text shows: each property every run
   * it covers shares, a property they differ on left out. Without a range,
   * the whole text; a collapsed range reads the run before it, whose style
   * text typed there takes. `undefined` for an element without text.
   */
  getTextStyle(
    fields: { readonly target: string; readonly range?: TextRange },
    options?: ReadOptions,
  ): Promise<ReadItem<Partial<PptxTextStyle>>>;
  /** Replaces the whole text of a shape, or the part a range covers. */
  replaceText(
    fields: PptxFields<PptxReplaceTextOperation>,
    options?: ApplyOptions,
  ): Promise<EditReceipt>;
  /** Changes run and paragraph properties; unspecified ones keep their bytes. */
  setTextStyle(
    fields: PptxFields<PptxSetTextStyleOperation>,
    options?: ApplyOptions,
  ): Promise<EditReceipt>;
  /** Changes or removes a shape's fill and line. */
  setShapeStyle(
    fields: PptxFields<PptxSetShapeStyleOperation>,
    options?: ApplyOptions,
  ): Promise<EditReceipt>;
  /** Moves an element to a point or by an offset, in slide space. */
  moveElement(
    fields: PptxFields<PptxMoveElementOperation>,
    options?: ApplyOptions,
  ): Promise<EditReceipt>;
  /** Gives an element new bounds; its rotation is kept. */
  resizeElement(
    fields: PptxFields<PptxResizeElementOperation>,
    options?: ApplyOptions,
  ): Promise<EditReceipt>;
  /** Removes an element and, for a group, every element in it. */
  deleteElement(
    fields: PptxFields<PptxDeleteElementOperation>,
    options?: ApplyOptions,
  ): Promise<EditReceipt>;
  /** Adds a text box; its id is in `createdIds`. */
  insertTextBox(
    fields: PptxFields<PptxInsertTextBoxOperation>,
    options?: ApplyOptions,
  ): Promise<EditReceipt>;
  /** Places a PNG or JPEG inside `rect` as a picture. */
  insertImage(
    fields: PptxFields<PptxInsertImageOperation>,
    options?: ApplyOptions,
  ): Promise<EditReceipt>;
  /** Adds a table with the given cell text. */
  insertTable(
    fields: PptxFields<PptxInsertTableOperation>,
    options?: ApplyOptions,
  ): Promise<EditReceipt>;
  /** Changes one cell's text. */
  setTableCell(
    fields: PptxFields<PptxSetTableCellOperation>,
    options?: ApplyOptions,
  ): Promise<EditReceipt>;
  /** Adds a slide from a layout. */
  insertSlide(
    fields: PptxFields<PptxInsertSlideOperation>,
    options?: ApplyOptions,
  ): Promise<EditReceipt>;
  /** Copies a slide. */
  duplicateSlide(
    fields: PptxFields<PptxDuplicateSlideOperation>,
    options?: ApplyOptions,
  ): Promise<EditReceipt>;
  /** Removes a slide; the last slide cannot be removed. */
  deleteSlide(
    fields: PptxFields<PptxDeleteSlideOperation>,
    options?: ApplyOptions,
  ): Promise<EditReceipt>;
  /** Reorders slides. */
  moveSlide(
    fields: PptxFields<PptxMoveSlideOperation>,
    options?: ApplyOptions,
  ): Promise<EditReceipt>;
}
