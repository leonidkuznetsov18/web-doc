import type { EditElement, EditSessionBase } from "../types.js";

export type PdfElementKind =
  | "text" // one text object as stored in the file (often a word or a line)
  | "image"
  | "shape" // a path object
  | "textBox" // a text box created by insertTextBox
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
  /** Text boxes only; a multiple of the font size. */
  readonly lineHeight?: number;
}

export interface PdfShapeStyle {
  readonly stroke?: { readonly color: string; readonly width: number };
  readonly fill?: { readonly color: string };
}

export interface PdfElement extends EditElement {
  readonly kind: PdfElementKind;
  /** Present for `text` and `textBox`. */
  readonly textStyle?: PdfTextStyle;
  /** Present for `shape` and `table`. */
  readonly shapeStyle?: PdfShapeStyle;
  /** Present for `table`. */
  readonly table?: { readonly rows: readonly (readonly string[])[] };
}

/**
 * The PDF operation union. Operations are added as they ship; a session with
 * no operations accepts only the empty batch.
 */
export type PdfOperation = never;

export interface PdfEditSession extends EditSessionBase<
  PdfOperation,
  PdfElement
> {
  readonly format: "pdf";
}
