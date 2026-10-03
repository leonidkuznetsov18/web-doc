import { readFileSync } from "node:fs";
import { createRequire } from "node:module";

import { MARK_NAME, MARK_PARAM } from "../../src/edit/pdf/engine/elements.js";
import { Pdfium } from "../../src/edit/pdf/engine/pdfium.js";
import { tinyJpeg } from "./tiny-jpeg.js";

/*
 * Deterministic PDF fixtures built with PDFium itself, so no binary files are
 * committed. Pages are described in PDF user space (points, y up).
 */

const wasm = readFileSync(
  createRequire(import.meta.url).resolve("@embedpdf/pdfium/pdfium.wasm"),
);

let shared: Promise<Pdfium> | undefined;

/** One PDFium instance for all fixture building in a test process. */
export function fixturePdfium(): Promise<Pdfium> {
  shared ??= Pdfium.load(wasm);
  return shared;
}

export interface FixtureText {
  readonly text: string;
  readonly x?: number;
  readonly y?: number;
  readonly fontSize?: number;
  readonly font?: "Helvetica" | "Helvetica-Bold" | "Times-Italic";
  /**
   * The linear part of the text matrix, `[a, b, c, d]`: producers often
   * write `1 Tf` and carry the size here, scaled or turned.
   */
  readonly matrix?: readonly [number, number, number, number];
  readonly color?: readonly [number, number, number];
  /** Tag the object with a WebDoc mark carrying these parameters. */
  readonly mark?: Readonly<Record<string, unknown>>;
  /** Tag the object with a mark whose parameter is not valid JSON. */
  readonly brokenMark?: boolean;
}

export interface FixturePage {
  readonly text?: string;
  readonly texts?: readonly FixtureText[];
  readonly width?: number;
  readonly height?: number;
  readonly rotation?: 0 | 1 | 2 | 3;
  /** User-space crop box: left, bottom, right, top. */
  readonly cropBox?: readonly [number, number, number, number];
  readonly image?: {
    readonly x: number;
    readonly y: number;
    readonly width: number;
    readonly height: number;
  };
  readonly rect?: {
    readonly x: number;
    readonly y: number;
    readonly width: number;
    readonly height: number;
    readonly fill?: readonly [number, number, number];
    readonly stroke?: readonly [number, number, number];
    readonly strokeWidth?: number;
  };
}

export async function buildPdf(
  pages: readonly (FixturePage | string)[],
): Promise<Uint8Array> {
  const pdfium = await fixturePdfium();
  const { lib } = pdfium;
  const document = pdfium.createDocument();
  try {
    pages.forEach((entry, index) => {
      const page = typeof entry === "string" ? { text: entry } : entry;
      const handle = lib.FPDFPage_New(
        document.handle,
        index,
        page.width ?? 612,
        page.height ?? 792,
      );
      if (page.cropBox) lib.FPDFPage_SetCropBox(handle, ...page.cropBox);
      const texts = [
        ...(page.text ? [{ text: page.text }] : []),
        ...(page.texts ?? []),
      ];
      for (const text of texts) addText(pdfium, document.handle, handle, text);
      if (page.rect) {
        const { rect } = page;
        const path = lib.FPDFPageObj_CreateNewRect(
          rect.x,
          rect.y,
          rect.width,
          rect.height,
        );
        if (rect.fill) lib.FPDFPageObj_SetFillColor(path, ...rect.fill, 255);
        if (rect.stroke) {
          lib.FPDFPageObj_SetStrokeColor(path, ...rect.stroke, 255);
          lib.FPDFPageObj_SetStrokeWidth(path, rect.strokeWidth ?? 1);
        }
        lib.FPDFPath_SetDrawMode(path, rect.fill ? 1 : 0, Boolean(rect.stroke));
        lib.FPDFPage_InsertObject(handle, path);
      }
      if (page.image) {
        const { image } = page;
        const object = lib.FPDFPageObj_NewImageObj(document.handle);
        pdfium.withFileAccess(tinyJpeg(), (fileAccess) =>
          lib.FPDFImageObj_LoadJpegFileInline(0, 0, object, fileAccess),
        );
        lib.FPDFImageObj_SetMatrix(
          object,
          image.width,
          0,
          0,
          image.height,
          image.x,
          image.y,
        );
        lib.FPDFPage_InsertObject(handle, object);
      }
      if (page.rotation) lib.FPDFPage_SetRotation(handle, page.rotation);
      lib.FPDFPage_GenerateContent(handle);
      lib.FPDF_ClosePage(handle);
    });
    return document.save("full");
  } finally {
    document.close();
  }
}

function addText(
  pdfium: Pdfium,
  document: number,
  page: number,
  text: FixtureText,
): void {
  const { lib } = pdfium;
  const font = lib.FPDFText_LoadStandardFont(
    document,
    text.font ?? "Helvetica",
  );
  const object = lib.FPDFPageObj_CreateTextObj(
    document,
    font,
    text.fontSize ?? 12,
  );
  const wide = pdfium.writeWideString(text.text);
  try {
    lib.FPDFText_SetText(object, wide);
  } finally {
    pdfium.free(wide);
  }
  if (text.color) lib.FPDFPageObj_SetFillColor(object, ...text.color, 255);
  const [a, b, c, d] = text.matrix ?? [1, 0, 0, 1];
  lib.FPDFPageObj_Transform(object, a, b, c, d, text.x ?? 72, text.y ?? 700);
  if (text.mark || text.brokenMark) {
    const mark = lib.FPDFPageObj_AddMark(object, MARK_NAME);
    lib.FPDFPageObjMark_SetStringParam(
      document,
      object,
      mark,
      MARK_PARAM,
      text.brokenMark ? "{not json" : JSON.stringify(text.mark),
    );
  }
  lib.FPDFPage_InsertObject(page, object);
}

/** Text PDFium extracts from a page, for checking saved files. */
export async function extractPageText(
  bytes: Uint8Array,
  pageIndex: number,
): Promise<string> {
  const pdfium = await fixturePdfium();
  const { lib } = pdfium;
  const document = pdfium.openDocument(bytes);
  try {
    const page = lib.FPDF_LoadPage(document.handle, pageIndex);
    const textPage = lib.FPDFText_LoadPage(page);
    try {
      const count = lib.FPDFText_CountChars(textPage);
      if (count === 0) return "";
      const bytesNeeded = (count + 1) * 2;
      const buffer = pdfium.malloc(bytesNeeded);
      try {
        lib.FPDFText_GetText(textPage, 0, count, buffer);
        return pdfium.readWideStringAt(buffer, bytesNeeded);
      } finally {
        pdfium.free(buffer);
      }
    } finally {
      lib.FPDFText_ClosePage(textPage);
      lib.FPDF_ClosePage(page);
    }
  } finally {
    document.close();
  }
}
