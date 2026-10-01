import { readFileSync } from "node:fs";
import { createRequire } from "node:module";

import { Pdfium } from "../../src/edit/pdf/engine/pdfium.js";

/*
 * Deterministic PDF fixtures built with PDFium itself, so no binary files are
 * committed. Every page shows its text in Helvetica at the given position.
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

export interface FixturePage {
  readonly text?: string;
  readonly width?: number;
  readonly height?: number;
  readonly rotation?: 0 | 1 | 2 | 3;
  readonly x?: number;
  readonly y?: number;
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
      if (page.text) {
        const font = lib.FPDFText_LoadStandardFont(
          document.handle,
          "Helvetica",
        );
        const object = lib.FPDFPageObj_CreateTextObj(document.handle, font, 12);
        const wide = pdfium.writeWideString(page.text);
        try {
          lib.FPDFText_SetText(object, wide);
        } finally {
          pdfium.free(wide);
        }
        lib.FPDFPageObj_Transform(
          object,
          1,
          0,
          0,
          1,
          page.x ?? 72,
          page.y ?? 700,
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
