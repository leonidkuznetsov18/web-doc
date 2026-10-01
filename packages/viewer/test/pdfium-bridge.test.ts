import assert from "node:assert/strict";
import { readFileSync } from "node:fs";
import { createRequire } from "node:module";
import { before, describe, it } from "node:test";

import { Pdfium, type PdfiumDocument } from "../src/edit/pdf/engine/pdfium.js";
import { ViewerError } from "../src/index.js";
import { tinyJpeg } from "./fixtures/tiny-jpeg.js";

const wasm = readFileSync(
  createRequire(import.meta.url).resolve("@embedpdf/pdfium/pdfium.wasm"),
);

/** Object types reported by FPDFPageObj_GetType. */
const TEXT_OBJECT = 1;
const IMAGE_OBJECT = 3;

function addText(
  pdfium: Pdfium,
  document: PdfiumDocument,
  page: number,
  text: string,
  x: number,
  y: number,
): number {
  const { lib } = pdfium;
  const font = lib.FPDFText_LoadStandardFont(document.handle, "Helvetica");
  const object = lib.FPDFPageObj_CreateTextObj(document.handle, font, 12);
  const wide = pdfium.writeWideString(text);
  try {
    assert.equal(lib.FPDFText_SetText(object, wide), true);
  } finally {
    pdfium.free(wide);
  }
  lib.FPDFPageObj_Transform(object, 1, 0, 0, 1, x, y);
  lib.FPDFPage_InsertObject(page, object);
  return object;
}

function pageText(
  pdfium: Pdfium,
  document: PdfiumDocument,
  index: number,
): string {
  const { lib } = pdfium;
  const page = lib.FPDF_LoadPage(document.handle, index);
  const textPage = lib.FPDFText_LoadPage(page);
  try {
    const count = lib.FPDFText_CountChars(textPage);
    if (count === 0) return "";
    const bytes = (count + 1) * 2;
    const buffer = pdfium.malloc(bytes);
    try {
      lib.FPDFText_GetText(textPage, 0, count, buffer);
      return pdfium.readWideStringAt(buffer, bytes);
    } finally {
      pdfium.free(buffer);
    }
  } finally {
    lib.FPDFText_ClosePage(textPage);
    lib.FPDF_ClosePage(page);
  }
}

/** A three-page document saved without incremental history. */
function createOriginal(pdfium: Pdfium): Uint8Array {
  const document = pdfium.createDocument();
  try {
    for (let index = 0; index < 3; index += 1) {
      const page = pdfium.lib.FPDFPage_New(document.handle, index, 612, 792);
      addText(pdfium, document, page, `Page ${index + 1}`, 72, 700);
      assert.equal(pdfium.lib.FPDFPage_GenerateContent(page), true);
      pdfium.lib.FPDF_ClosePage(page);
    }
    return document.save("full");
  } finally {
    document.close();
  }
}

function appendText(pdfium: Pdfium, original: Uint8Array): Uint8Array {
  const document = pdfium.openDocument(original);
  try {
    const page = pdfium.lib.FPDF_LoadPage(document.handle, 0);
    addText(pdfium, document, page, "Added", 72, 600);
    assert.equal(pdfium.lib.FPDFPage_GenerateContent(page), true);
    pdfium.lib.FPDF_ClosePage(page);
    return document.save("incremental");
  } finally {
    document.close();
  }
}

function startsWith(bytes: Uint8Array, prefix: Uint8Array): boolean {
  return (
    bytes.byteLength > prefix.byteLength &&
    prefix.every((value, index) => bytes[index] === value)
  );
}

describe("PDFium bridge", () => {
  let pdfium: Pdfium;
  let original: Uint8Array;

  before(async () => {
    pdfium = await Pdfium.load(wasm);
    original = createOriginal(pdfium);
  });

  it("saves an incremental update that starts with the original bytes", () => {
    const edited = appendText(pdfium, original);
    assert.equal(startsWith(edited, original), true);

    const reopened = pdfium.openDocument(edited);
    try {
      assert.equal(pageText(pdfium, reopened, 0), "Page 1\r\nAdded");
      assert.equal(pageText(pdfium, reopened, 1), "Page 2");
    } finally {
      reopened.close();
    }
  });

  it("produces identical bytes for the same edit across instances and saves", async () => {
    const first = appendText(pdfium, original);
    // A second instance has a fresh heap; extra allocations in the first one
    // move every later address. Neither may leak into the output.
    const other = await Pdfium.load(wasm);
    const scratch = pdfium.malloc(4096);
    const second = appendText(other, original);
    const third = appendText(pdfium, original);
    pdfium.free(scratch);
    assert.deepEqual(second, first);
    assert.deepEqual(third, first);
  });

  it("keeps WebDoc marked-content parameters across save and reopen", () => {
    const { lib } = pdfium;
    const parameters = JSON.stringify({ kind: "textBox", text: "Привіт — ok" });
    const document = pdfium.openDocument(original);
    let edited: Uint8Array;
    try {
      const page = lib.FPDF_LoadPage(document.handle, 0);
      const object = addText(pdfium, document, page, "Tagged", 72, 650);
      const mark = lib.FPDFPageObj_AddMark(object, "WebDoc");
      assert.notEqual(mark, 0);
      assert.equal(
        lib.FPDFPageObjMark_SetStringParam(
          document.handle,
          object,
          mark,
          "webdoc",
          parameters,
        ),
        true,
      );
      assert.equal(lib.FPDFPage_GenerateContent(page), true);
      lib.FPDF_ClosePage(page);
      edited = document.save("incremental");
    } finally {
      document.close();
    }

    const reopened = pdfium.openDocument(edited);
    try {
      const page = lib.FPDF_LoadPage(reopened.handle, 0);
      const found: string[] = [];
      for (let index = 0; index < lib.FPDFPage_CountObjects(page); index += 1) {
        const object = lib.FPDFPage_GetObject(page, index);
        for (let m = 0; m < lib.FPDFPageObj_CountMarks(object); m += 1) {
          const mark = lib.FPDFPageObj_GetMark(object, m);
          const name = pdfium.readWideStringOut((buffer, bytes, out) =>
            lib.FPDFPageObjMark_GetName(mark, buffer, bytes, out),
          );
          if (name !== "WebDoc") continue;
          found.push(
            pdfium.readWideStringOut((buffer, bytes, out) =>
              lib.FPDFPageObjMark_GetParamStringValue(
                mark,
                "webdoc",
                buffer,
                bytes,
                out,
              ),
            ),
          );
        }
      }
      lib.FPDF_ClosePage(page);
      assert.deepEqual(found, [parameters]);
    } finally {
      reopened.close();
    }
  });

  it("inserts, moves, rotates and deletes pages in an incremental save", () => {
    const { lib } = pdfium;
    const document = pdfium.openDocument(original);
    let edited: Uint8Array;
    try {
      const inserted = lib.FPDFPage_New(document.handle, 1, 300, 400);
      assert.equal(lib.FPDFPage_GenerateContent(inserted), true);
      lib.FPDF_ClosePage(inserted);
      // [1, new, 2, 3] → move "Page 3" to the front → [3, 1, new, 2]
      const indices = pdfium.writeInt32Array([3]);
      try {
        assert.equal(lib.FPDF_MovePages(document.handle, indices, 1, 0), true);
      } finally {
        pdfium.free(indices);
      }
      const rotated = lib.FPDF_LoadPage(document.handle, 2);
      lib.FPDFPage_SetRotation(rotated, 1);
      lib.FPDF_ClosePage(rotated);
      lib.FPDFPage_Delete(document.handle, 3);
      edited = document.save("incremental");
    } finally {
      document.close();
    }
    assert.equal(startsWith(edited, original), true);

    const reopened = pdfium.openDocument(edited);
    try {
      assert.equal(lib.FPDF_GetPageCount(reopened.handle), 3);
      assert.equal(pageText(pdfium, reopened, 0), "Page 3");
      assert.equal(pageText(pdfium, reopened, 1), "Page 1");
      const page = lib.FPDF_LoadPage(reopened.handle, 2);
      // A quarter turn reports the displayed size, so width and height swap.
      assert.deepEqual(
        [
          lib.FPDFPage_GetRotation(page),
          lib.FPDF_GetPageWidthF(page),
          lib.FPDF_GetPageHeightF(page),
        ],
        [1, 400, 300],
      );
      lib.FPDF_ClosePage(page);
    } finally {
      reopened.close();
    }
  });

  it("loads JPEG data inline through a WASM table callback", () => {
    const { lib } = pdfium;
    const document = pdfium.openDocument(original);
    let edited: Uint8Array;
    try {
      const page = lib.FPDF_LoadPage(document.handle, 0);
      const image = lib.FPDFPageObj_NewImageObj(document.handle);
      const loaded = pdfium.withFileAccess(tinyJpeg(), (fileAccess) =>
        lib.FPDFImageObj_LoadJpegFileInline(0, 0, image, fileAccess),
      );
      assert.equal(loaded, true);
      assert.equal(
        lib.FPDFImageObj_SetMatrix(image, 160, 0, 0, 80, 300, 500),
        true,
      );
      lib.FPDFPage_InsertObject(page, image);
      assert.equal(lib.FPDFPage_GenerateContent(page), true);
      lib.FPDF_ClosePage(page);
      edited = document.save("incremental");
    } finally {
      document.close();
    }

    const reopened = pdfium.openDocument(edited);
    try {
      const page = lib.FPDF_LoadPage(reopened.handle, 0);
      const types = Array.from(
        { length: lib.FPDFPage_CountObjects(page) },
        (_, index) =>
          lib.FPDFPageObj_GetType(lib.FPDFPage_GetObject(page, index)),
      );
      lib.FPDF_ClosePage(page);
      assert.deepEqual(types, [TEXT_OBJECT, IMAGE_OBJECT]);
    } finally {
      reopened.close();
    }
  });

  it("rejects malformed input with a typed error and frees nothing twice", () => {
    assert.throws(
      () => pdfium.openDocument(new TextEncoder().encode("not a pdf")),
      (error: unknown) =>
        error instanceof ViewerError && error.code === "invalid-file",
    );
    const document = pdfium.openDocument(original);
    document.close();
    document.close();
    assert.throws(
      () => document.save("full"),
      (error: unknown) =>
        error instanceof ViewerError && error.code === "lifecycle-error",
    );
  });
});
