import assert from "node:assert/strict";
import { existsSync, readdirSync, readFileSync } from "node:fs";
import { describe, it } from "node:test";
import { fileURLToPath, pathToFileURL } from "node:url";

import {
  compactPdf,
  PdfCompactionError,
} from "../src/edit/pdf/engine/compact.js";
import { PdfEditDocument } from "../src/edit/pdf/engine/document.js";
import type { PdfOperation } from "../src/index.js";
import { extractPageText, fixturePdfium } from "./fixtures/pdf-builder.js";

/*
 * The compaction pass of a full PDF save: a lexer over PDFium's output that
 * keeps the objects reachable from the trailer. Checked on hand-written
 * files with the constructs string searching would trip over, on every
 * corpus PDF (two of which use object and cross-reference streams, which
 * PDFium's full save turns into classic objects), and reopened with PDFium
 * and PDF.js afterwards.
 */

// Tests run from `.test-dist/test/`; resolve against the package directory.
const PACKAGE = pathToFileURL(`${process.cwd()}/`);
const CORPUS = new URL("../../.cache/corpus/", PACKAGE);
const op = <T extends PdfOperation>(operation: T): T => operation;

/** A classic PDF from object bodies; `trailer` names the catalog. */
function classicPdf(objects: readonly string[], extraTrailer = ""): Uint8Array {
  let out = "%PDF-1.7\n%âãÏÓ\n";
  const offsets: number[] = [];
  objects.forEach((body, index) => {
    offsets.push(out.length);
    out += `${index + 1} 0 obj\n${body}\nendobj\n`;
  });
  const xref = out.length;
  out += `xref\n0 ${objects.length + 1}\n0000000000 65535 f \n`;
  for (const offset of offsets)
    out += `${String(offset).padStart(10, "0")} 00000 n \n`;
  out += `trailer\n<< /Size ${objects.length + 1} /Root 1 0 R${extraTrailer} >>\nstartxref\n${xref}\n%%EOF\n`;
  return Uint8Array.from(out, (character) => character.charCodeAt(0) & 0xff);
}

const latin1 = (bytes: Uint8Array): string =>
  Array.from(bytes, (byte) => String.fromCharCode(byte)).join("");

async function pageTexts(bytes: Uint8Array): Promise<string[]> {
  const pdfium = await fixturePdfium();
  const document = pdfium.openDocument(bytes);
  const count = pdfium.lib.FPDF_GetPageCount(document.handle);
  document.close();
  const texts: string[] = [];
  for (let index = 0; index < count; index += 1)
    texts.push(await extractPageText(bytes, index));
  return texts;
}

/** PDF.js in Node: page count and the text of the first page. */
async function pdfjsSummary(bytes: Uint8Array) {
  const pdfjs = await import("pdfjs-dist/legacy/build/pdf.mjs");
  pdfjs.GlobalWorkerOptions.workerSrc = new URL(
    "../../node_modules/pdfjs-dist/legacy/build/pdf.worker.mjs",
    PACKAGE,
  ).href;
  const task = pdfjs.getDocument({
    data: bytes.slice(),
    disableFontFace: true,
    useSystemFonts: false,
    // The Node build reads these from the file system, as paths.
    standardFontDataUrl: fileURLToPath(
      new URL("../../node_modules/pdfjs-dist/standard_fonts/", PACKAGE),
    ),
    cMapUrl: fileURLToPath(
      new URL("../../node_modules/pdfjs-dist/cmaps/", PACKAGE),
    ),
    cMapPacked: true,
  });
  try {
    const document = await task.promise;
    const page = await document.getPage(1);
    const content = await page.getTextContent();
    return {
      pages: document.numPages,
      text: content.items
        .map((item) => ("str" in item ? item.str : ""))
        .join(""),
    };
  } finally {
    await task.destroy();
  }
}

describe("pdf compaction", () => {
  const content = "BT /F1 12 Tf 72 700 Td (Kept) Tj ET";

  it("keeps reachable objects through strings, comments, names and split references", async () => {
    const bytes = classicPdf([
      // 1: catalog; the reference to the page tree is split over a line.
      "<< /Type /Catalog /Pages\n2\n0 R /Upstream (a string with endobj and stream inside) >>",
      "<< /Type /Pages /Kids [ 3 0 R ] /Count 1 >>",
      // 3: page; a comment that would confuse a search for "endobj".
      "<< /Type /Page /Parent 2 0 R /MediaBox [0 0 612 792] % endobj stream\n/Resources << /Font << /F1 4 0 R >> >> /Contents 5 0 R >>",
      "<< /Type /Font /Subtype /Type1 /BaseFont /Helvetica /Encoding /WinAnsiEncoding >>",
      // 5: content stream with an indirect length that is defined later.
      `<< /Length 6 0 R >>\nstream\n${content}\nendstream`,
      String(content.length),
      // 7: unreachable: an orphaned stream whose data says "endstream".
      "<< /Length 20 >>\nstream\nORPHAN endstream dat\nendstream",
      // 8: unreachable dictionary with a hex string.
      "<< /Note <656e646f626a> >>",
    ]);
    const compacted = compactPdf(bytes);
    const text = latin1(compacted);
    assert.ok(
      text.includes("/Upstream (a string"),
      "the catalog survived intact",
    );
    assert.ok(text.includes("(Kept)"), "the content stream survived");
    assert.equal(text.includes("ORPHAN"), false, "the orphan stream is gone");
    assert.equal(
      text.includes("656e646f626a"),
      false,
      "the orphan dict is gone",
    );
    assert.ok(text.includes("6 0 obj"), "the indirect length object is kept");
    assert.deepEqual(await pageTexts(compacted), ["Kept"]);
    assert.deepEqual(await pdfjsSummary(compacted), { pages: 1, text: "Kept" });
    // Idempotent: compacting the output changes nothing.
    assert.deepEqual(compactPdf(compacted), compacted);
  });

  it("refuses files it cannot read safely", () => {
    assert.throws(
      () =>
        compactPdf(
          new TextEncoder().encode("%PDF-1.7\n1 0 obj\n<< >>\nendobj\n"),
        ),
      (error: unknown) =>
        error instanceof PdfCompactionError &&
        error.details?.reason === "pdf-compaction",
    );
    assert.throws(
      () =>
        compactPdf(
          classicPdf([
            "<< /Type /Catalog >>",
            "<< /Length 99 >>\nstream\nshort",
          ]),
        ),
      PdfCompactionError,
    );
  });

  it("round-trips every corpus PDF through a full save and reopens it in PDFium and PDF.js", async (t) => {
    if (!existsSync(CORPUS)) {
      t.skip("corpus not fetched; run npm run corpus:fetch");
      return;
    }
    const pdfium = await fixturePdfium();
    const files = readdirSync(CORPUS).filter((name) => name.endsWith(".pdf"));
    assert.ok(files.length >= 7, "the public corpus has PDFs");
    for (const name of files) {
      const original = new Uint8Array(readFileSync(new URL(name, CORPUS)));
      const before = await pageTexts(original);
      const model = new PdfEditDocument(pdfium, original);
      try {
        // The box has to lie on the page: the corpus holds a 200×50 pt one.
        const page = model.pageLayout(0)!;
        model.apply([
          op({
            op: "insertTextBox",
            pageIndex: 0,
            rect: {
              x: Math.min(36, page.width / 10),
              y: Math.min(36, page.height / 10),
              width: Math.min(200, page.width * 0.8),
              height: Math.min(30, page.height * 0.8),
            },
            text: "Compacted",
          }),
        ]);
        const shownStart = performance.now();
        const shown = model.materialize("show");
        const shownMs = performance.now() - shownStart;
        const fullStart = performance.now();
        const full = model.materialize("save", "full");
        const fullMs = performance.now() - fullStart;
        console.log(
          `compaction ${name}: ${original.length} B, show ${shownMs.toFixed(1)} ms (${shown.length} B), full+compact ${fullMs.toFixed(1)} ms (${full.length} B)`,
        );
        const after = await pageTexts(full);
        assert.equal(after.length, before.length, `${name}: page count`);
        assert.ok(after[0]!.includes("Compacted"), `${name}: the box is there`);
        for (let index = 1; index < before.length; index += 1)
          assert.equal(after[index], before[index], `${name}: page ${index}`);
        const viaPdfjs = await pdfjsSummary(full);
        assert.equal(
          viaPdfjs.pages,
          before.length,
          `${name}: PDF.js page count`,
        );
        assert.ok(viaPdfjs.text.includes("Compacted"), `${name}: PDF.js text`);
        assert.ok(
          full.length <= shown.length * 1.05,
          `${name}: compact output`,
        );
      } finally {
        model.dispose();
      }
    }
  });

  it("stays within budget on a heavy image-only file", async () => {
    // Two hundred pages of raw RGB image data: the shape of a scanned PDF,
    // about 50 MB, without an encoder in the test.
    const pages = 200;
    const sample = 300 * 300 * 3;
    const objects: string[] = [
      "<< /Type /Catalog /Pages 2 0 R >>",
      `<< /Type /Pages /Count ${pages} /Kids [${Array.from({ length: pages }, (_, index) => `${3 + index * 3} 0 R`).join(" ")}] >>`,
    ];
    // Pseudo-random bytes so the streams do not deflate to nothing.
    let state = 0x9e3779b9;
    const noise = Array.from({ length: sample }, () => {
      state = (Math.imul(state, 1_664_525) + 1_013_904_223) >>> 0;
      return String.fromCharCode(state >>> 24);
    }).join("");
    for (let index = 0; index < pages; index += 1) {
      const page = 3 + index * 3;
      objects.push(
        `<< /Type /Page /Parent 2 0 R /MediaBox [0 0 612 792] /Resources << /XObject << /Im ${page + 1} 0 R >> >> /Contents ${page + 2} 0 R >>`,
        `<< /Type /XObject /Subtype /Image /Width 300 /Height 300 /ColorSpace /DeviceRGB /BitsPerComponent 8 /Length ${sample} >>\nstream\n${noise}\nendstream`,
        "<< /Length 35 >>\nstream\nq 612 0 0 792 0 0 cm /Im Do Q      \nendstream",
      );
    }
    const original = classicPdf(objects);
    assert.ok(original.length > 50_000_000, `${original.length} bytes`);
    const pdfium = await fixturePdfium();
    const openStart = performance.now();
    const model = new PdfEditDocument(pdfium, original);
    try {
      const openMs = performance.now() - openStart;
      const applyStart = performance.now();
      model.apply([
        op({
          op: "insertTextBox",
          pageIndex: 100,
          rect: { x: 36, y: 36, width: 200, height: 30 },
          text: "Heavy",
        }),
      ]);
      const applyMs = performance.now() - applyStart;
      const shownStart = performance.now();
      const shown = model.materialize("show");
      const shownMs = performance.now() - shownStart;
      const fullStart = performance.now();
      const full = model.materialize("save", "full");
      const fullMs = performance.now() - fullStart;
      console.log(
        `heavy ${(original.length / 1048576).toFixed(1)} MB: open ${openMs.toFixed(0)} ms, apply ${applyMs.toFixed(0)} ms, show ${shownMs.toFixed(0)} ms (${(shown.length / 1048576).toFixed(1)} MB), full+compact ${fullMs.toFixed(0)} ms (${(full.length / 1048576).toFixed(1)} MB)`,
      );
      assert.equal(await extractPageText(full, 100), "Heavy");
      assert.ok(
        fullMs < 10_000,
        "a full save with compaction stays under 10 s",
      );
    } finally {
      model.dispose();
    }
  });
});
