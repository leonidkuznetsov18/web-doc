import assert from "node:assert/strict";
import { existsSync, readdirSync, readFileSync } from "node:fs";
import { describe, it } from "node:test";
import { fileURLToPath, pathToFileURL } from "node:url";

import {
  attachInfoDictionary,
  compactPdf,
  PdfCompactionError,
  readObjects,
} from "../src/edit/pdf/engine/compact.js";
import { PdfEditDocument } from "../src/edit/pdf/engine/document.js";
import type { PdfOperation } from "../src/index.js";
import { extractPageText, fixturePdfium } from "./fixtures/pdf-builder.js";
import { pdfSession } from "./fixtures/pdf-session.js";

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
function classicPdf(
  objects: readonly string[],
  extraTrailer = "",
  offsetBase = 0,
): Uint8Array {
  let out = "%PDF-1.7\n%âãÏÓ\n";
  const offsets: number[] = [];
  objects.forEach((body, index) => {
    offsets.push(out.length + offsetBase);
    out += `${index + 1} 0 obj\n${body}\nendobj\n`;
  });
  const xref = out.length + offsetBase;
  out += `xref\n0 ${objects.length + 1}\n0000000000 65535 f \n`;
  for (const offset of offsets)
    out += `${String(offset).padStart(10, "0")} 00000 n \n`;
  out += `trailer\n<< /Size ${objects.length + 1} /Root 1 0 R${extraTrailer} >>\nstartxref\n${xref}\n%%EOF\n`;
  return Uint8Array.from(out, (character) => character.charCodeAt(0) & 0xff);
}

const latin1 = (bytes: Uint8Array): string =>
  Array.from(bytes, (byte) => String.fromCharCode(byte)).join("");

/** A public synthetic input whose incremental PDFium save keeps an xref stream. */
function xrefStreamPdf(
  info: "linked" | "missing" = "linked",
  infoBody = "<< /Producer (Synthetic incremental test) >>",
): Uint8Array {
  const text = "BT /F1 24 Tf 72 700 Td (#1) Tj ET";
  const objects = [
    "<< /Type /Catalog /Pages 2 0 R >>",
    "<< /Type /Pages /Kids [3 0 R] /Count 1 >>",
    "<< /Type /Page /Parent 2 0 R /MediaBox [0 0 612 792] /Resources << /Font << /F1 5 0 R >> >> /Contents 4 0 R >>",
    `<< /Length ${text.length} >>\nstream\n${text}\nendstream`,
    "<< /Type /Font /Subtype /Type1 /BaseFont /Helvetica >>",
    infoBody,
  ];
  let body = "%PDF-1.7\n";
  const offsets = [0];
  objects.forEach((object, index) => {
    offsets.push(body.length);
    body += `${index + 1} 0 obj\n${object}\nendobj\n`;
  });
  offsets.push(body.length);
  const entries = new Uint8Array(offsets.length * 7);
  const view = new DataView(entries.buffer);
  offsets.forEach((offset, index) => {
    entries[index * 7] = index === 0 ? 0 : 1;
    view.setUint32(index * 7 + 1, offset);
    view.setUint16(index * 7 + 5, index === 0 ? 65535 : 0);
  });
  const xref = `${offsets.length - 1} 0 obj\n<< /Type /XRef /Size ${offsets.length} /Root 1 0 R${info === "linked" ? " /Info 6 0 R" : ""} /W [1 4 2] /Length ${entries.length} >>\nstream\n`;
  const ending = `\nendstream\nendobj\nstartxref\n${body.length}\n%%EOF\n`;
  return Uint8Array.from(body + xref + latin1(entries) + ending, (character) =>
    character.charCodeAt(0),
  );
}

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

  for (const info of ["linked", "missing"] as const)
    it(`saves an xref-stream input with ${info} Info incrementally without changing its prefix, page identities or history`, async () => {
      const original = xrefStreamPdf(info);
      const { session, end } = await pdfSession(original);
      try {
        await session.insertPage({ index: 0 });
        await session.replaceText({ target: "p0:o0", text: "#2" });
        const full = (await session.save({ mode: "full" })).bytes;
        const state = session.state;
        for (let attempt = 0; attempt < 2; attempt += 1) {
          const saved = await session.save({ mode: "incremental" });
          assert.deepEqual(saved.warnings, []);
          assert.deepEqual(saved.bytes.subarray(0, original.length), original);
          assert.deepEqual(session.state, state);
          const reopened = new PdfEditDocument(
            await fixturePdfium(),
            saved.bytes,
          );
          try {
            assert.equal(reopened.getElement("p0:o0")?.pageIndex, 1);
            assert.equal(reopened.getElement("p0:o0")?.text, "#2");
            assert.equal(await extractPageText(saved.bytes, 1), "#2");
          } finally {
            reopened.dispose();
          }
        }
        assert.deepEqual((await session.save({ mode: "full" })).bytes, full);
        await session.undo();
        assert.equal((await session.getElement("p0:o0")).item?.text, "#1");
        await session.redo();
        assert.equal((await session.getElement("p0:o0")).item?.text, "#2");
        const restored = (await session.save({ mode: "full" })).bytes;
        const reopened = new PdfEditDocument(await fixturePdfium(), restored);
        try {
          assert.equal(reopened.getElement("p0:o0")?.pageIndex, 1);
          assert.equal(reopened.getElement("p0:o0")?.text, "#2");
          assert.equal(await extractPageText(restored, 1), "#2");
        } finally {
          reopened.dispose();
        }
        assert.deepEqual(
          (await session.save({ mode: "full" })).bytes,
          restored,
        );
      } finally {
        await end();
      }
    });

  it("links orphaned Info in an xref stream without changing any existing byte or offset", async () => {
    const original = xrefStreamPdf(
      "missing",
      "<< /WebDocPageKeys (metadata) >>",
    );
    const linked = attachInfoDictionary(original, "WebDocPageKeys");
    assert.equal(
      latin1(linked).replace(" /Info 6 0 R ", ""),
      latin1(original),
      "only the Info reference is inserted; startxref, xref entries, stream data and Length stay identical",
    );
    assert.equal(await extractPageText(linked, 0), "#1");
    assert.equal(attachInfoDictionary(linked, "WebDocPageKeys"), linked);
  });

  it("only trusts linked Info in the authoritative cross-reference dictionary", () => {
    const bytes = xrefStreamPdf();
    assert.equal(attachInfoDictionary(bytes, "WebDocPageKeys"), bytes);
    const source = latin1(bytes);
    const altered = (text: string) =>
      Uint8Array.from(text, (character) => character.charCodeAt(0));
    const typeless = source.replace("/Type /XRef ", "");
    const typelessBytes = altered(typeless);
    assert.equal(
      attachInfoDictionary(typelessBytes, "WebDocPageKeys"),
      typelessBytes,
    );
    for (const fake of [
      "/Nested << /Info 6 0 R >>",
      "/Note (/Info 6 0 R)",
      "/Names [/Info 6 0 R]",
    ])
      assert.throws(
        () =>
          attachInfoDictionary(
            altered(source.replace("/Info 6 0 R", fake)),
            "WebDocPageKeys",
          ),
        PdfCompactionError,
      );
    for (const malformed of [
      source.replace("/Type /XRef", "/Type /Metadata"),
      source.replace(
        /startxref\n\d+/,
        `startxref\n${source.indexOf("6 0 obj")}`,
      ),
      source.replace(/startxref\n\d+/, "startxref\n99999999999999999999"),
      source.replace(/startxref\n\d+/, "startxref\n0"),
      typeless.replace("/W [1 4 2]", "/Widths [1 4 2]"),
      typeless.replace("/W [1 4 2]", "/W [0 0 0]"),
      typeless.replace("/W [1 4 2]", "/W [1 -1 2]"),
      typeless.replace("/W [1 4 2]", "/W [1 4 2 0]"),
      typeless.replace("/Size 8", "/Size 0"),
      typeless.replace("/Size 8", "/OtherSize 8"),
      typeless.replace("/Root 1 0 R", "/Root (1 0 R)"),
      typeless.replace("/Root 1 0 R", "/Root 0 0 R"),
      typeless.replace("/Root 1 0 R", "/Root 1 -1 R"),
      typeless.replace("/Root 1 0 R", "/OtherRoot 1 0 R"),
    ])
      assert.throws(
        () => attachInfoDictionary(altered(malformed), "WebDocPageKeys"),
        PdfCompactionError,
      );
  });

  it("attaches orphaned page metadata without adopting names in nested values or streams", () => {
    const objects = [
      "<< /Type /Catalog /Fake << /WebDocPageKeys (not Info) >> >>",
      "<< /Note (/WebDocPageKeys endobj trailer) >>",
      "<< /Length 20 >>\nstream\n/WebDocPageKeys fake!\nendstream",
      "<< /WebDocPageKeys (metadata) >>",
    ];
    const bytes = classicPdf(objects);
    const linked = attachInfoDictionary(bytes, "WebDocPageKeys");
    assert.match(latin1(linked), /\/Info 4 0 R/);
    assert.equal(
      readObjects(compactPdf(linked)).get(4)?.value?.trim(),
      "<< /WebDocPageKeys (metadata) >>",
    );
    assert.deepEqual(attachInfoDictionary(linked, "WebDocPageKeys"), linked);

    // An incremental repair must not tokenize or alter the earlier signed
    // prefix, which may use a cross-reference stream instead of a trailer.
    const prefix = Uint8Array.from("<< /Type /XRef >>\n% signed bytes\n", (c) =>
      c.charCodeAt(0),
    );
    const appended = classicPdf(objects, "", prefix.length);
    const incremental = new Uint8Array(prefix.length + appended.length);
    incremental.set(prefix);
    incremental.set(appended, prefix.length);
    const repaired = attachInfoDictionary(
      incremental,
      "WebDocPageKeys",
      prefix.length,
    );
    assert.deepEqual(repaired.subarray(0, prefix.length), prefix);
    assert.match(latin1(repaired.subarray(prefix.length)), /\/Info 4 0 R/);
  });

  it("refuses ambiguous orphaned Info metadata instead of choosing another object", () => {
    const bytes = classicPdf([
      "<< /Type /Catalog >>",
      "<< /WebDocPageKeys (first) >>",
      "<< /WebDocPageKeys (second) >>",
    ]);
    assert.throws(
      () => attachInfoDictionary(bytes, "WebDocPageKeys"),
      PdfCompactionError,
    );
  });

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
