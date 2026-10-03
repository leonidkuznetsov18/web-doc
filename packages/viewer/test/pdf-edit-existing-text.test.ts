import assert from "node:assert/strict";
import { readFileSync } from "node:fs";
import { before, describe, it } from "node:test";

import { createPdfEditHandler } from "../src/edit/pdf/engine/handler.js";
import {
  loadPdfEditEngine,
  type PdfEditEngineClient,
} from "../src/edit/pdf/provider.js";
import type { OperationIssue, PdfElement, PdfOperation } from "../src/index.js";
import { defaultResourceLimits } from "../src/index.js";
import { loopbackWorker } from "./fixtures/loopback-worker.js";
import {
  buildPdf,
  extractPageText,
  fixturePdfium,
} from "./fixtures/pdf-builder.js";

const ttf = new Uint8Array(
  readFileSync(
    new URL("../../fonts/noto-sans-latin-cyrillic.ttf", import.meta.url),
  ),
);
const FALLBACK_URL = "https://fonts.test/noto.ttf";
const signal = new AbortController().signal;
const op = <T extends PdfOperation>(operation: T): T => operation;

async function engineFor(original: Uint8Array): Promise<PdfEditEngineClient> {
  const pair = loopbackWorker(
    createPdfEditHandler({
      loadPdfium: () => fixturePdfium(),
      decodeImage: async () => {
        throw new Error("no images");
      },
      fetchBytes: async (url) => {
        if (url === FALLBACK_URL) return ttf;
        throw new Error(`No font at ${url}`);
      },
    }),
  );
  return loadPdfEditEngine(
    original,
    { format: "pdf", limits: defaultResourceLimits, signal },
    { createWorker: () => pair.worker, fallbackFontUrl: FALLBACK_URL },
  );
}

async function run(engine: PdfEditEngineClient, operation: PdfOperation) {
  const issues = await engine.validate([operation], signal);
  if (issues.length > 0) return { issues };
  return { change: await engine.apply([operation], signal) };
}

const codes = (issues: readonly OperationIssue[] | undefined) =>
  (issues ?? []).map((issue) => `${issue.path}:${issue.code}`);

async function element(
  engine: PdfEditEngineClient,
  id: string,
): Promise<PdfElement> {
  return (await engine.getElement!(id, signal)) as PdfElement;
}

/** A page whose text uses the embedded Noto Sans face, saved by PDFium. */
async function embeddedFontPdf(text: string): Promise<Uint8Array> {
  const pdfium = await fixturePdfium();
  const { lib } = pdfium;
  const document = pdfium.createDocument();
  try {
    const page = lib.FPDFPage_New(document.handle, 0, 612, 792);
    const data = pdfium.writeBytes(ttf);
    const font = lib.FPDFText_LoadFont(
      document.handle,
      data,
      ttf.length,
      1,
      true,
    );
    const object = lib.FPDFPageObj_CreateTextObj(document.handle, font, 14);
    const wide = pdfium.writeWideString(text);
    lib.FPDFText_SetText(object, wide);
    pdfium.free(wide);
    lib.FPDFPageObj_Transform(object, 1, 0, 0, 1, 72, 700);
    lib.FPDFPage_InsertObject(page, object);
    lib.FPDFPage_GenerateContent(page);
    lib.FPDF_ClosePage(page);
    const bytes = document.save("full");
    pdfium.free(data);
    return bytes;
  } finally {
    document.close();
  }
}

describe("edits to existing text objects", () => {
  let original: Uint8Array;

  before(async () => {
    original = await buildPdf([
      {
        texts: [
          { text: "Hello", x: 72, y: 700 },
          { text: "Red", x: 72, y: 600, color: [255, 0, 0], fontSize: 20 },
        ],
      },
    ]);
  });

  it("replaces text in place when the font can draw it", async () => {
    const engine = await engineFor(original);
    try {
      const before = await element(engine, "p0:o0");
      const result = await run(
        engine,
        op({ op: "replaceText", target: "p0:o0", text: "Goodbye, world" }),
      );
      assert.deepEqual(result.change?.warnings, []);
      const after = await element(engine, "p0:o0");
      assert.equal(after.text, "Goodbye, world");
      assert.equal(after.textStyle?.fontFamily, "Helvetica");
      assert.ok(Math.abs(after.bounds.x - before.bounds.x) < 0.5);
      assert.ok(Math.abs(after.bounds.y - before.bounds.y) < 0.5);
      assert.ok(after.bounds.width > before.bounds.width);
      const saved = await engine.materialize(signal);
      assert.equal(await extractPageText(saved, 0), "Goodbye, world\r\nRed");
    } finally {
      await engine.dispose();
    }
  });

  it("falls back to a covering font when the original cannot draw the text", async () => {
    const engine = await engineFor(original);
    try {
      const before = await element(engine, "p0:o1");
      const result = await run(
        engine,
        op({ op: "replaceText", target: "p0:o1", text: "Червоний" }),
      );
      assert.equal(result.change?.warnings[0]?.code, "font-substitution");
      const after = await element(engine, "p0:o1");
      assert.equal(after.id, "p0:o1");
      assert.equal(after.text, "Червоний");
      assert.equal(after.textStyle?.fontFamily, "Noto Sans");
      assert.equal(after.textStyle?.fontSize, 20);
      assert.equal(after.textStyle?.color, "#ff0000");
      assert.ok(Math.abs(after.bounds.x - before.bounds.x) < 1.5);
      // The baseline is kept, so the boxes overlap vertically even though
      // Cyrillic descenders make the new ink box taller.
      assert.ok(after.bounds.y < before.bounds.y + before.bounds.height);
      assert.ok(after.bounds.y + after.bounds.height > before.bounds.y);
      assert.equal(
        await extractPageText(await engine.materialize(signal), 0),
        "Hello\r\nЧервоний",
      );
    } finally {
      await engine.dispose();
    }
  });

  it("edits embedded-font text in place when its cmap covers the new text", async () => {
    const engine = await engineFor(await embeddedFontPdf("Привіт"));
    try {
      const kept = await run(
        engine,
        op({ op: "replaceText", target: "p0:o0", text: "Бувай" }),
      );
      assert.deepEqual(kept.change?.warnings, []);
      assert.equal(
        await extractPageText(await engine.materialize(signal), 0),
        "Бувай",
      );
      const cjk = await run(
        engine,
        op({ op: "replaceText", target: "p0:o0", text: "日本" }),
      );
      assert.deepEqual(codes(cjk.issues), ["/text:font-unavailable"]);
    } finally {
      await engine.dispose();
    }
  });

  it("changes colour and size on existing text and rejects other styles", async () => {
    const engine = await engineFor(original);
    try {
      const before = await element(engine, "p0:o0");
      const result = await run(
        engine,
        op({
          op: "setTextStyle",
          target: "p0:o0",
          style: { color: "#0000ff", fontSize: 24 },
        }),
      );
      assert.deepEqual(result.change?.warnings, []);
      const after = await element(engine, "p0:o0");
      assert.deepEqual(after.textStyle, {
        fontFamily: "Helvetica",
        fontSize: 24,
        bold: false,
        italic: false,
        color: "#0000ff",
      });
      assert.ok(Math.abs(after.bounds.x - before.bounds.x) < 0.5);
      assert.ok(Math.abs(after.bounds.y - before.bounds.y) < 0.5);
      assert.ok(after.bounds.height > before.bounds.height * 1.8);
      const issues = await engine.validate(
        [
          op({ op: "setTextStyle", target: "p0:o0", style: { bold: true } }),
          op({ op: "replaceText", target: "p0:o0", text: "שלום" }),
          op({ op: "replaceText", target: "p0:o9", text: "x" }),
        ],
        signal,
      );
      assert.deepEqual(
        issues.map(
          (issue) => `${issue.operationIndex}${issue.path}:${issue.code}`,
        ),
        [
          "0/style/bold:unsupported-style",
          "1/text:unsupported-script",
          "2/target:unknown-target",
        ],
      );
    } finally {
      await engine.dispose();
    }
  });

  // ACTION-908: a title written as `1 Tf` with its size in the text matrix
  // read as size 1, and setting 24 drew it at 24 × 24.
  it("reads and sets the size a text matrix carries, keeping its turn", async () => {
    const scaled = await buildPdf([
      {
        texts: [
          {
            text: "Upright title",
            fontSize: 1,
            matrix: [24, 0, 0, 24],
            y: 700,
          },
          {
            text: "Turned title",
            fontSize: 1,
            matrix: [0, 24, -24, 0],
            x: 300,
            y: 300,
          },
        ],
      },
    ]);
    const engine = await engineFor(scaled);
    try {
      for (const id of ["p0:o0", "p0:o1"]) {
        const before = await element(engine, id);
        assert.equal(before.textStyle?.fontSize, 24, `${id} reads 24 pt`);
        await run(
          engine,
          op({ op: "setTextStyle", target: id, style: { fontSize: 12 } }),
        );
        const after = await element(engine, id);
        assert.equal(after.textStyle?.fontSize, 12, `${id} reads 12 pt`);
        assert.equal(after.text, before.text);
        // Half the size along the text, and still turned the same way.
        const along = (box: PdfElement["bounds"]) =>
          Math.max(box.width, box.height);
        const across = (box: PdfElement["bounds"]) =>
          Math.min(box.width, box.height);
        assert.ok(
          Math.abs(along(after.bounds) / along(before.bounds) - 0.5) < 0.05,
          `${id} along`,
        );
        assert.ok(
          Math.abs(across(after.bounds) / across(before.bounds) - 0.5) < 0.1,
          `${id} across`,
        );
        assert.equal(
          after.bounds.width > after.bounds.height,
          before.bounds.width > before.bounds.height,
        );
      }
      // Replacing the text keeps the size it shows.
      await run(
        engine,
        op({ op: "replaceText", target: "p0:o0", text: "New title" }),
      );
      assert.equal((await element(engine, "p0:o0")).textStyle?.fontSize, 12);
    } finally {
      await engine.dispose();
    }
  });

  it("replays existing-text edits deterministically", async () => {
    const batch = [
      op({ op: "replaceText", target: "p0:o0", text: "Replayed" }),
      op({ op: "setTextStyle", target: "p0:o1", style: { fontSize: 10 } }),
    ];
    const first = await engineFor(original);
    const second = await engineFor(original);
    try {
      await first.apply(batch, signal);
      await second.restore([batch], signal);
      assert.deepEqual(
        await second.materialize(signal),
        await first.materialize(signal),
      );
    } finally {
      await first.dispose();
      await second.dispose();
    }
  });
});
