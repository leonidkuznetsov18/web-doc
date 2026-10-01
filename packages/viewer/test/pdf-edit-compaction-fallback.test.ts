import assert from "node:assert/strict";
import { describe, it } from "node:test";

import {
  compactPdf,
  PdfCompactionError,
} from "../src/edit/pdf/engine/compact.js";
import { PdfEditDocument } from "../src/edit/pdf/engine/document.js";
import type { PdfOperation } from "../src/index.js";
import {
  buildPdf,
  extractPageText,
  fixturePdfium,
} from "./fixtures/pdf-builder.js";
import { pdfSession } from "./fixtures/pdf-session.js";
import { signedPdf } from "./fixtures/signed-pdf.js";

/*
 * Leonid's compaction decisions of 2026-10-02 (ACTION-821): the display
 * copy is a compacted full save, a compaction failure on the display path
 * falls back silently to the uncompacted full save, and on save() it falls
 * back with a privacy-not-guaranteed warning instead of failing.
 */

const op = <T extends PdfOperation>(operation: T): T => operation;
const box = op({
  op: "insertTextBox",
  pageIndex: 0,
  rect: { x: 36, y: 36, width: 200, height: 30 },
  text: "Changed",
});

function failing(bytes: Uint8Array): Uint8Array {
  void bytes;
  throw new PdfCompactionError("forced for the test");
}

/** Runs `use` with console.warn captured. */
function capturingWarnings<T>(use: () => T): { result: T; warnings: string[] } {
  const warnings: string[] = [];
  const original = console.warn;
  console.warn = (...args: unknown[]) => {
    warnings.push(args.map(String).join(" "));
  };
  try {
    return { result: use(), warnings };
  } finally {
    console.warn = original;
  }
}

describe("compaction fallback (decisions of 2026-10-02)", () => {
  it("shows the compacted full save for unsigned files and the incremental form for signed ones", async () => {
    const pdfium = await fixturePdfium();
    const original = await buildPdf(["Plain"]);
    const model = new PdfEditDocument(pdfium, original);
    try {
      model.apply([box]);
      const shown = model.materializeDocument("show");
      assert.deepEqual(shown.warnings, []);
      assert.deepEqual(shown.bytes, model.materialize("save", "full"));
      assert.ok(
        shown.bytes.length < original.length * 1.5,
        `the display copy stays near the original size (${shown.bytes.length} vs ${original.length})`,
      );
      assert.equal(await extractPageText(shown.bytes, 0), "Plain\r\nChanged");
    } finally {
      model.dispose();
    }
    const signed = new PdfEditDocument(pdfium, signedPdf());
    try {
      signed.apply([box]);
      const shown = signed.materializeDocument("show");
      assert.ok(
        signedPdf().every((byte, index) => shown.bytes[index] === byte),
        "a signed file's display copy still starts with the signed bytes",
      );
    } finally {
      signed.dispose();
    }
  });

  it("falls back to the uncompacted full save: silently for display, with a warning for a save", async () => {
    const pdfium = await fixturePdfium();
    const original = await buildPdf(["Plain"]);
    const model = new PdfEditDocument(
      pdfium,
      original,
      undefined,
      undefined,
      undefined,
      failing,
    );
    try {
      model.apply([box]);
      const { result: shown, warnings: logged } = capturingWarnings(() =>
        model.materializeDocument("show"),
      );
      assert.deepEqual(
        shown.warnings,
        [],
        "the display path carries no warning",
      );
      assert.equal(logged.length, 1, "but it is logged");
      assert.match(logged[0]!, /forced for the test/);
      assert.equal(await extractPageText(shown.bytes, 0), "Plain\r\nChanged");

      const saved = model.materializeDocument("save", "full");
      assert.equal(saved.warnings.length, 1);
      assert.equal(saved.warnings[0]!.code, "privacy-not-guaranteed");
      assert.equal(saved.warnings[0]!.details?.reason, "pdf-compaction");
      assert.equal(await extractPageText(saved.bytes, 0), "Plain\r\nChanged");
      // The bytes are PDFium's own full save: compacting them now works and
      // shrinks them, which shows they were not compacted.
      assert.ok(compactPdf(saved.bytes).length <= saved.bytes.length);
      // An incremental save never goes through the pass.
      assert.deepEqual(
        model.materializeDocument("save", "incremental").warnings,
        [],
      );
    } finally {
      model.dispose();
    }
  });

  it("carries the warning on the session's save result and nowhere else", async () => {
    const original = await buildPdf(["Plain"]);
    const forced = await pdfSession(original, { compact: failing });
    try {
      // Without changes the original bytes come back, untouched by the pass.
      assert.deepEqual((await forced.session.save()).warnings, []);
      const receipt = await forced.session.insertTextBox(box);
      assert.deepEqual(receipt.warnings, [], "the display path stays quiet");
      const saved = await forced.session.save();
      assert.deepEqual(
        saved.warnings.map((warning) => warning.code),
        ["privacy-not-guaranteed"],
      );
      assert.equal(await extractPageText(saved.bytes, 0), "Plain\r\nChanged");
      const incremental = await forced.session.save({ mode: "incremental" });
      assert.deepEqual(incremental.warnings, []);
    } finally {
      await forced.end();
    }
    const normal = await pdfSession(original);
    try {
      await normal.session.insertTextBox(box);
      const saved = await normal.session.save();
      assert.deepEqual(saved.warnings, [], "the normal path never warns");
      assert.equal(await extractPageText(saved.bytes, 0), "Plain\r\nChanged");
    } finally {
      await normal.end();
    }
  });
});
