import assert from "node:assert/strict";
import { readFileSync } from "node:fs";
import { before, describe, it } from "node:test";

import { createPdfEditHandler } from "../src/edit/pdf/engine/handler.js";
import { parseCmap } from "../src/edit/pdf/engine/fonts.js";
import {
  loadPdfEditEngine,
  type PdfEditEngineClient,
} from "../src/edit/pdf/provider.js";
import type { OperationIssue, PdfOperation } from "../src/index.js";
import { defaultResourceLimits } from "../src/index.js";
import { loopbackWorker } from "./fixtures/loopback-worker.js";
import {
  buildPdf,
  extractPageText,
  fixturePdfium,
} from "./fixtures/pdf-builder.js";

// Tests run compiled from .test-dist/test, two levels below the package root.
const ttf = new Uint8Array(
  readFileSync(
    new URL("../../fonts/noto-sans-latin-cyrillic.ttf", import.meta.url),
  ),
);
const FALLBACK_URL = "https://fonts.test/noto.ttf";
const signal = new AbortController().signal;

const op = <T extends PdfOperation>(operation: T): T => operation;

let nextY = 200;

/** Each box gets its own band of the page, so extracted lines stay apart. */
function textBox(text: string, style: Record<string, unknown> = {}) {
  nextY += 60;
  return op({
    op: "insertTextBox",
    pageIndex: 0,
    rect: { x: 72, y: nextY, width: 400, height: 50 },
    text,
    style,
  });
}

/** An engine over the loopback worker, with a counting font fetcher. */
async function engineFor(
  original: Uint8Array,
  fonts: {
    family: string;
    weight?: number;
    style?: "normal" | "italic";
  }[] = [],
) {
  const fetched: string[] = [];
  const handler = createPdfEditHandler({
    loadPdfium: () => fixturePdfium(),
    decodeImage: async () => {
      throw new Error("no images");
    },
    fetchBytes: async (url) => {
      fetched.push(url);
      if (url === FALLBACK_URL) return ttf;
      throw new Error(`No font at ${url}`);
    },
  });
  const pair = loopbackWorker(handler);
  const engine = await loadPdfEditEngine(
    original,
    {
      format: "pdf",
      limits: defaultResourceLimits,
      signal,
      fonts: fonts.map((font) => ({ ...font, source: ttf })),
    },
    { createWorker: () => pair.worker, fallbackFontUrl: FALLBACK_URL },
  );
  return { engine, fetched };
}

async function applyOne(engine: PdfEditEngineClient, operation: PdfOperation) {
  const issues = await engine.validate([operation], signal);
  if (issues.length > 0) return { issues };
  const change = await engine.apply([operation], signal);
  return { change };
}

function codes(issues: readonly OperationIssue[] | undefined): string[] {
  return (issues ?? []).map((issue) => `${issue.path}:${issue.code}`);
}

describe("fonts for inserted PDF text", () => {
  let original: Uint8Array;

  before(async () => {
    original = await buildPdf(["Existing"]);
  });

  it("reads coverage from a font's cmap", () => {
    const coverage = parseCmap(ttf);
    assert.equal(coverage.has("п".codePointAt(0)!), true);
    assert.equal(coverage.has("A".codePointAt(0)!), true);
    assert.equal(coverage.has("日".codePointAt(0)!), false);
    assert.equal(parseCmap(new Uint8Array(10)).has(65), false);
  });

  it("falls back to the bundled Noto Sans for text the standard fonts cannot encode", async () => {
    const { engine, fetched } = await engineFor(original);
    try {
      const latin = await applyOne(engine, textBox("Plain Latin, € included"));
      assert.deepEqual(latin.change?.warnings, []);
      assert.deepEqual(fetched, [], "no font fetched for WinAnsi text");

      const cyrillic = await applyOne(engine, textBox("Привіт, світе!"));
      assert.deepEqual(fetched, [FALLBACK_URL]);
      assert.equal(cyrillic.change?.warnings[0]?.code, "font-substitution");
      const saved = await engine.materialize(signal);
      assert.equal(
        await extractPageText(saved, 0),
        "Existing\r\nPlain Latin, € included\r\nПривіт, світе!",
      );
      const box = await engine.getElement!("p0:n2.0.0", signal);
      assert.equal(box?.kind, "textBox");
      assert.equal(
        (box as { textStyle?: { fontFamily: string } }).textStyle?.fontFamily,
        "Noto Sans",
      );

      // Bold is asked for but the fallback has one face: drawn regular, with a note.
      const bold = await applyOne(engine, textBox("Жирний", { bold: true }));
      assert.equal(bold.change?.warnings.length, 1);
      assert.deepEqual(fetched, [FALLBACK_URL], "fetched once");

      const cjk = await applyOne(engine, textBox("日本語"));
      assert.deepEqual(codes(cjk.issues), ["/text:font-unavailable"]);
      const unknown = await applyOne(
        engine,
        textBox("x", { fontFamily: "Comic" }),
      );
      assert.deepEqual(codes(unknown.issues), [
        "/style/fontFamily:unknown-font",
      ]);
    } finally {
      await engine.dispose();
    }
  });

  it("uses a registered family without a substitution note", async () => {
    const { engine, fetched } = await engineFor(original, [
      { family: "House Sans" },
    ]);
    try {
      const result = await applyOne(
        engine,
        textBox("Зареєстрований шрифт", { fontFamily: "House Sans" }),
      );
      assert.deepEqual(result.change?.warnings, []);
      assert.deepEqual(fetched, [], "registered bytes need no fetch");
      assert.equal(
        await extractPageText(await engine.materialize(signal), 0),
        "Existing\r\nЗареєстрований шрифт",
      );
      // A registered family also serves Latin text, in place of a standard font.
      const latin = await applyOne(
        engine,
        textBox("Latin too", { fontFamily: "house sans", italic: true }),
      );
      assert.equal(latin.change?.warnings[0]?.code, "font-substitution");
    } finally {
      await engine.dispose();
    }
  });

  it("keeps fonts working across undo, redo and reopen", async () => {
    const { engine } = await engineFor(original);
    try {
      const box = textBox("Привіт");
      await engine.apply([box], signal);
      const first = await engine.materialize(signal);
      await engine.restore([], signal);
      assert.deepEqual(await engine.materialize(signal), original);
      await engine.restore([[box]], signal);
      assert.deepEqual(await engine.materialize(signal), first);
      await engine.apply(
        [op({ op: "replaceText", target: "p0:n1.0.0", text: "Бувай" })],
        signal,
      );
      assert.equal(
        await extractPageText(await engine.materialize(signal), 0),
        "Existing\r\nБувай",
      );
    } finally {
      await engine.dispose();
    }
  });
});
