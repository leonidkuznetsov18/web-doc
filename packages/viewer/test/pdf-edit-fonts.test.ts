import assert from "node:assert/strict";
import { createRequire } from "node:module";
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
      const file = new URL(url).pathname.split("/").at(-1);
      if (
        file &&
        /^LiberationSans-(Regular|Bold|Italic|BoldItalic)\.ttf$/.test(file)
      )
        return new Uint8Array(
          readFileSync(
            createRequire(import.meta.url).resolve(
              `pdfjs-dist/standard_fonts/${file}`,
            ),
          ),
        );
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

      // A styled request uses a real covering face, without changing regular fallback.
      const bold = await applyOne(engine, textBox("Жирний", { bold: true }));
      assert.equal(bold.change?.warnings.length, 1);
      assert.equal(
        fetched.filter((url: string) => url === FALLBACK_URL).length,
        1,
      );
      assert.equal(
        fetched.filter((url: string) =>
          url.endsWith("/LiberationSans-Bold.ttf"),
        ).length,
        1,
      );

      const cjk = await applyOne(engine, textBox("日本語"));
      assert.deepEqual(codes(cjk.issues), ["/text:font-unavailable"]);
      const unknownFamily = textBox("x", { fontFamily: "Comic" });
      for (const bold of [false, true]) {
        const unknown = await applyOne(engine, {
          ...unknownFamily,
          style: { fontFamily: "Comic", bold },
        });
        assert.deepEqual(codes(unknown.issues), [
          "/style/fontFamily:unknown-font",
        ]);
      }
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

  // ACTION-879: a Helvetica box made with Latin text, then given Ukrainian
  // text with a dash, a euro sign and a line break, as QA typed it.
  it("draws Cyrillic replacing the text of a Helvetica box", async () => {
    const { engine } = await engineFor(original);
    try {
      const inserted = await applyOne(engine, textBox("QA PDF 0.8.0"));
      const id = inserted.change!.createdIds[0]!;
      for (const text of [
        "Україна — PDF fallback after restart",
        "QA PDF 0.8.0\nУкраїнська перевірка — € 123",
      ]) {
        const replaced = await applyOne(
          engine,
          op({ op: "replaceText", target: id, text }),
        );
        assert.deepEqual(codes(replaced.issues), [], text);
      }
      assert.match(
        await extractPageText(await engine.materialize(signal), 0),
        /Українська перевірка — € 123/,
      );
      // A character no font has is named, not the first Cyrillic letter.
      const refused = await applyOne(
        engine,
        op({ op: "replaceText", target: id, text: "Україна ≠ 日本" }),
      );
      assert.match(refused.issues?.[0]?.message ?? "", /"≠" \(U\+2260\)/);
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
