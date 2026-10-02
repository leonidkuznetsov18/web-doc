import assert from "node:assert/strict";
import { describe, it } from "node:test";
import { fileURLToPath, pathToFileURL } from "node:url";

import type { TextRun, TextSelection } from "../src/contracts.js";
import { mapRangeThrough } from "../src/edit/pdf/range-map.js";
import { resolveSelection } from "../src/edit/pdf/selection.js";
import type { PageLayout, TextRange } from "../src/index.js";
import { buildPdf } from "./fixtures/pdf-builder.js";
import { pdfSession } from "./fixtures/pdf-session.js";

/*
 * The selection and range mapping of ACTION-825, task 34. Selections are
 * built the way the viewer builds them: PDF.js (the viewer's renderer, run
 * here under Node) extracts the runs of the fixture page, and the selected
 * span is cut out of them by logical offsets.
 */

const PACKAGE = pathToFileURL(`${process.cwd()}/`);

/** The page's runs as the PDF adapter produces them: PDF.js items in page space. */
async function pdfjsRuns(bytes: Uint8Array): Promise<TextRun[]> {
  const pdfjs = await import("pdfjs-dist/legacy/build/pdf.mjs");
  pdfjs.GlobalWorkerOptions.workerSrc = new URL(
    "../../node_modules/pdfjs-dist/legacy/build/pdf.worker.mjs",
    PACKAGE,
  ).href;
  const task = pdfjs.getDocument({
    data: bytes.slice(),
    disableFontFace: true,
    useSystemFonts: false,
    standardFontDataUrl: fileURLToPath(
      new URL("../../node_modules/pdfjs-dist/standard_fonts/", PACKAGE),
    ),
  });
  try {
    const document = await task.promise;
    const page = await document.getPage(1);
    const viewport = page.getViewport({ scale: 1 });
    const content = await page.getTextContent();
    const runs: TextRun[] = [];
    let logicalStart = 0;
    for (const item of content.items) {
      if (!("str" in item) || item.str.length === 0) continue;
      const style = (content.styles[item.fontName] ?? {}) as {
        readonly ascent?: number;
        readonly descent?: number;
      };
      // The adapter's textItemToRun, without links.
      const t = multiply(viewport.transform, item.transform);
      const angle = Math.atan2(t[1]!, t[0]!);
      const fontHeight = Math.max(1, Math.hypot(t[2]!, t[3]!));
      const ascent =
        style.ascent !== undefined
          ? style.ascent * fontHeight
          : style.descent !== undefined
            ? (1 + style.descent) * fontHeight
            : fontHeight;
      const left = t[4]! + fontHeight * Math.sin(angle) * (ascent / fontHeight);
      const top = t[5]! - fontHeight * Math.cos(angle) * (ascent / fontHeight);
      runs.push({
        text: item.str,
        x: left,
        y: top,
        width: Math.max(1, item.width * viewport.scale),
        height: Math.max(1, item.height * viewport.scale, fontHeight),
        textLayer: "pdf",
        coordinateWidth: viewport.width,
        coordinateHeight: viewport.height,
        logicalStart,
        logicalEnd: logicalStart + item.str.length,
      });
      logicalStart += item.str.length;
    }
    return runs;
  } finally {
    await task.destroy();
  }
}

function multiply(a: readonly number[], b: readonly number[]): number[] {
  return [
    a[0]! * b[0]! + a[2]! * b[1]!,
    a[1]! * b[0]! + a[3]! * b[1]!,
    a[0]! * b[2]! + a[2]! * b[3]!,
    a[1]! * b[2]! + a[3]! * b[3]!,
    a[0]! * b[4]! + a[2]! * b[5]! + a[4]!,
    a[1]! * b[4]! + a[3]! * b[5]! + a[5]!,
  ];
}

/** The viewer's selection of logical offsets [start, end) on one page. */
function select(
  runs: readonly TextRun[],
  start: number,
  end: number,
): TextSelection {
  const selected: TextRun[] = [];
  let offset = 0;
  for (const run of runs) {
    const runStart = offset;
    const runEnd = offset + run.text.length;
    offset = runEnd;
    const sliceStart = Math.max(start, runStart);
    const sliceEnd = Math.min(end, runEnd);
    if (sliceStart >= sliceEnd) continue;
    const localStart = sliceStart - runStart;
    const localEnd = sliceEnd - runStart;
    selected.push({
      ...run,
      text: run.text.slice(localStart, localEnd),
      x: run.x + (run.width * localStart) / run.text.length,
      width: (run.width * (localEnd - localStart)) / run.text.length,
    });
  }
  return {
    pageIndex: 0,
    startOffset: start,
    endOffset: end,
    text: runs
      .map((run) => run.text)
      .join("")
      .slice(start, end),
    runs: selected,
  };
}

const range = (elementId: string, start: number, end: number): TextRange => ({
  start: { elementId, offset: start },
  end: { elementId, offset: end },
});

describe("selection mapping (overlay primitives)", () => {
  it("resolves PDF.js selections through the three rungs of the ladder", async () => {
    const bytes = await buildPdf([
      {
        texts: [
          { text: "Alpha beta", x: 72, y: 700, fontSize: 24 },
          { text: "gamma delta", x: 72, y: 650, fontSize: 24 },
        ],
      },
    ]);
    const runs = await pdfjsRuns(bytes);
    assert.deepEqual(
      runs.map((run) => run.text),
      ["Alpha beta", "gamma delta"],
    );
    const { session, end } = await pdfSession(bytes);
    try {
      const { item: page } = await session.getPageLayout(0);
      const pages = [page as PageLayout];
      // Rung 1: a whole line, covered by its run.
      assert.deepEqual(resolveSelection(select(runs, 0, 10), pages), [
        range("p0:o0", 0, 10),
      ]);
      // Rung 2: a word inside the line, contained by it.
      assert.deepEqual(resolveSelection(select(runs, 6, 10), pages), [
        range("p0:o0", 6, 10),
      ]);
      // Across both lines: two ranges in reading order.
      assert.deepEqual(resolveSelection(select(runs, 6, 15), pages), [
        range("p0:o0", 6, 10),
        range("p0:o1", 0, 5),
      ]);
      // Rung 3: the geometry is off, the text still matches.
      const displaced = select(runs, 10, 15);
      const moved: TextSelection = {
        ...displaced,
        runs: displaced.runs.map((run) => ({
          ...run,
          x: run.x + 300,
          y: run.y + 50,
        })),
      };
      assert.deepEqual(resolveSelection(moved, pages), [range("p0:o1", 0, 5)]);
      // Nothing selected, nothing resolved; so for a run of spaces.
      assert.deepEqual(resolveSelection(select(runs, 3, 3), pages), []);
      const spaces: TextSelection = {
        ...displaced,
        runs: displaced.runs.map((run) => ({ ...run, text: "   " })),
      };
      assert.deepEqual(resolveSelection(spaces, pages), []);
      // The session method does the same with the engine's layouts.
      const viaSession = await session.elementsForSelection(
        select(runs, 6, 15),
      );
      assert.equal(viaSession.sessionId, session.sessionId);
      assert.deepEqual(viaSession.items, [
        range("p0:o0", 6, 10),
        range("p0:o1", 0, 5),
      ]);
    } finally {
      await end();
    }
  });

  it("maps ranges through replacements, deletions, undo, redo and reset", async () => {
    const original = await buildPdf(["Hello brave world"]);
    const { session, end } = await pdfSession(original);
    try {
      const brave = range("p0:o0", 6, 11);
      // Nothing happened yet: the range is its own image.
      assert.deepEqual((await session.mapRange(brave, 0)).item, brave);
      // A whole-text replacement keeps offsets up to the new text's length.
      await session.replaceText({ target: "p0:o0", text: "Hi all" });
      assert.deepEqual(
        (await session.mapRange(brave, 0)).item,
        range("p0:o0", 6, 6),
      );
      await session.undo();
      assert.deepEqual((await session.mapRange(brave, 0)).item, brave);
      // A deleted element is gone, and back after an undo.
      await session.deleteElement({ target: "p0:o0" });
      assert.equal((await session.mapRange(brave, 0)).item, undefined);
      await session.undo();
      assert.deepEqual((await session.mapRange(brave, 0)).item, brave);
      await session.redo();
      assert.equal((await session.mapRange(brave, 0)).item, undefined);
      await session.undo();
      // A created element does not survive a reset.
      const created = await session.insertTextBox({
        pageIndex: 0,
        rect: { x: 36, y: 36, width: 200, height: 40 },
        text: "box",
      });
      const inBox = range(created.createdIds[0]!, 0, 3);
      const at = session.state.revision;
      await session.reset();
      assert.equal((await session.mapRange(inBox, at)).item, undefined);
      assert.deepEqual((await session.mapRange(brave, at)).item, brave);
      // Revisions the session cannot reach are unknown.
      assert.equal((await session.mapRange(brave, 99)).item, undefined);
      // A dry run does not count as a mutation.
      const before = session.state.revision;
      await session.replaceText(
        { target: "p0:o0", text: "x" },
        { dryRun: true },
      );
      assert.equal(session.state.revision, before);
      assert.deepEqual((await session.mapRange(brave, before)).item, brave);
    } finally {
      await end();
    }
  });

  it("shifts offsets around a ranged replacement, forwards and back", () => {
    const receipt = {
      sessionId: "s",
      revision: 1,
      dryRun: false,
      operationCount: 1,
      createdIds: [],
      removedIds: [],
      changedPages: [0],
      pageCount: 1,
      warnings: [],
    };
    const replace = {
      op: "replaceText",
      target: "p0:o0",
      text: "bold",
      range: range("p0:o0", 6, 11),
    };
    const forward = [
      { revision: 1, kind: "apply" as const, operations: [replace], receipt },
    ];
    // Before the span: unchanged. Covering it: the end follows the new length.
    assert.deepEqual(
      mapRangeThrough(range("p0:o0", 0, 5), forward),
      range("p0:o0", 0, 5),
    );
    assert.deepEqual(
      mapRangeThrough(range("p0:o0", 6, 11), forward),
      range("p0:o0", 6, 10),
    );
    // After the span: shifted by the length difference (-1).
    assert.deepEqual(
      mapRangeThrough(range("p0:o0", 12, 17), forward),
      range("p0:o0", 11, 16),
    );
    // Inside the span: clamped to the replacement.
    assert.deepEqual(
      mapRangeThrough(range("p0:o0", 8, 9), forward),
      range("p0:o0", 8, 9),
    );
    assert.deepEqual(
      mapRangeThrough(range("p0:o0", 9, 11), forward),
      range("p0:o0", 9, 10),
    );
    // Taken back by an undo: the shift reverses.
    const back = [
      ...forward,
      {
        revision: 2,
        kind: "undo" as const,
        operations: [replace],
        receipt: { ...receipt, revision: 2 },
      },
    ];
    assert.deepEqual(
      mapRangeThrough(range("p0:o0", 12, 17), back),
      range("p0:o0", 12, 17),
    );
    // Another element is untouched; a renamed one is followed.
    assert.deepEqual(
      mapRangeThrough(range("p0:o1", 2, 4), forward),
      range("p0:o1", 2, 4),
    );
    const renamed = [
      {
        revision: 1,
        kind: "apply" as const,
        operations: [],
        receipt: { ...receipt, remappedIds: { "p0:o0": "p0:n1.0.0" } },
      },
    ];
    assert.deepEqual(
      mapRangeThrough(range("p0:o0", 1, 2), renamed),
      range("p0:n1.0.0", 1, 2),
    );
  });
});
