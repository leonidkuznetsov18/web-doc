import assert from "node:assert/strict";
import { describe, it } from "node:test";

import { prepareDocxForDisplay } from "../src/adapters/docx-prepass.js";
import { createOoxmlEditHandler } from "../src/edit/pptx/handler.js";
import { OoxmlPackage } from "../src/edit/ooxml/package.js";
import {
  defaultResourceLimits,
  ViewerError,
  type EditOperation,
} from "../src/index.js";
import { buildDocx, paragraph, sectPr } from "./fixtures/docx-builder.js";
import { docxSession, run, type FakePages } from "./fixtures/docx-session.js";

/*
 * Task 54, the main-thread side: the DOCX session joins the engine's
 * elements with the renderer's runs (fragments per page, tables through
 * their cells, inline objects through their paragraph), answers hit tests
 * and text searches with run geometry, reads only the pages a query names
 * or the viewer has cached, and serves everything through the worker
 * protocol.
 */

const limits = defaultResourceLimits;

/** The pre-pass ids of the main part's paragraphs, in document order. */
async function idsOf(bytes: Uint8Array): Promise<string[]> {
  const display = await prepareDocxForDisplay(bytes, limits);
  const pkg = await OoxmlPackage.open(display.bytes, { limits });
  const xml = new TextDecoder().decode(await pkg.part("/word/document.xml"));
  return [...xml.matchAll(/w:name="_wd([0-9A-F]{8})"/g)].map((m) => m[1]!);
}

const DOCUMENT = buildDocx({
  body:
    paragraph("Intro paragraph") +
    `<w:tbl><w:tr><w:tc>${paragraph("Cell A")}</w:tc><w:tc>${paragraph("Cell B")}</w:tc></w:tr></w:tbl>` +
    paragraph("Long paragraph that continues") +
    `<w:p/>` +
    paragraph("Last") +
    sectPr(),
});

describe("DOCX edit session: geometry join (docx-edit T54)", () => {
  it("places paragraphs, tables and the rest from the runs of the pages a query names", async () => {
    const [intro, cellA, cellB, long, , last] = await idsOf(DOCUMENT);
    const pages: FakePages = [
      [
        run(intro!, "Intro ", 72, 72),
        run(intro!, "paragraph", 108, 72),
        run(cellA!, "Cell A", 72, 100),
        run(cellB!, "Cell B", 300, 100),
        run(long!, "Long paragraph", 72, 130),
      ],
      [run(long!, "that continues", 72, 72), run(last!, "Last", 72, 90)],
    ];
    const { session, reads, end } = await docxSession(DOCUMENT, pages);
    try {
      const page0 = (await session.getElements({ pageIndex: 0 })).items;
      assert.deepEqual(reads, [0]);
      assert.deepEqual(
        page0.map((element) => [element.kind, element.pageIndex]),
        [
          ["paragraph", 0],
          ["table", 0],
          ["paragraph", 0],
          ["paragraph", 0],
          ["paragraph", 0],
        ],
      );
      assert.deepEqual(page0[0]!.bounds, {
        x: 72,
        y: 72,
        width: 36 + 54,
        height: 12,
      });
      assert.deepEqual(page0[1]!.bounds, {
        x: 72,
        y: 100,
        width: 300 + 36 - 72,
        height: 12,
      });
      assert.deepEqual(page0[4]!.fragments, [
        { pageIndex: 0, bounds: { x: 72, y: 130, width: 84, height: 12 } },
      ]);
      // The empty paragraph draws no run and "Last" is on page 1.
      assert.equal(page0.length, 5);

      const page1 = (await session.getElements({ pageIndex: 1 })).items;
      assert.deepEqual(
        page1.map((element) => [element.text, element.pageIndex]),
        [
          ["Long paragraph that continues", 1],
          ["Last", 1],
        ],
      );

      // Without a page, only the pages the viewer has laid out are read:
      // both now, so the long paragraph has two fragments.
      const all = (await session.getElements()).items;
      assert.deepEqual(reads, [0, 1, 0, 1]);
      const long2 = all.find((element) => element.id === `p:${long}`)!;
      assert.equal(long2.pageIndex, 0);
      assert.equal(long2.fragments!.length, 2);
      assert.deepEqual(long2.bounds, { x: 72, y: 130, width: 84, height: 12 });
      const empty = all[5]!;
      assert.equal(empty.text, "");
      assert.equal(empty.pageIndex, -1);
      assert.deepEqual(empty.fragments, []);

      const intersecting = (
        await session.getElements({
          pageIndex: 0,
          intersects: { x: 290, y: 95, width: 20, height: 20 },
        })
      ).items;
      assert.deepEqual(
        intersecting.map((element) => element.id),
        [`tbl:${cellA}`, `p:${cellB}`],
      );
      assert.deepEqual(
        (await session.getElements({ pageIndex: 0, kinds: ["table"] })).items
          .length,
        1,
      );
      const one = await session.getElement(`tbl:${cellA}`);
      assert.equal(one.item?.pageIndex, 0);
      assert.equal((await session.getElement("p:FFFFFFFF")).item, undefined);
    } finally {
      await end();
    }
  });

  it("reads nothing for a query without a page when no page is cached", async () => {
    const { session, reads, end } = await docxSession(DOCUMENT, [[], []]);
    try {
      const all = (await session.getElements()).items;
      assert.deepEqual(reads, []);
      assert.ok(all.every((element) => element.pageIndex === -1));
      assert.equal(all.length, 7);
    } finally {
      await end();
    }
  });

  it("hit-tests through the runs and returns the paragraph with its table", async () => {
    const [intro, cellA, cellB] = await idsOf(DOCUMENT);
    const pages: FakePages = [
      [
        run(intro!, "Intro", 72, 72),
        run(cellA!, "Cell A", 72, 100),
        run(cellB!, "Cell B", 300, 100),
      ],
    ];
    const { session, end } = await docxSession(DOCUMENT, pages);
    try {
      const hit = (await session.elementsAt(0, { x: 310, y: 105 })).items;
      assert.deepEqual(
        hit.map((element) => element.id),
        [`p:${cellB}`, `tbl:${cellA}`],
      );
      assert.equal(hit[1]!.bounds.width, 300 + 36 - 72);
      assert.deepEqual((await session.elementsAt(0, { x: 5, y: 5 })).items, []);
      const intro0 = (await session.elementsAt(0, { x: 80, y: 80 })).items;
      assert.deepEqual(
        intro0.map((element) => element.id),
        [`p:${intro}`],
      );
    } finally {
      await end();
    }
  });

  it("finds text with rectangles cut from the runs and pages located on demand", async () => {
    const [intro, , , long, , last] = await idsOf(DOCUMENT);
    const pages: FakePages = [
      [
        run(intro!, "Intro ", 72, 72, 60),
        run(intro!, "paragraph", 132, 72, 90),
      ],
      [run(long!, "that continues", 72, 72, 140), run(last!, "Last", 72, 90)],
      [run(last!, "nothing here", 10, 10)],
    ];
    const { session, reads, end } = await docxSession(DOCUMENT, pages);
    try {
      const hits = (await session.findText("paragraph")).items;
      assert.equal(hits.length, 2);
      // "Intro paragraph": the match covers the second run exactly.
      assert.deepEqual(hits[0], {
        pageIndex: 0,
        text: "paragraph",
        rects: [{ x: 132, y: 72, width: 90, height: 12 }],
        elementIds: [`p:${intro}`],
        ranges: [
          {
            start: { elementId: `p:${intro}`, offset: 6 },
            end: { elementId: `p:${intro}`, offset: 15 },
          },
        ],
      });
      // "Long paragraph that continues": page 0 holds no run of it in this
      // fixture, page 1 does; the match is absent from the run text, so
      // the paragraph's runs on that page stand in.
      assert.equal(hits[1]!.pageIndex, 1);
      assert.deepEqual(hits[1]!.rects, [
        { x: 72, y: 72, width: 140, height: 12 },
      ]);
      // Pages were read in order until the paragraph was found, plus the
      // next one to see whether the paragraph continues there.
      assert.deepEqual(reads, [0, 1, 2]);

      const partial = (await session.findText("cont", { pageRange: [1, 1] }))
        .items;
      assert.equal(partial.length, 1);
      assert.deepEqual(partial[0]!.rects, [
        { x: 72 + 5 * 10, y: 72, width: 40, height: 12 },
      ]);
      const second = (await session.findText("a", { maxResults: 3 })).items;
      assert.equal(second.length, 3);
      assert.deepEqual(await session.findText("").then((r) => r.items), []);
    } finally {
      await end();
    }
  });

  it("returns identical bytes without changes and refuses the presentation reads", async () => {
    const { session, end } = await docxSession(DOCUMENT, [[]]);
    try {
      const saved = await session.save();
      assert.deepEqual(saved.bytes, DOCUMENT);
      assert.equal(session.state.dirty, false);
      assert.equal(session.format, "docx");
      assert.deepEqual(Object.keys(session.schemas.operations), []);
      await assert.rejects(
        session.applyJson([
          { op: "replaceText", target: "p:x", text: "y" } as EditOperation,
        ]),
        (error: unknown) =>
          error instanceof ViewerError && error.code === "invalid-operation",
      );
    } finally {
      await end();
    }
    const handler = createOoxmlEditHandler();
    const signal = new AbortController().signal;
    const context = {
      signal,
      reportProgress: () => {},
      reportWarning: () => {},
    };
    await handler(
      "edit-open",
      { data: DOCUMENT.slice().buffer, limits, format: "docx" },
      context,
    );
    await assert.rejects(
      handler("edit-pptx-slides", undefined, context),
      (error: unknown) =>
        error instanceof ViewerError && error.code === "edit-unsupported",
    );
    await assert.rejects(
      handler(
        "edit-open",
        { data: DOCUMENT.slice().buffer, limits, format: "pdf" },
        context,
      ),
      (error: unknown) =>
        error instanceof ViewerError && error.code === "edit-unsupported",
    );
    await handler("edit-dispose", undefined, context);
  });
});
