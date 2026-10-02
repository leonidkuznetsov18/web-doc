import assert from "node:assert/strict";
import { describe, it } from "node:test";

import type { ResourceLimits } from "../src/contracts.js";
import { OoxmlPackage } from "../src/edit/ooxml/package.js";
import { PptxEditEngine } from "../src/edit/pptx/engine.js";
import { emuToPx } from "../src/edit/pptx/geometry.js";
import { createOoxmlEditHandler } from "../src/edit/pptx/handler.js";
import { loadPptxEditEngine } from "../src/edit/pptx/provider.js";
import { PptxSession } from "../src/edit/pptx/session.js";
import {
  EditSessionController,
  type EditSessionHost,
} from "../src/edit/session.js";
import type { WorkerOperation } from "../src/worker-protocol.js";
import {
  defaultResourceLimits,
  type PptxEditSession,
  type PptxElement,
  type PptxOperation,
} from "../src/index.js";
import { loopbackWorker } from "./fixtures/loopback-worker.js";
import {
  buildDeck,
  group,
  picture,
  textShape,
} from "./fixtures/pptx-builder.js";
import { pptxSession } from "./fixtures/pptx-session.js";
import { tinyJpeg } from "./fixtures/tiny-jpeg.js";

/*
 * Review of the PPTX module's coverage: the behaviours the spec promises or
 * the code branches on that the task tests do not reach — rotated and
 * flipped groups, the text model at its boundaries, collapsed style ranges,
 * malformed slide parts, decks without layouts or slides, id allocation
 * against the layout, external relationships, restores from a checkpoint
 * base, the worker handler's lifecycle and abort signals.
 */

const signal = new AbortController().signal;
const SLIDE = "/ppt/slides/slide1.xml";
const PRESENTATION = "/ppt/presentation.xml";
const REL =
  "http://schemas.openxmlformats.org/officeDocument/2006/relationships/";
const XML = '<?xml version="1.0" encoding="UTF-8" standalone="yes"?>\n';
const limits = defaultResourceLimits;

const run = (engine: PptxEditEngine, operations: PptxOperation[]) =>
  engine.apply(operations, signal);
const check = (engine: PptxEditEngine, operations: PptxOperation[]) =>
  engine.validate(operations, signal);

async function open(bytes: Uint8Array): Promise<PptxEditEngine> {
  return PptxEditEngine.open(bytes, limits, signal);
}

async function partOf(bytes: Uint8Array, name: string): Promise<string> {
  const pkg = await OoxmlPackage.open(bytes, { limits });
  return new TextDecoder().decode(await pkg.part(name));
}

async function partText(engine: PptxEditEngine, name: string): Promise<string> {
  return partOf(await engine.materialize("save", {}, signal), name);
}

async function element(
  engine: PptxEditEngine,
  id: string,
): Promise<PptxElement> {
  const found = await engine.getElement(id, signal);
  assert.ok(found, `element ${id}`);
  return found;
}

/** The package with some parts replaced byte for byte, scanning nothing. */
async function withParts(
  bytes: Uint8Array,
  parts: Readonly<Record<string, Uint8Array | string>>,
): Promise<Uint8Array> {
  const pkg = await OoxmlPackage.open(bytes, { limits });
  const set = new Map<string, Uint8Array>();
  for (const [name, data] of Object.entries(parts))
    set.set(
      name,
      typeof data === "string" ? new TextEncoder().encode(data) : data,
    );
  pkg.applyOverlay({ set, remove: new Set() });
  return pkg.save();
}

function px(box: { x: number; y: number; cx: number; cy: number }) {
  return {
    x: emuToPx(box.x),
    y: emuToPx(box.y),
    width: emuToPx(box.cx),
    height: emuToPx(box.cy),
  };
}

function close(actual: number, expected: number, tolerance = 1e-6): void {
  assert.ok(
    Math.abs(actual - expected) <= tolerance,
    `expected ${actual} to be within ${tolerance} of ${expected}`,
  );
}

function sameRect(
  actual: { x: number; y: number; width: number; height: number },
  expected: { x: number; y: number; width: number; height: number },
  tolerance = 1e-6,
): void {
  close(actual.x, expected.x, tolerance);
  close(actual.y, expected.y, tolerance);
  close(actual.width, expected.width, tolerance);
  close(actual.height, expected.height, tolerance);
}

const RANGE = (id: string, start: number, end: number) => ({
  start: { elementId: id, offset: start },
  end: { elementId: id, offset: end },
});

/** The session fixture's wiring with the limits a test chooses. */
async function sessionWith(
  original: Uint8Array,
  own: ResourceLimits,
): Promise<{ session: PptxEditSession; end(): Promise<void> }> {
  const count = async (bytes: Uint8Array): Promise<number> => {
    const engine = await PptxEditEngine.open(bytes, own);
    try {
      return engine.pageCount;
    } finally {
      await engine.dispose();
    }
  };
  const pair = loopbackWorker(createOoxmlEditHandler());
  const engine = await loadPptxEditEngine(
    original,
    { format: "pptx", limits: own, signal },
    { createWorker: () => pair.worker },
  );
  const host: EditSessionHost = {
    format: "pptx",
    limits: own,
    prepareDocument: async (bytes) => ({ pageCount: await count(bytes) }),
    commitDocument: (prepared) => prepared.pageCount,
    discardDocument: () => {},
    emit: () => {},
  };
  const core = new EditSessionController(
    engine,
    host,
    original,
    await count(original),
  );
  return { session: new PptxSession(core), end: () => core.end() };
}

describe("PPTX review coverage (pptx-edit)", () => {
  it("maps a child through a rotated group: frame, bounds, hit-testing and writes through the inverse", async () => {
    // A 192 × 96 px group rotated 90° about its centre (96, 48), child space
    // at scale 1, holding a 96 × 48 child at the child-space origin.
    const engine = await open(
      buildDeck({
        slides: [
          {
            shapes: [
              group({
                id: 2,
                x: 0,
                y: 0,
                cx: 1828800,
                cy: 914400,
                rotation: 90,
                child: { x: 0, y: 0, cx: 1828800, cy: 914400 },
                children: [
                  textShape({ id: 3, x: 0, y: 0, cx: 914400, cy: 457200 }),
                ],
              }),
            ],
          },
        ],
      }),
    );
    const child = await element(engine, "sld1:3");
    assert.equal(child.parentId, "sld1:2");
    assert.equal(child.rotation, 90);
    assert.deepEqual(child.frame, {
      x: 72,
      y: -24,
      width: 96,
      height: 48,
      rotation: 90,
      flipH: false,
      flipV: false,
    });
    sameRect(child.bounds, { x: 96, y: -48, width: 48, height: 96 });
    sameRect((await element(engine, "sld1:2")).bounds, {
      x: 48,
      y: -48,
      width: 96,
      height: 192,
    });
    // Hit-testing uses the rotated frames: inside the child, inside the
    // group only, and inside the group's axis-aligned box but outside its frame.
    const ids = async (x: number, y: number): Promise<string[]> =>
      (await engine.elementsAt(0, { x, y }, signal)).map((item) => item.id);
    assert.deepEqual(await ids(100, 0), ["sld1:3", "sld1:2"]);
    assert.deepEqual(await ids(60, 100), ["sld1:2"]);
    assert.deepEqual(await ids(140, 60), ["sld1:2"]);
    assert.deepEqual(await ids(0, 0), []);

    // A slide-space move of (10, 20) is a child-space move of (20, -10).
    await run(engine, [
      { op: "moveElement", target: "sld1:3", by: { dx: 10, dy: 20 } },
    ]);
    sameRect((await element(engine, "sld1:3")).bounds, {
      x: 106,
      y: -28,
      width: 48,
      height: 96,
    });
    assert.ok(
      (await partText(engine, SLIDE)).includes(
        '<a:off x="190500" y="-95250"/><a:ext cx="914400" cy="457200"/>',
      ),
    );
    // New bounds for the rotated child: the frame is solved for the box.
    await run(engine, [
      {
        op: "resizeElement",
        target: "sld1:3",
        rect: { x: 96, y: -48, width: 24, height: 48 },
      },
    ]);
    const resized = await element(engine, "sld1:3");
    sameRect(resized.bounds, { x: 96, y: -48, width: 24, height: 48 });
    assert.equal(resized.rotation, 90);
    assert.ok(
      (await partText(engine, SLIDE)).includes(
        '<a:off x="0" y="228600"/><a:ext cx="457200" cy="228600"/>',
      ),
    );
    await engine.dispose();
  });

  it("reports the orientation of a child inside a flipped group (bug: the rotation is 180° off when the group has exactly one flip)", async () => {
    // flipH then rot 90° on the group (DrawingML applies the flip first):
    // the child's transform is R(90)·F_h, so its frame is rotation 90 with
    // flipH, not rotation 270. The bounds agree either way (a box is
    // symmetric under 180°), so only `frame.rotation`/`rotation` is wrong.
    const engine = await open(
      buildDeck({
        slides: [
          {
            shapes: [
              group({
                id: 2,
                x: 0,
                y: 0,
                cx: 1828800,
                cy: 914400,
                rotation: 90,
                flipH: true,
                child: { x: 0, y: 0, cx: 1828800, cy: 914400 },
                children: [
                  textShape({ id: 3, x: 0, y: 0, cx: 914400, cy: 457200 }),
                ],
              }),
              group({
                id: 4,
                x: 0,
                y: 0,
                cx: 1828800,
                cy: 914400,
                rotation: 90,
                flipV: true,
                child: { x: 0, y: 0, cx: 1828800, cy: 914400 },
                children: [
                  textShape({ id: 5, x: 0, y: 0, cx: 914400, cy: 457200 }),
                ],
              }),
            ],
          },
        ],
      }),
    );
    const flippedH = await element(engine, "sld1:3");
    sameRect(flippedH.bounds, { x: 96, y: 48, width: 48, height: 96 });
    assert.deepEqual(
      [flippedH.frame!.rotation, flippedH.frame!.flipH],
      [90, true],
    );
    // flipV then rot 90° is R(90)·F_v = R(270)·F_h.
    const flippedV = await element(engine, "sld1:5");
    assert.deepEqual(
      [flippedV.frame!.rotation, flippedV.frame!.flipH],
      [270, true],
    );
    await engine.dispose();
  });

  it("round-trips line and paragraph breaks at the text boundaries and edits runs around a:br", async () => {
    const engine = await open(
      buildDeck({
        slides: [
          {
            shapes: [
              textShape({
                id: 2,
                x: 0,
                y: 0,
                cx: 914400,
                cy: 914400,
                paragraphs: [["Alpha"]],
              }),
            ],
          },
        ],
      }),
    );
    const text = "\u000bLead\nMid\u000b\u000bEnd\u000b\n";
    await run(engine, [{ op: "replaceText", target: "sld1:2", text }]);
    assert.equal((await element(engine, "sld1:2")).text, text);
    const BR = '<a:br><a:rPr lang="en-US"/></a:br>';
    const RUN = (value: string) =>
      `<a:r><a:rPr lang="en-US"/><a:t>${value}</a:t></a:r>`;
    const END = '<a:endParaRPr lang="en-US"/>';
    const body = async (): Promise<string> =>
      /<a:lstStyle\/>(.*)<\/p:txBody>/s.exec(
        await partText(engine, SLIDE),
      )![1]!;
    assert.equal(
      await body(),
      `<a:p>${BR}${RUN("Lead")}${END}</a:p>` +
        `<a:p>${RUN("Mid")}${BR}${BR}${RUN("End")}${BR}${END}</a:p>` +
        `<a:p>${END}</a:p>`,
    );
    // Deleting the leading break keeps the run after it byte for byte.
    await run(engine, [
      {
        op: "replaceText",
        target: "sld1:2",
        text: "",
        range: RANGE("sld1:2", 0, 1),
      },
    ]);
    assert.equal(
      (await element(engine, "sld1:2")).text,
      "Lead\nMid\u000b\u000bEnd\u000b\n",
    );
    assert.ok((await body()).startsWith(`<a:p>${RUN("Lead")}${END}</a:p>`));
    // Replacing two breaks by one: the runs around them keep their bytes.
    await run(engine, [
      {
        op: "replaceText",
        target: "sld1:2",
        text: "\u000b",
        range: RANGE("sld1:2", 8, 10),
      },
    ]);
    assert.equal(
      (await element(engine, "sld1:2")).text,
      "Lead\nMid\u000bEnd\u000b\n",
    );
    assert.ok(
      (await body()).includes(
        `<a:p>${RUN("Mid")}${BR}${RUN("End")}${BR}${END}</a:p>`,
      ),
    );
    // A style over a break writes its properties into the a:br; the break
    // after the range and the paragraph end keep theirs.
    await run(engine, [
      {
        op: "setTextStyle",
        target: "sld1:2",
        range: RANGE("sld1:2", 5, 12),
        style: { bold: true },
      },
    ]);
    assert.ok(
      (await body()).includes(
        '<a:p><a:r><a:rPr lang="en-US" b="1"/><a:t>Mid</a:t></a:r><a:br><a:rPr lang="en-US" b="1"/></a:br><a:r><a:rPr lang="en-US" b="1"/><a:t>End</a:t></a:r>' +
          `${BR}${END}</a:p>`,
      ),
      await body(),
    );
    // A range that covers exactly one break replaces it with text in the
    // style of the run before it.
    await run(engine, [
      {
        op: "replaceText",
        target: "sld1:2",
        text: "-",
        range: RANGE("sld1:2", 8, 9),
      },
    ]);
    assert.equal(
      (await element(engine, "sld1:2")).text,
      "Lead\nMid-End\u000b\n",
    );
    assert.ok(
      (await body()).includes(
        '<a:r><a:rPr lang="en-US" b="1"/><a:t>Mid</a:t></a:r><a:r><a:rPr lang="en-US" b="1"/><a:t>-</a:t></a:r><a:r><a:rPr lang="en-US" b="1"/><a:t>End</a:t></a:r>',
      ),
    );
    await engine.dispose();
  });

  it("styles a collapsed range: no run changes, the paragraph end at a caret at its end, alignment of the caret's paragraph", async () => {
    const engine = await open(
      buildDeck({
        slides: [
          {
            shapes: [
              textShape({
                id: 2,
                x: 0,
                y: 0,
                cx: 914400,
                cy: 914400,
                paragraphs: [
                  ["Hello ", { text: "world", rPr: 'b="1"' }],
                  ["Second"],
                ],
              }),
            ],
          },
        ],
      }),
    );
    const body = async (): Promise<string> =>
      /<p:txBody>.*<\/p:txBody>/s.exec(await partText(engine, SLIDE))![0];
    const before = await body();
    await run(engine, [
      {
        op: "setTextStyle",
        target: "sld1:2",
        range: RANGE("sld1:2", 3, 3),
        style: { italic: true },
      },
    ]);
    assert.equal(await body(), before, "a caret inside a run styles nothing");
    assert.equal((await element(engine, "sld1:2")).textStyle?.italic, false);

    await run(engine, [
      {
        op: "setTextStyle",
        target: "sld1:2",
        range: RANGE("sld1:2", 11, 11),
        style: { italic: true },
      },
    ]);
    const atEnd = await body();
    assert.ok(
      atEnd.includes(
        '<a:r><a:rPr lang="en-US"/><a:t>Hello </a:t></a:r><a:r><a:rPr lang="en-US" b="1"/><a:t>world</a:t></a:r><a:endParaRPr lang="en-US" i="1"/></a:p>',
      ),
      atEnd,
    );
    assert.ok(
      atEnd.includes(
        '<a:p><a:r><a:rPr lang="en-US"/><a:t>Second</a:t></a:r><a:endParaRPr lang="en-US"/></a:p>',
      ),
      "the next paragraph is untouched",
    );

    await run(engine, [
      {
        op: "setTextStyle",
        target: "sld1:2",
        range: RANGE("sld1:2", 14, 14),
        style: { align: "center" },
      },
    ]);
    const aligned = await body();
    assert.equal(aligned.match(/<a:pPr algn="ctr"\/>/g)!.length, 1);
    assert.ok(
      aligned.includes(
        '<a:p><a:pPr algn="ctr"/><a:r><a:rPr lang="en-US"/><a:t>Second</a:t></a:r>',
      ),
    );
    assert.ok(
      aligned.startsWith(
        '<p:txBody><a:bodyPr wrap="square" rtlCol="0"></a:bodyPr><a:lstStyle/><a:p><a:r>',
      ),
    );
    assert.equal((await element(engine, "sld1:2")).textStyle?.align, "left");
    await engine.dispose();
  });

  it("surfaces malformed and non-UTF-8 slide parts by their package codes and keeps the deck editable elsewhere", async () => {
    const shape = (text: string) =>
      textShape({
        id: 2,
        x: 0,
        y: 0,
        cx: 914400,
        cy: 914400,
        paragraphs: [[text]],
      });
    const clean = buildDeck({
      slides: [
        { shapes: [shape("Broken")] },
        { shapes: [shape("Fine")] },
        { shapes: [shape("Latin#")] },
      ],
    });
    const slide1 = (await partOf(clean, SLIDE)).replace("</p:sp>", "");
    const slide3 = new TextEncoder().encode(
      await partOf(clean, "/ppt/slides/slide3.xml"),
    );
    slide3[slide3.indexOf(0x23)] = 0xe9; // "#" → a lone Latin-1 é
    const deck = await withParts(clean, {
      [SLIDE]: slide1,
      "/ppt/slides/slide3.xml": slide3,
    });

    const engine = await open(deck);
    assert.equal(engine.pageCount, 3);
    assert.deepEqual(
      (await engine.getElements({ pageIndex: 1 }, signal)).map(
        (item) => item.text,
      ),
      ["Fine"],
    );
    await assert.rejects(engine.getElements({ pageIndex: 0 }, signal), {
      code: "malformed-xml",
    });
    await assert.rejects(engine.getElements({ pageIndex: 2 }, signal), {
      code: "unsupported-part",
    });
    await assert.rejects(engine.getElement("sld3:2", signal), {
      code: "unsupported-part",
    });
    await assert.rejects(
      check(engine, [{ op: "replaceText", target: "sld1:2", text: "x" }]),
      { code: "malformed-xml" },
    );
    // Validation does not read the slide a text box goes to; applying does,
    // and a failed batch leaves the package as it was.
    const box: PptxOperation = {
      op: "insertTextBox",
      pageIndex: 2,
      rect: { x: 0, y: 0, width: 10, height: 10 },
      text: "x",
    };
    assert.deepEqual(await check(engine, [box]), []);
    await assert.rejects(run(engine, [box]), { code: "unsupported-part" });
    assert.equal(engine.package.revision, 0);
    assert.deepEqual(await engine.materialize("save", {}, signal), deck);
    await engine.dispose();

    const { session, end } = await pptxSession(deck);
    try {
      await assert.rejects(session.getElements({ pageIndex: 0 }), {
        code: "malformed-xml",
      });
      await assert.rejects(
        session.replaceText({ target: "sld3:2", text: "x" }),
        { code: "unsupported-part" },
      );
      assert.equal(session.state.revision, 0);
      const receipt = await session.replaceText({
        target: "sld2:2",
        text: "Still fine",
      });
      assert.equal(receipt.revision, 1);
      assert.equal(
        (await session.getElement("sld2:2")).item?.text,
        "Still fine",
      );
    } finally {
      await end();
    }
  });

  it("reads a deck without masters or layouts, and inserts the first slide of an empty deck", async () => {
    const built = buildDeck({
      slides: [
        {
          layout: 2,
          shapes: [
            textShape({
              id: 2,
              inherit: true,
              placeholder: { type: "ctrTitle" },
              x: 0,
              y: 0,
              cx: 0,
              cy: 0,
              paragraphs: [["Title"]],
            }),
            textShape({
              id: 3,
              x: 914400,
              y: 914400,
              cx: 914400,
              cy: 914400,
              paragraphs: [["Box"]],
            }),
          ],
        },
      ],
    });
    const presentation = await partOf(built, PRESENTATION);
    assert.ok(presentation.includes("<p:sldMasterIdLst>"));
    const engine = await open(
      await withParts(built, {
        [PRESENTATION]: presentation.replace(
          /<p:sldMasterIdLst>.*?<\/p:sldMasterIdLst>/s,
          "",
        ),
      }),
    );
    assert.deepEqual(await engine.layouts(signal), []);
    assert.deepEqual(await engine.slides(signal), [
      { pageIndex: 0, key: "sld1", layout: "", hidden: false },
    ]);
    const title = await element(engine, "sld1:2");
    assert.deepEqual(title.bounds, { x: 0, y: 0, width: 0, height: 0 });
    assert.equal(title.frame, undefined);
    assert.equal(title.rotation, undefined);
    // Nothing to inherit from: the presentation defaults and the stock theme faces.
    assert.deepEqual(title.textStyle, {
      fontFamily: "Calibri",
      fontSize: 18,
      bold: false,
      italic: false,
      underline: false,
      color: { theme: "tx1" },
      align: "left",
    });
    sameRect(
      (await element(engine, "sld1:3")).bounds,
      px({ x: 914400, y: 914400, cx: 914400, cy: 914400 }),
    );
    const issues = await check(engine, [
      { op: "moveElement", target: "sld1:2", by: { dx: 1, dy: 1 } },
      {
        op: "resizeElement",
        target: "sld1:2",
        rect: { x: 0, y: 0, width: 1, height: 1 },
      },
      { op: "moveElement", target: "sld1:3", by: { dx: 1, dy: 1 } },
      { op: "replaceText", target: "sld1:2", text: "Still editable" },
    ]);
    assert.deepEqual(
      issues.map((issue) => [issue.operationIndex, issue.path, issue.code]),
      [
        [0, "/target", "invalid-target"],
        [1, "/target", "invalid-target"],
      ],
    );
    await run(engine, [
      { op: "replaceText", target: "sld1:2", text: "Still editable" },
    ]);
    assert.equal((await element(engine, "sld1:2")).text, "Still editable");
    await engine.dispose();

    // A deck whose slide list is empty: the first slide goes at index 0.
    const empty = await open(buildDeck({ slides: [] }));
    assert.equal(empty.pageCount, 0);
    assert.deepEqual(await empty.getElements({}, signal), []);
    assert.deepEqual(await empty.findText("x", {}, signal), []);
    const change = await run(empty, [{ op: "insertSlide", index: 0 }]);
    assert.deepEqual(change.createdIds, ["sld1:2", "sld1:3"]);
    assert.equal(change.pageCount, 1);
    assert.ok(
      /<p:sldIdLst><p:sldId id="256" r:id="rId3"\/><\/p:sldIdLst>/.test(
        await partText(empty, PRESENTATION),
      ),
    );
    assert.deepEqual(
      (await empty.slides(signal)).map((slide) => [slide.key, slide.layout]),
      [["sld1", "layout1"]],
    );
    await empty.dispose();
  });

  it("refuses insertSlide at validation when the deck has no layout and when the presentation has no p:sldIdLst (bug: both pass validation and apply throws a plain Error)", async () => {
    const noLayouts = buildDeck({ slides: [{ shapes: [] }] });
    const noLayoutsEngine = await open(
      await withParts(noLayouts, {
        [PRESENTATION]: (await partOf(noLayouts, PRESENTATION)).replace(
          /<p:sldMasterIdLst>.*?<\/p:sldMasterIdLst>/s,
          "",
        ),
      }),
    );
    const issues = await check(noLayoutsEngine, [
      { op: "insertSlide", index: 1 },
    ]);
    assert.deepEqual(
      issues.map((issue) => [issue.path, issue.code]),
      [["/layout", "unknown-layout"]],
    );
    await noLayoutsEngine.dispose();

    // Without p:sldIdLst the engine opens the deck with no slides but cannot
    // add the first one: insertSlide should create the list before p:sldSz.
    const noList = buildDeck({ slides: [] });
    const noListEngine = await open(
      await withParts(noList, {
        [PRESENTATION]: (await partOf(noList, PRESENTATION)).replace(
          "<p:sldIdLst></p:sldIdLst>",
          "",
        ),
      }),
    );
    assert.equal(noListEngine.pageCount, 0);
    const change = await run(noListEngine, [{ op: "insertSlide", index: 0 }]);
    assert.equal(change.pageCount, 1);
    assert.ok(
      /<p:sldIdLst><p:sldId id="256" r:id="rId3"\/><\/p:sldIdLst><p:sldSz/.test(
        await partText(noListEngine, PRESENTATION),
      ),
    );
    await noListEngine.dispose();
  });

  it("allocates ids from the slide alone, relates a JPEG asset once and builds a one-cell table through the session", async () => {
    // Layout 1 holds cNvPr ids 2 and 3; the slide holds only 2, so the next
    // id is 3 (ECMA-376 scopes ids per part; the renderer prefers the slide).
    const bytes = buildDeck({
      slides: [{ shapes: [textShape({ id: 2, x: 0, y: 0, cx: 10, cy: 10 })] }],
    });
    const { session, end } = await pptxSession(bytes);
    try {
      const jpeg = tinyJpeg();
      const asset = await session.addAsset(jpeg, { mimeType: "image/jpeg" });
      const receipt = await session.apply([
        {
          op: "insertTextBox",
          pageIndex: 0,
          rect: { x: 0, y: 0, width: 10, height: 10 },
          text: "Box",
        },
        {
          op: "insertImage",
          pageIndex: 0,
          rect: { x: 0, y: 0, width: 16, height: 8 },
          data: asset,
          mimeType: "image/jpeg",
        },
        {
          op: "insertTable",
          pageIndex: 0,
          rect: { x: 10, y: 20, width: 300, height: 40 },
          rows: [["only"]],
        },
        {
          op: "setTableCell",
          target: "$2",
          row: 0,
          column: 0,
          text: "one\u000btwo",
        },
        {
          op: "insertImage",
          pageIndex: 0,
          rect: { x: 50, y: 50, width: 16, height: 8 },
          data: jpeg,
          mimeType: "image/jpeg",
        },
      ]);
      assert.deepEqual(receipt.createdIds, [
        "sld1:3",
        "sld1:4",
        "sld1:5",
        "sld1:6",
      ]);
      const saved = (await session.save()).bytes;
      const pkg = await OoxmlPackage.open(saved, { limits });
      assert.deepEqual(
        pkg.currentPartNames.filter((name) => name.startsWith("/ppt/media/")),
        ["/ppt/media/image1.jpeg"],
        "the same JPEG, as an asset and inline, is stored once",
      );
      assert.deepEqual(await pkg.part("/ppt/media/image1.jpeg"), jpeg);
      const rels = await partOf(saved, "/ppt/slides/_rels/slide1.xml.rels");
      assert.equal(
        rels.match(/Target="\.\.\/media\/image1\.jpeg"/g)!.length,
        2,
        "each picture has its own relationship",
      );
      const xml = await partOf(saved, SLIDE);
      assert.ok(
        xml.includes(
          '<a:tblPr firstRow="1" bandRow="1"/><a:tblGrid><a:gridCol w="2857500"/></a:tblGrid><a:tr h="381000"><a:tc><a:txBody><a:bodyPr/><a:lstStyle/><a:p><a:r><a:t>one</a:t></a:r><a:br/><a:r><a:t>two</a:t></a:r></a:p></a:txBody><a:tcPr/></a:tc></a:tr></a:tbl>',
        ),
        xml,
      );
      const grid = (await session.getElement("sld1:5")).item!;
      assert.deepEqual(grid.table, { rows: [["one\u000btwo"]] });
      assert.deepEqual(grid.bounds, { x: 10, y: 20, width: 300, height: 40 });
      const pictures = (await session.getElements({ kinds: ["image"] })).items;
      assert.deepEqual(
        pictures.map((item) => [item.id, item.name]),
        [
          ["sld1:4", "Picture 3"],
          ["sld1:6", "Picture 5"],
        ],
      );
      await session.undo();
      assert.deepEqual((await session.save()).bytes, bytes);
    } finally {
      await end();
    }
  });

  it("copies external hyperlinks when duplicating, drops them with their shape, and deletes a slide whose notes share media", async () => {
    const png = new Uint8Array([0x89, 0x50, 0x4e, 0x47, 0, 0, 0, 1]);
    const engine = await open(
      buildDeck({
        parts: [
          { name: "ppt/media/image1.png", data: png },
          {
            name: "ppt/notesSlides/notesSlide1.xml",
            data: `${XML}<p:notes xmlns:p="http://schemas.openxmlformats.org/presentationml/2006/main"/>`,
            contentType:
              "application/vnd.openxmlformats-officedocument.presentationml.notesSlide+xml",
          },
          {
            name: "ppt/notesSlides/_rels/notesSlide1.xml.rels",
            data: `${XML}<Relationships xmlns="http://schemas.openxmlformats.org/package/2006/relationships"><Relationship Id="rId1" Type="${REL}slide" Target="../slides/slide1.xml"/><Relationship Id="rId2" Type="${REL}image" Target="../media/image1.png"/></Relationships>`,
          },
        ],
        slides: [
          {
            relationships: [
              {
                id: "rId2",
                type: `${REL}hyperlink`,
                target: "https://example.com/",
              },
              {
                id: "rId3",
                type: `${REL}notesSlide`,
                target: "../notesSlides/notesSlide1.xml",
              },
            ],
            shapes: [
              textShape({
                id: 2,
                x: 0,
                y: 0,
                cx: 914400,
                cy: 914400,
                paragraphs: [["Link"]],
              }).replace(
                '<a:rPr lang="en-US"/><a:t>Link</a:t>',
                '<a:rPr lang="en-US"><a:hlinkClick r:id="rId2"/></a:rPr><a:t>Link</a:t>',
              ),
            ],
          },
          {
            relationships: [
              {
                id: "rId2",
                type: `${REL}image`,
                target: "../media/image1.png",
              },
            ],
            shapes: [
              picture({
                id: 2,
                rId: "rId2",
                x: 0,
                y: 0,
                cx: 914400,
                cy: 914400,
              }),
            ],
          },
        ],
      }),
    );
    // The builder writes relationships without TargetMode; make the link external.
    const original = await engine.materialize("save", {}, signal);
    const rels1 = (
      await partOf(original, "/ppt/slides/_rels/slide1.xml.rels")
    ).replace(
      'Target="https://example.com/"/>',
      'Target="https://example.com/" TargetMode="External"/>',
    );
    assert.ok(rels1.includes('TargetMode="External"'));
    const linked = await open(
      await withParts(original, { "/ppt/slides/_rels/slide1.xml.rels": rels1 }),
    );
    await engine.dispose();

    const copy = await run(linked, [{ op: "duplicateSlide", pageIndex: 0 }]);
    assert.deepEqual(copy.createdIds, ["sld3:2"]);
    assert.deepEqual(copy.warnings, []);
    let saved = await linked.materialize("save", {}, signal);
    const rels3 = await partOf(saved, "/ppt/slides/_rels/slide3.xml.rels");
    assert.ok(
      rels3.includes(
        `<Relationship Id="rId2" Type="${REL}hyperlink" Target="https://example.com/" TargetMode="External"/>`,
      ),
      rels3,
    );
    assert.ok(!rels3.includes("notesSlide"));
    assert.deepEqual(
      [...linked.package.changedParts]
        .filter((name) => !name.includes("slide3"))
        .sort(),
      [
        "/[Content_Types].xml",
        "/ppt/_rels/presentation.xml.rels",
        "/ppt/presentation.xml",
      ],
      "nothing but the copy, its rels and the registrations changed",
    );

    const removed = await run(linked, [
      { op: "deleteElement", target: "sld1:2" },
    ]);
    assert.deepEqual(removed.removedIds, ["sld1:2"]);
    saved = await linked.materialize("save", {}, signal);
    const afterDelete = await partOf(
      saved,
      "/ppt/slides/_rels/slide1.xml.rels",
    );
    assert.ok(!afterDelete.includes('Id="rId2"'), "the hyperlink is gone");
    assert.ok(afterDelete.includes('Id="rId3"'), "the notes stay");
    assert.ok(
      (await partOf(saved, "/ppt/slides/_rels/slide3.xml.rels")).includes(
        'Id="rId2"',
      ),
      "the copy keeps its own link",
    );

    const gone = await run(linked, [{ op: "deleteSlide", pageIndex: 0 }]);
    assert.deepEqual(gone.removedIds, []);
    assert.deepEqual(gone.changedPages, [0, 1]);
    assert.deepEqual(gone.warnings, []);
    const pkg = await OoxmlPackage.open(
      await linked.materialize("save", {}, signal),
      { limits },
    );
    for (const name of [
      "/ppt/slides/slide1.xml",
      "/ppt/notesSlides/notesSlide1.xml",
      "/ppt/notesSlides/_rels/notesSlide1.xml.rels",
    ])
      assert.ok(!pkg.has(name), `${name} is gone`);
    assert.ok(pkg.has("/ppt/media/image1.png"), "the shared media stays");
    assert.deepEqual(
      (await linked.slides(signal)).map((slide) => slide.key),
      ["sld3", "sld2"],
    );
    const remaining = await linked.getElements({ pageIndex: 1 }, signal);
    assert.deepEqual(
      remaining.map((item) => [item.id, item.kind]),
      [["sld2:2", "image"]],
    );
    await linked.dispose();
  });

  it("commits a same-index moveSlide as a batch that changes no page and saves identical bytes", async () => {
    const bytes = buildDeck({ slides: [{ shapes: [] }, { shapes: [] }] });
    const { session, end } = await pptxSession(bytes);
    try {
      const receipt = await session.moveSlide({ from: 1, to: 1 });
      assert.deepEqual(receipt.changedPages, []);
      assert.equal(receipt.operationCount, 1);
      // Every committed batch is a history entry (edit-core, apply step 8).
      assert.equal(receipt.revision, 1);
      assert.equal(session.state.canUndo, true);
      assert.deepEqual((await session.save()).bytes, bytes);
      assert.deepEqual(
        (await session.getSlides()).items.map((slide) => slide.key),
        ["sld1", "sld2"],
      );
      const undone = await session.undo();
      assert.equal(undone.operationCount, 1);
      assert.equal(undone.revision, 2);
      assert.deepEqual((await session.save()).bytes, bytes);
    } finally {
      await end();
    }
  });

  it("restores from a checkpoint base in the middle of a history, on the engine and through session checkpoints", async () => {
    const bytes = buildDeck({
      slides: [
        {
          shapes: [
            textShape({
              id: 2,
              x: 0,
              y: 0,
              cx: 914400,
              cy: 914400,
              paragraphs: [["Alpha"]],
            }),
          ],
        },
      ],
    });
    const engine = await open(bytes);
    const first = await run(engine, [
      {
        op: "insertTextBox",
        pageIndex: 0,
        rect: { x: 10, y: 10, width: 100, height: 20 },
        text: "One",
      },
    ]);
    assert.deepEqual(first.createdIds, ["sld1:3"]);
    const base = await engine.materialize("save", {}, signal);
    const later: PptxOperation[][] = [
      [{ op: "replaceText", target: "sld1:3", text: "Two" }],
      [{ op: "moveElement", target: "sld1:3", by: { dx: 10, dy: 0 } }],
    ];
    for (const operations of later) await run(engine, operations);
    const linearSlide = await partText(engine, SLIDE);
    const linearElements = await engine.getElements({ pageIndex: 0 }, signal);

    await engine.restore(
      {
        base,
        batches: later.map((operations, index) => ({
          stateId: index + 2,
          operations,
        })),
      },
      signal,
    );
    assert.equal(await partText(engine, SLIDE), linearSlide);
    assert.deepEqual(
      await engine.getElements({ pageIndex: 0 }, signal),
      linearElements,
    );
    assert.equal((await element(engine, "sld1:3")).text, "Two");
    await engine.restore({ base, batches: [] }, signal);
    assert.deepEqual(await engine.materialize("save", {}, signal), base);
    assert.equal((await element(engine, "sld1:3")).text, "One");
    // A base that is not a package leaves the engine at its current state.
    await assert.rejects(
      engine.restore({ base: new Uint8Array([1, 2, 3]), batches: [] }, signal),
    );
    assert.equal((await element(engine, "sld1:3")).text, "One");
    await engine.dispose();

    // With a four-entry history every committed state is a checkpoint, so
    // undo and redo restore from bases, not from the original plus replays.
    const { session, end } = await sessionWith(bytes, {
      ...limits,
      maxEditHistory: 4,
    });
    try {
      const saves: Uint8Array[] = [];
      for (const text of ["A", "B", "C"]) {
        await session.replaceText({ target: "sld1:2", text });
        saves.push((await session.save()).bytes);
      }
      await session.undo();
      assert.deepEqual((await session.save()).bytes, saves[1]);
      await session.undo();
      assert.deepEqual((await session.save()).bytes, saves[0]);
      await session.redo();
      assert.deepEqual((await session.save()).bytes, saves[1]);
      assert.equal((await session.getElement("sld1:2")).item?.text, "B");
      // A new batch after an undo drops the redo tail (C) and its
      // checkpoint; the history is now A, B, D.
      await session.replaceText({ target: "sld1:2", text: "D" });
      assert.equal(session.state.canRedo, false);
      const d = (await session.save()).bytes;
      await session.undo();
      assert.deepEqual((await session.save()).bytes, saves[1]);
      await session.redo();
      assert.deepEqual((await session.save()).bytes, d);
      await session.undo();
      await session.undo();
      assert.deepEqual((await session.save()).bytes, saves[0]);
      // State 0 has no checkpoint: the original is reopened without replays.
      await session.undo();
      assert.deepEqual((await session.save()).bytes, bytes);
      assert.equal(session.state.canUndo, false);
      assert.equal((await session.getElement("sld1:2")).item?.text, "Alpha");
    } finally {
      await end();
    }
  });

  it("keeps the handler consistent when a reopen fails (bug: edit-open disposes the live engine before opening the next, so an aborted open leaves a disposed engine as the state)", async () => {
    const handler = createOoxmlEditHandler();
    const context = {
      signal,
      reportProgress: () => {},
      reportWarning: () => {},
    };
    const deck = buildDeck({ slides: [{ shapes: [] }, { shapes: [] }] });
    const payload = () => ({
      data: deck.slice().buffer,
      limits,
      format: "pptx",
    });
    assert.deepEqual(await handler("edit-open", payload(), context), {
      pageCount: 2,
    });
    await assert.rejects(
      handler("edit-open", payload(), {
        ...context,
        signal: AbortSignal.abort(),
      }),
      { code: "aborted" },
    );
    // The previous deck is still served: the next engine is opened before
    // the live one is disposed.
    assert.equal(
      ((await handler("edit-pptx-slides", undefined, context)) as unknown[])
        .length,
      2,
    );
  });

  it("serves the worker handler's open, reopen and dispose cycle, refuses other formats, and honours abort signals", async () => {
    const deckA = buildDeck({ slides: [{ shapes: [] }, { shapes: [] }] });
    const deckB = buildDeck({
      slides: [{ shapes: [textShape({ id: 2, x: 0, y: 0, cx: 10, cy: 10 })] }],
    });
    const handler = createOoxmlEditHandler();
    const context = {
      signal,
      reportProgress: () => {},
      reportWarning: () => {},
    };
    const call = (operation: string, payload?: unknown, own = context) =>
      handler(operation as WorkerOperation, payload, own);
    const openPayload = (deck: Uint8Array, format = "pptx") => ({
      data: deck.slice().buffer,
      limits,
      format,
    });

    assert.equal(await call("edit-init"), undefined);
    await assert.rejects(call("edit-elements", { query: {} }), {
      code: "lifecycle-error",
    });
    await assert.rejects(call("edit-open", openPayload(deckA, "pdf")), {
      code: "edit-unsupported",
      details: { format: "pdf", reason: "no-engine" },
    });
    assert.deepEqual(await call("edit-open", openPayload(deckA)), {
      pageCount: 2,
    });
    assert.equal(((await call("edit-pptx-slides")) as unknown[]).length, 2);
    // Opening again replaces the engine without a dispose in between.
    assert.deepEqual(await call("edit-open", openPayload(deckB)), {
      pageCount: 1,
    });
    assert.deepEqual(
      ((await call("edit-elements", { query: {} })) as PptxElement[]).map(
        (item) => item.id,
      ),
      ["sld1:2"],
    );
    assert.equal(await call("edit-dispose"), undefined);
    await assert.rejects(
      call("edit-apply", { batch: { stateId: 1, operations: [] } }),
      { code: "lifecycle-error" },
    );
    assert.equal(await call("edit-dispose"), undefined, "idempotent");
    assert.deepEqual(await call("edit-open", openPayload(deckA)), {
      pageCount: 2,
    });
    await assert.rejects(call("edit-nope"), { code: "internal" });
    await call("edit-dispose");
    // An aborted open on an empty handler leaves it empty.
    await assert.rejects(
      call("edit-open", openPayload(deckB), {
        ...context,
        signal: AbortSignal.abort(),
      }),
      { code: "aborted" },
    );
    await assert.rejects(call("edit-pptx-slides"), {
      code: "lifecycle-error",
      message: "No package is open for editing",
    });

    // The engine: an aborted open, and an aborted batch that changes nothing.
    const aborted = AbortSignal.abort();
    await assert.rejects(PptxEditEngine.open(deckB, limits, aborted), {
      code: "aborted",
    });
    const engine = await open(deckB);
    const batch: PptxOperation[] = [
      {
        op: "insertTextBox",
        pageIndex: 0,
        rect: { x: 0, y: 0, width: 10, height: 10 },
        text: "x",
      },
    ];
    await assert.rejects(engine.apply(batch, aborted), { code: "aborted" });
    assert.equal(engine.package.revision, 0);
    assert.deepEqual(await engine.materialize("save", {}, signal), deckB);
    await engine.dispose();
    await assert.rejects(engine.getElements({}, signal), {
      code: "lifecycle-error",
    });

    // The session: a call with an aborted signal rejects before it is queued.
    const { session, end } = await pptxSession(deckB);
    try {
      await assert.rejects(
        session.replaceText(
          { target: "sld1:2", text: "x" },
          { signal: aborted },
        ),
        { code: "aborted" },
      );
      await assert.rejects(session.getElements({}, { signal: aborted }), {
        code: "aborted",
      });
      assert.equal(session.state.revision, 0);
      assert.deepEqual((await session.save()).bytes, deckB);
    } finally {
      await end();
    }
  });
});
