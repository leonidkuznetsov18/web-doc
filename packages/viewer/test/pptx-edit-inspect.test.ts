import assert from "node:assert/strict";
import { existsSync, readFileSync } from "node:fs";
import { describe, it } from "node:test";
import { pathToFileURL } from "node:url";

import { PptxEditEngine } from "../src/edit/pptx/engine.js";
import { emuToPx } from "../src/edit/pptx/geometry.js";
import { pptxOperationSchemaDrafts } from "../src/edit/pptx/schemas.js";
import { assertSupportedSchema } from "../src/edit/schema.js";
import {
  defaultResourceLimits,
  type EditOperation,
  type PptxElement,
} from "../src/index.js";
import {
  alternateContent,
  buildDeck,
  connector,
  graphicFrame,
  group,
  LAYOUT2_SUBTITLE,
  LAYOUT2_TITLE,
  MASTER_BODY,
  MASTER_TITLE,
  picture,
  syntheticDeck,
  table,
  textShape,
} from "./fixtures/pptx-builder.js";
import { pptxSession } from "./fixtures/pptx-session.js";

/*
 * Task 44 of the PPTX module: the engine opens a deck, lists every element
 * of a slide with frames read from the XML (own, inherited through the
 * placeholder chain, or mapped through groups), text, styles and the
 * operations each accepts; hit-testing and text search; the slide and
 * layout reads; identity without changes; the worker protocol.
 */

const PACKAGE_DIR = pathToFileURL(`${process.cwd()}/`);
const SAMPLE = new URL("../../.cache/corpus/sample.pptx", PACKAGE_DIR);
const signal = new AbortController().signal;

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

async function open(bytes: Uint8Array): Promise<PptxEditEngine> {
  return PptxEditEngine.open(bytes, defaultResourceLimits, signal);
}

describe("PPTX inspection (pptx-edit)", () => {
  it("lists the corpus deck's placeholders with frames and styles inherited from the layout and master", async (t) => {
    if (!existsSync(SAMPLE)) {
      t.skip("no corpus; run npm run corpus:fetch");
      return;
    }
    const engine = await open(new Uint8Array(readFileSync(SAMPLE)));
    assert.equal(engine.pageCount, 2);
    const elements = (await engine.getElements({}, signal)) as PptxElement[];
    assert.deepEqual(
      elements.map((element) => [element.id, element.kind, element.name]),
      [
        ["sld1:2", "shape", "The Title"],
        ["sld1:3", "shape", "Another Subtitle"],
        ["sld2:2", "shape", "Title 1"],
        ["sld2:3", "shape", "Content Placeholder 2"],
      ],
    );
    const [title, subtitle, , content] = elements as [
      PptxElement,
      PptxElement,
      PptxElement,
      PptxElement,
    ];
    // Slide 1 uses the "Title Slide" layout: its ctrTitle and subTitle frames.
    sameRect(
      title.bounds,
      px({ x: 685800, y: 2130425, cx: 7772400, cy: 1470025 }),
    );
    sameRect(
      subtitle.bounds,
      px({ x: 1371600, y: 3886200, cx: 6400800, cy: 1752600 }),
    );
    assert.equal(title.placeholder?.type, "ctrTitle");
    assert.deepEqual(subtitle.placeholder, { type: "subTitle", idx: 1 });
    assert.equal(title.text, "Title of the first slide");
    assert.equal(
      subtitle.text,
      "Subtitle of the first slide\n\nThis bit is in italic green",
    );
    // Title style: master titleStyle (44 pt, centred, +mj-lt → theme major).
    assert.deepEqual(title.textStyle, {
      fontFamily: "Calibri",
      fontSize: 44,
      bold: false,
      italic: false,
      underline: false,
      color: { theme: "tx1" },
      align: "center",
    });
    // Subtitle style: the layout's lstStyle (centred, tinted tx1), size from the master bodyStyle.
    assert.deepEqual(subtitle.textStyle, {
      fontFamily: "Calibri",
      fontSize: 32,
      bold: false,
      italic: false,
      underline: false,
      color: { theme: "tx1", mods: { tint: 75000 } },
      align: "center",
    });
    // Slide 2's content placeholder inherits the master body frame through
    // the "Title and Content" layout, which has no frame of its own.
    sameRect(
      content.bounds,
      px({ x: 457200, y: 1600200, cx: 8229600, cy: 4525963 }),
    );
    assert.equal(content.textStyle?.fontSize, 32);
    assert.equal(content.textStyle?.align, "left");
    assert.deepEqual(content.operations, [
      "replaceText",
      "setTextStyle",
      "setShapeStyle",
      "moveElement",
      "resizeElement",
      "deleteElement",
    ]);
    const layouts = await engine.layouts(signal);
    assert.equal(layouts.length, 11);
    assert.deepEqual(layouts[0], {
      id: "layout1",
      name: "Title Slide",
      type: "title",
      master: "master1",
    });
    const slides = await engine.slides(signal);
    assert.deepEqual(slides, [
      { pageIndex: 0, key: "sld1", layout: "layout1", hidden: false },
      { pageIndex: 1, key: "sld2", layout: "layout2", hidden: false },
    ]);
    await engine.dispose();
  });

  it("reads every element kind of a built deck: frames, rotation, groups, hidden shapes, tables and fallbacks", async () => {
    const png = new Uint8Array([0x89, 0x50, 0x4e, 0x47, 0, 0, 0, 0]);
    const deck = buildDeck({
      parts: [{ name: "ppt/media/image1.png", data: png }],
      slides: [
        {
          layout: 2,
          relationships: [
            {
              id: "rId2",
              type: "http://schemas.openxmlformats.org/officeDocument/2006/relationships/image",
              target: "../media/image1.png",
            },
          ],
          shapes: [
            textShape({
              id: 2,
              name: "Title",
              inherit: true,
              placeholder: { type: "ctrTitle" },
              x: 0,
              y: 0,
              cx: 0,
              cy: 0,
              paragraphs: [["Inherited title"]],
            }),
            textShape({
              id: 3,
              name: "Rotated",
              x: 914400,
              y: 914400,
              cx: 1828800,
              cy: 914400,
              rotation: 90,
              paragraphs: [
                [{ text: "Bold", rPr: 'b="1" sz="2000"' }, " and plain"],
                ["Second", { br: true }, "line"],
              ],
              fill: '<a:solidFill><a:srgbClr val="FF0000"/></a:solidFill>',
              line: '<a:ln w="25400"><a:solidFill><a:schemeClr val="accent1"><a:lumMod val="75000"/></a:schemeClr></a:solidFill></a:ln>',
            }),
            picture({
              id: 4,
              name: "Logo",
              rId: "rId2",
              x: 0,
              y: 0,
              cx: 914400,
              cy: 914400,
            }),
            table({
              id: 5,
              name: "Grid",
              x: 914400,
              y: 2743200,
              cx: 3657600,
              cy: 914400,
              rows: [
                ["A", "B"],
                ["C", "D"],
              ],
            }),
            group({
              id: 6,
              name: "Group",
              x: 4572000,
              y: 2743200,
              cx: 1828800,
              cy: 1828800,
              child: { x: 0, y: 0, cx: 914400, cy: 914400 },
              children: [
                textShape({
                  id: 7,
                  name: "In group",
                  x: 0,
                  y: 0,
                  cx: 457200,
                  cy: 457200,
                  paragraphs: [["Child"]],
                }),
              ],
            }),
            connector({
              id: 8,
              name: "Line",
              x: 0,
              y: 5486400,
              cx: 2743200,
              cy: 0,
            }),
            graphicFrame({
              id: 9,
              name: "Chart",
              x: 5486400,
              y: 914400,
              cx: 2743200,
              cy: 1828800,
              uri: "http://schemas.openxmlformats.org/drawingml/2006/chart",
              inner:
                '<c:chart xmlns:c="http://schemas.openxmlformats.org/drawingml/2006/chart" r:id="rId99"/>',
            }),
            textShape({
              id: 10,
              name: "Hidden",
              hidden: true,
              x: 0,
              y: 0,
              cx: 914400,
              cy: 457200,
              paragraphs: [["Hidden text"]],
            }),
            alternateContent(
              textShape({ id: 11, name: "Choice", x: 0, y: 0, cx: 10, cy: 10 }),
              textShape({
                id: 11,
                name: "Fallback",
                x: 0,
                y: 0,
                cx: 914400,
                cy: 914400,
                paragraphs: [["Fallback text"]],
              }),
            ),
            textShape({
              id: 3,
              name: "Duplicate id",
              x: 0,
              y: 0,
              cx: 914400,
              cy: 914400,
            }),
          ],
        },
      ],
    });
    const engine = await open(deck);
    const elements = (await engine.getElements(
      { pageIndex: 0 },
      signal,
    )) as PptxElement[];
    const byId = new Map(elements.map((element) => [element.id, element]));
    assert.deepEqual(
      [...byId.keys()],
      [
        "sld1:2",
        "sld1:3",
        "sld1:4",
        "sld1:5",
        "sld1:6",
        "sld1:7",
        "sld1:8",
        "sld1:9",
        "sld1:10",
        "sld1:11",
        "sld1:3#2",
      ],
    );
    assert.deepEqual(
      elements.map((element) => element.kind),
      [
        "shape",
        "shape",
        "image",
        "table",
        "group",
        "shape",
        "connector",
        "other",
        "shape",
        "shape",
        "shape",
      ],
    );

    // Inherited through layout 2's ctrTitle.
    sameRect(byId.get("sld1:2")!.bounds, px(LAYOUT2_TITLE));
    assert.ok(
      MASTER_TITLE.x !== LAYOUT2_TITLE.x,
      "the layout overrides the master",
    );

    // A frame rotated by 90°: the bounds swap width and height about the centre.
    const rotated = byId.get("sld1:3")!;
    assert.equal(rotated.rotation, 90);
    assert.deepEqual(rotated.frame, {
      ...px({ x: 914400, y: 914400, cx: 1828800, cy: 914400 }),
      rotation: 90,
      flipH: false,
      flipV: false,
    });
    sameRect(
      rotated.bounds,
      px({ x: 914400 + 457200, y: 914400 - 457200, cx: 914400, cy: 1828800 }),
    );
    assert.equal(rotated.text, "Bold and plain\nSecond\u000bline");
    assert.deepEqual(rotated.textStyle, {
      fontFamily: "Calibri",
      fontSize: 20,
      bold: true,
      italic: false,
      underline: false,
      color: { theme: "tx1" },
      align: "left",
    });
    assert.deepEqual(rotated.shapeStyle, {
      fill: "#FF0000",
      line: { color: { theme: "accent1", mods: { lumMod: 75000 } }, width: 2 },
    });

    assert.deepEqual(byId.get("sld1:4")!.operations, [
      "moveElement",
      "resizeElement",
      "deleteElement",
    ]);
    const grid = byId.get("sld1:5")!;
    assert.deepEqual(grid.table, {
      rows: [
        ["A", "B"],
        ["C", "D"],
      ],
    });
    assert.equal(grid.text, "A\tB\nC\tD");
    assert.deepEqual(grid.operations, [
      "setTableCell",
      "moveElement",
      "resizeElement",
      "deleteElement",
    ]);

    // The group doubles its child space: the child's 457200 box becomes 914400 at the group's origin.
    const child = byId.get("sld1:7")!;
    assert.equal(child.parentId, "sld1:6");
    sameRect(
      child.bounds,
      px({ x: 4572000, y: 2743200, cx: 914400, cy: 914400 }),
    );
    sameRect(
      byId.get("sld1:6")!.bounds,
      px({ x: 4572000, y: 2743200, cx: 1828800, cy: 1828800 }),
    );

    assert.deepEqual(byId.get("sld1:8")!.shapeStyle, {
      line: { color: "#FF0000", width: 1.5 },
    });
    assert.equal(byId.get("sld1:8")!.bounds.height, 0);
    assert.deepEqual(byId.get("sld1:9")!.operations, [
      "moveElement",
      "resizeElement",
      "deleteElement",
    ]);
    assert.equal(byId.get("sld1:10")!.hidden, true);
    assert.ok(byId.get("sld1:10")!.operations.includes("replaceText"));
    const fallback = byId.get("sld1:11")!;
    assert.equal(fallback.name, "Fallback");
    assert.deepEqual(fallback.operations, []);
    assert.equal(byId.get("sld1:3#2")!.name, "Duplicate id");

    assert.equal((await engine.getElement("sld1:7", signal))?.name, "In group");
    assert.equal(await engine.getElement("sld9:7", signal), undefined);
    assert.deepEqual(
      (
        await engine.getElements(
          { pageIndex: 0, kinds: ["image", "table"] },
          signal,
        )
      ).map((element) => element.id),
      ["sld1:4", "sld1:5"],
    );
    assert.deepEqual(
      (
        await engine.getElements(
          {
            pageIndex: 0,
            intersects: { x: 480, y: 288, width: 10, height: 10 },
          },
          signal,
        )
      ).map((element) => element.id),
      ["sld1:2", "sld1:6", "sld1:7"],
    );
    await engine.dispose();
  });

  it("hit-tests topmost first and finds text with ranges", async () => {
    const deck = buildDeck({
      slides: [
        {
          shapes: [
            textShape({
              id: 2,
              name: "Back",
              x: 0,
              y: 0,
              cx: 1828800,
              cy: 1828800,
              paragraphs: [["alpha beta"]],
            }),
            textShape({
              id: 3,
              name: "Front",
              x: 914400,
              y: 914400,
              cx: 1828800,
              cy: 1828800,
              paragraphs: [["Beta gamma"]],
            }),
          ],
        },
        {
          shapes: [
            textShape({
              id: 2,
              name: "Other",
              x: 0,
              y: 0,
              cx: 914400,
              cy: 914400,
              paragraphs: [["beta"]],
            }),
          ],
        },
      ],
    });
    const engine = await open(deck);
    assert.deepEqual(
      (await engine.elementsAt(0, { x: 100, y: 100 }, signal)).map(
        (element) => element.id,
      ),
      ["sld1:3", "sld1:2"],
    );
    assert.deepEqual(
      (await engine.elementsAt(0, { x: 10, y: 10 }, signal)).map(
        (element) => element.id,
      ),
      ["sld1:2"],
    );
    assert.deepEqual(
      await engine.elementsAt(0, { x: 5000, y: 5000 }, signal),
      [],
    );
    assert.deepEqual(await engine.elementsAt(7, { x: 1, y: 1 }, signal), []);

    const found = await engine.findText("beta", {}, signal);
    assert.deepEqual(
      found.map((target) => [
        target.pageIndex,
        target.text,
        target.elementIds[0],
        target.ranges[0]!.start.offset,
      ]),
      [
        [0, "beta", "sld1:2", 6],
        [0, "Beta", "sld1:3", 0],
        [1, "beta", "sld2:2", 0],
      ],
    );
    assert.equal(
      (await engine.findText("beta", { caseSensitive: true }, signal)).length,
      2,
    );
    assert.equal(
      (await engine.findText("beta", { pageRange: [1, 1] }, signal)).length,
      1,
    );
    assert.equal(
      (await engine.findText("beta", { maxResults: 1 }, signal)).length,
      1,
    );
    await engine.dispose();
  });

  it("returns identical bytes without changes, after a restore and for a macro-enabled deck", async () => {
    const deck = syntheticDeck(12);
    const engine = await open(deck);
    assert.equal(engine.pageCount, 12);
    assert.deepEqual(await engine.materialize("save", {}, signal), deck);
    await engine.restore({ batches: [] }, signal);
    assert.deepEqual(await engine.materialize("show", {}, signal), deck);
    await engine.dispose();

    const macro = buildDeck({
      macro: true,
      slides: [{ shapes: [textShape({ id: 2, x: 0, y: 0, cx: 10, cy: 10 })] }],
    });
    const macroEngine = await open(macro);
    assert.deepEqual(await macroEngine.materialize("save", {}, signal), macro);
    await macroEngine.dispose();
  });

  it("serves a session over the worker protocol with slide and layout reads", async () => {
    const deck = buildDeck({
      slides: [
        {
          shapes: [
            textShape({ id: 2, name: "A", x: 0, y: 0, cx: 914400, cy: 914400 }),
          ],
        },
        { layout: 2, hidden: true, shapes: [] },
      ],
    });
    const { session, end } = await pptxSession(deck);
    try {
      assert.equal(session.format, "pptx");
      assert.equal(session.state.pageCount, 2);
      const elements = await session.getElements();
      assert.deepEqual(
        elements.items.map((element) => element.id),
        ["sld1:2"],
      );
      assert.equal(elements.revision, 0);
      const slides = await session.getSlides();
      assert.deepEqual(slides.items, [
        { pageIndex: 0, key: "sld1", layout: "layout1", hidden: false },
        { pageIndex: 1, key: "sld2", layout: "layout2", hidden: true },
      ]);
      const layouts = await session.getLayouts();
      assert.deepEqual(
        layouts.items.map((layout) => [layout.id, layout.name, layout.type]),
        [
          ["layout1", "Title and Content", "obj"],
          ["layout2", "Title Slide", "title"],
        ],
      );
      const saved = await session.save();
      assert.deepEqual(saved.bytes, deck);
      assert.equal(saved.revision, 0);
      const unknown: EditOperation = {
        op: "replaceText",
        ...{ target: "sld1:2", text: "x" },
      };
      await assert.rejects(session.applyJson([unknown]), {
        code: "invalid-operation",
      });
    } finally {
      await end();
    }
  });

  it("declares schemas the in-repo validator supports", () => {
    for (const schema of Object.values(pptxOperationSchemaDrafts))
      assertSupportedSchema(schema);
  });
});
