import assert from "node:assert/strict";
import { describe, it } from "node:test";

import { PptxEditEngine } from "../src/edit/pptx/engine.js";
import { emuToPx } from "../src/edit/pptx/geometry.js";
import { OoxmlPackage } from "../src/edit/ooxml/package.js";
import { defaultResourceLimits, type PptxOperation } from "../src/index.js";
import {
  buildDeck,
  graphicFrame,
  LAYOUT2_TITLE,
  MASTER_BODY,
  MASTER_TITLE,
  picture,
  textShape,
} from "./fixtures/pptx-builder.js";
import { pptxSession } from "./fixtures/pptx-session.js";

/*
 * Task 48 of the PPTX module: a new slide instantiates its layout's
 * placeholders; a duplicate copies the slide and clones the parts only it
 * may own while sharing media and the layout; deletion takes the notes
 * slide along and never removes the last slide; reordering touches only
 * p:sldIdLst.
 */

const signal = new AbortController().signal;
const REL =
  "http://schemas.openxmlformats.org/officeDocument/2006/relationships/";
const XML = '<?xml version="1.0" encoding="UTF-8" standalone="yes"?>\n';

const run = (engine: PptxEditEngine, operations: PptxOperation[]) =>
  engine.apply(operations, signal);
const check = (engine: PptxEditEngine, operations: PptxOperation[]) =>
  engine.validate(operations, signal);

async function open(bytes: Uint8Array): Promise<PptxEditEngine> {
  return PptxEditEngine.open(bytes, defaultResourceLimits, signal);
}

async function saved(engine: PptxEditEngine): Promise<OoxmlPackage> {
  return OoxmlPackage.open(await engine.materialize("save", {}, signal), {
    limits: defaultResourceLimits,
  });
}

async function partText(pkg: OoxmlPackage, name: string): Promise<string> {
  return new TextDecoder().decode(await pkg.part(name));
}

async function keys(engine: PptxEditEngine): Promise<string[]> {
  return (await engine.slides(signal)).map((slide) => slide.key);
}

function px(box: { x: number; y: number; cx: number; cy: number }) {
  return {
    x: emuToPx(box.x),
    y: emuToPx(box.y),
    width: emuToPx(box.cx),
    height: emuToPx(box.cy),
  };
}

describe("PPTX slide operations (pptx-edit)", () => {
  it("inserts slides that instantiate their layout's placeholders", async () => {
    const engine = await open(
      buildDeck({
        slides: [
          { shapes: [textShape({ id: 2, x: 0, y: 0, cx: 10, cy: 10 })] },
          { layout: 2, shapes: [] },
        ],
      }),
    );
    const change = await run(engine, [{ op: "insertSlide", index: 1 }]);
    assert.deepEqual(change.createdIds, ["sld3:2", "sld3:3"]);
    assert.deepEqual(change.changedPages, [1, 2]);
    assert.equal(change.pageCount, 3);
    assert.deepEqual(await keys(engine), ["sld1", "sld3", "sld2"]);
    const elements = await engine.getElements({ pageIndex: 1 }, signal);
    assert.deepEqual(
      elements.map((element) => [
        element.id,
        element.name,
        element.placeholder,
        element.text,
      ]),
      [
        ["sld3:2", "Title 1", { type: "title" }, ""],
        ["sld3:3", "Content Placeholder 2", { type: "body", idx: 1 }, ""],
      ],
    );
    assert.deepEqual(elements[0]!.bounds, px(MASTER_TITLE));
    assert.deepEqual(elements[1]!.bounds, px(MASTER_BODY));
    const pkg = await saved(engine);
    const slide = await partText(pkg, "/ppt/slides/slide3.xml");
    assert.ok(slide.startsWith(`${XML}<p:sld xmlns:a=`));
    assert.ok(
      slide.includes(
        '<p:nvPr><p:ph type="title"/></p:nvPr></p:nvSpPr><p:spPr/><p:txBody><a:bodyPr/><a:lstStyle/><a:p><a:endParaRPr lang="en-US"/></a:p></p:txBody>',
      ),
    );
    assert.ok(slide.includes('<p:ph idx="1"/>'));
    assert.ok(
      (await partText(pkg, "/ppt/slides/_rels/slide3.xml.rels")).includes(
        'Target="../slideLayouts/slideLayout1.xml"',
      ),
    );
    const presentation = await partText(pkg, "/ppt/presentation.xml");
    assert.ok(
      /<p:sldIdLst><p:sldId id="256" r:id="rId3"\/><p:sldId id="258" r:id="rId5"\/><p:sldId id="257" r:id="rId4"\/><\/p:sldIdLst>/.test(
        presentation,
      ),
      presentation,
    );
    assert.ok(
      (await partText(pkg, "/ppt/_rels/presentation.xml.rels")).includes(
        'Id="rId5" Type="' + REL + 'slide" Target="slides/slide3.xml"',
      ),
    );
    assert.ok(
      (await partText(pkg, "/[Content_Types].xml")).includes(
        'PartName="/ppt/slides/slide3.xml"',
      ),
    );

    // An explicit layout at the end: the title slide's placeholders.
    const second = await run(engine, [
      { op: "insertSlide", index: 3, layout: "layout2" },
    ]);
    assert.deepEqual(second.createdIds, ["sld4:2", "sld4:3"]);
    const titleSlide = await engine.getElements({ pageIndex: 3 }, signal);
    assert.deepEqual(titleSlide[0]!.placeholder, { type: "ctrTitle" });
    assert.deepEqual(titleSlide[0]!.bounds, px(LAYOUT2_TITLE));
    const issues = await check(engine, [
      { op: "insertSlide", index: 9 },
      { op: "insertSlide", index: 0, layout: "layout9" },
    ]);
    assert.deepEqual(
      issues.map((issue) => [issue.operationIndex, issue.path, issue.code]),
      [
        [0, "/index", "range"],
        [1, "/layout", "unknown-layout"],
      ],
    );
    await engine.dispose();
  });

  it("duplicates a slide, cloning the parts it owns and sharing media, the layout and no notes", async () => {
    const engine = await open(
      buildDeck({
        parts: [
          {
            name: "ppt/media/image1.png",
            data: new Uint8Array([0x89, 0x50, 0x4e, 0x47, 9]),
          },
          {
            name: "ppt/charts/chart1.xml",
            data: `${XML}<c:chartSpace xmlns:c="http://schemas.openxmlformats.org/drawingml/2006/chart"/>`,
            contentType:
              "application/vnd.openxmlformats-officedocument.drawingml.chart+xml",
          },
          {
            name: "ppt/charts/_rels/chart1.xml.rels",
            data: `${XML}<Relationships xmlns="http://schemas.openxmlformats.org/package/2006/relationships"><Relationship Id="rId1" Type="http://schemas.microsoft.com/office/2011/relationships/chartColorStyle" Target="colors1.xml"/></Relationships>`,
          },
          {
            name: "ppt/charts/colors1.xml",
            data: `${XML}<cs:colorStyle xmlns:cs="http://schemas.microsoft.com/office/drawing/2012/chartStyle"/>`,
            contentType: "application/vnd.ms-office.chartcolorstyle+xml",
          },
          {
            name: "ppt/notesSlides/notesSlide1.xml",
            data: `${XML}<p:notes xmlns:p="http://schemas.openxmlformats.org/presentationml/2006/main"/>`,
            contentType:
              "application/vnd.openxmlformats-officedocument.presentationml.notesSlide+xml",
          },
        ],
        slides: [
          {
            relationships: [
              {
                id: "rId2",
                type: `${REL}image`,
                target: "../media/image1.png",
              },
              {
                id: "rId3",
                type: `${REL}chart`,
                target: "../charts/chart1.xml",
              },
              {
                id: "rId4",
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
                paragraphs: [["Source"]],
              }),
              picture({
                id: 3,
                rId: "rId2",
                x: 0,
                y: 0,
                cx: 914400,
                cy: 914400,
              }),
              graphicFrame({
                id: 4,
                x: 0,
                y: 0,
                cx: 914400,
                cy: 914400,
                uri: "http://schemas.openxmlformats.org/drawingml/2006/chart",
                inner:
                  '<c:chart xmlns:c="http://schemas.openxmlformats.org/drawingml/2006/chart" r:id="rId3"/>',
              }),
            ],
          },
          { shapes: [] },
        ],
      }),
    );
    const change = await run(engine, [{ op: "duplicateSlide", pageIndex: 0 }]);
    assert.deepEqual(change.createdIds, ["sld3:2", "sld3:3", "sld3:4"]);
    assert.deepEqual(change.changedPages, [1, 2]);
    assert.deepEqual(change.warnings, []);
    assert.deepEqual(await keys(engine), ["sld1", "sld3", "sld2"]);
    const copy = await engine.getElements({ pageIndex: 1 }, signal);
    assert.deepEqual(
      copy.map((element) => [element.kind, element.text]),
      [
        ["shape", "Source"],
        ["image", undefined],
        ["other", undefined],
      ],
    );
    const pkg = await saved(engine);
    assert.deepEqual(
      await pkg.part("/ppt/slides/slide3.xml"),
      await pkg.part("/ppt/slides/slide1.xml"),
    );
    const rels = await partText(pkg, "/ppt/slides/_rels/slide3.xml.rels");
    assert.ok(
      rels.includes(
        'Id="rId1" Type="' +
          REL +
          'slideLayout" Target="../slideLayouts/slideLayout1.xml"',
      ),
    );
    assert.ok(
      rels.includes(
        'Id="rId2" Type="' + REL + 'image" Target="../media/image1.png"',
      ),
      "media is shared",
    );
    assert.ok(
      rels.includes(
        'Id="rId3" Type="' + REL + 'chart" Target="../charts/chart2.xml"',
      ),
      "the chart is cloned",
    );
    assert.ok(!rels.includes("notesSlide"), "notes are not copied");
    assert.deepEqual(
      await pkg.part("/ppt/charts/chart2.xml"),
      await pkg.part("/ppt/charts/chart1.xml"),
    );
    assert.ok(
      (await partText(pkg, "/ppt/charts/_rels/chart2.xml.rels")).includes(
        'Target="colors2.xml"',
      ),
    );
    assert.deepEqual(
      await pkg.part("/ppt/charts/colors2.xml"),
      await pkg.part("/ppt/charts/colors1.xml"),
    );
    const media = pkg.currentPartNames.filter((name) =>
      name.startsWith("/ppt/media/"),
    );
    assert.deepEqual(media, ["/ppt/media/image1.png"]);
    const types = await partText(pkg, "/[Content_Types].xml");
    assert.ok(
      types.includes('PartName="/ppt/charts/chart2.xml"') &&
        types.includes('PartName="/ppt/charts/colors2.xml"'),
    );
    const atEnd = await run(engine, [
      { op: "duplicateSlide", pageIndex: 2, index: 0 },
    ]);
    assert.deepEqual(atEnd.createdIds, []);
    assert.deepEqual(await keys(engine), ["sld4", "sld1", "sld3", "sld2"]);
    await engine.dispose();
  });

  it("deletes slides with their notes and keeps the last one", async () => {
    const engine = await open(
      buildDeck({
        parts: [
          {
            name: "ppt/media/image1.png",
            data: new Uint8Array([0x89, 0x50, 0x4e, 0x47, 9]),
          },
          {
            name: "ppt/notesSlides/notesSlide1.xml",
            data: `${XML}<p:notes xmlns:p="http://schemas.openxmlformats.org/presentationml/2006/main"/>`,
            contentType:
              "application/vnd.openxmlformats-officedocument.presentationml.notesSlide+xml",
          },
          {
            name: "ppt/notesSlides/_rels/notesSlide1.xml.rels",
            data: `${XML}<Relationships xmlns="http://schemas.openxmlformats.org/package/2006/relationships"><Relationship Id="rId1" Type="${REL}slide" Target="../slides/slide1.xml"/></Relationships>`,
          },
        ],
        slides: [
          {
            relationships: [
              {
                id: "rId2",
                type: `${REL}image`,
                target: "../media/image1.png",
              },
              {
                id: "rId3",
                type: `${REL}notesSlide`,
                target: "../notesSlides/notesSlide1.xml",
              },
            ],
            shapes: [
              textShape({ id: 2, x: 0, y: 0, cx: 10, cy: 10 }),
              picture({ id: 3, rId: "rId2", x: 0, y: 0, cx: 10, cy: 10 }),
            ],
          },
          { shapes: [textShape({ id: 2, x: 0, y: 0, cx: 10, cy: 10 })] },
        ],
      }),
    );
    const change = await run(engine, [{ op: "deleteSlide", pageIndex: 0 }]);
    assert.deepEqual(change.removedIds, ["sld1:2", "sld1:3"]);
    assert.deepEqual(change.changedPages, [0]);
    assert.equal(change.pageCount, 1);
    assert.deepEqual(change.warnings, []);
    assert.deepEqual(await keys(engine), ["sld2"]);
    const pkg = await saved(engine);
    for (const name of [
      "/ppt/slides/slide1.xml",
      "/ppt/slides/_rels/slide1.xml.rels",
      "/ppt/notesSlides/notesSlide1.xml",
      "/ppt/notesSlides/_rels/notesSlide1.xml.rels",
    ])
      assert.ok(!pkg.has(name), `${name} is gone`);
    assert.ok(pkg.has("/ppt/media/image1.png"), "media stays");
    const types = await partText(pkg, "/[Content_Types].xml");
    assert.ok(
      !types.includes("slide1.xml") && !types.includes("notesSlide1.xml"),
    );
    assert.ok(
      !(await partText(pkg, "/ppt/_rels/presentation.xml.rels")).includes(
        "slide1.xml",
      ),
    );
    assert.ok(
      /<p:sldIdLst><p:sldId id="257" r:id="rId4"\/><\/p:sldIdLst>/.test(
        await partText(pkg, "/ppt/presentation.xml"),
      ),
    );
    const issues = await check(engine, [
      { op: "deleteSlide", pageIndex: 0 },
      { op: "deleteSlide", pageIndex: 4 },
    ]);
    assert.deepEqual(
      issues.map((issue) => [issue.operationIndex, issue.path, issue.code]),
      [
        [0, "/pageIndex", "last-slide"],
        [1, "/pageIndex", "range"],
      ],
    );
    await assert.rejects(run(engine, [{ op: "deleteSlide", pageIndex: 0 }]), {
      code: "invalid-operation",
    });
    await engine.dispose();
  });

  it("moves slides by reordering the slide list only", async () => {
    const engine = await open(
      buildDeck({ slides: [{ shapes: [] }, { shapes: [] }, { shapes: [] }] }),
    );
    const original = await partText(
      await saved(engine),
      "/ppt/presentation.xml",
    );
    const change = await run(engine, [{ op: "moveSlide", from: 0, to: 2 }]);
    assert.deepEqual(change.changedPages, [0, 1, 2]);
    assert.deepEqual(await keys(engine), ["sld2", "sld3", "sld1"]);
    assert.deepEqual(
      (await engine.getElements({ pageIndex: 2 }, signal)).length,
      0,
    );
    assert.deepEqual(engine.package.changedParts, ["/ppt/presentation.xml"]);
    const pkg = await saved(engine);
    assert.ok(
      /<p:sldIdLst><p:sldId id="257" r:id="rId4"\/><p:sldId id="258" r:id="rId5"\/><p:sldId id="256" r:id="rId3"\/><\/p:sldIdLst>/.test(
        await partText(pkg, "/ppt/presentation.xml"),
      ),
    );
    await run(engine, [{ op: "moveSlide", from: 2, to: 0 }]);
    assert.deepEqual(await keys(engine), ["sld1", "sld2", "sld3"]);
    // The list is back as it was; the part is rewritten, so only its text is compared.
    assert.equal(
      await partText(await saved(engine), "/ppt/presentation.xml"),
      original,
    );
    const same = await run(engine, [{ op: "moveSlide", from: 1, to: 1 }]);
    assert.deepEqual(same.changedPages, []);
    const issues = await check(engine, [{ op: "moveSlide", from: 3, to: 0 }]);
    assert.deepEqual(
      issues.map((issue) => [issue.path, issue.code]),
      [["/from", "range"]],
    );
    await engine.dispose();
  });

  it("round-trips slide operations through the session with undo, redo and a batch on the new slide", async () => {
    const bytes = buildDeck({
      slides: [{ shapes: [textShape({ id: 2, x: 0, y: 0, cx: 10, cy: 10 })] }],
    });
    const { session, end } = await pptxSession(bytes);
    try {
      const receipt = await session.apply([
        { op: "insertSlide", index: 1, layout: "layout2" },
        {
          op: "insertTextBox",
          pageIndex: 1,
          rect: { x: 10, y: 10, width: 100, height: 20 },
          text: "On the new slide",
        },
        { op: "replaceText", target: "$0", text: "Title text" },
      ]);
      assert.deepEqual(receipt.createdIds, ["sld2:2", "sld2:3", "sld2:4"]);
      assert.equal(receipt.pageCount, 2);
      assert.equal(session.state.pageCount, 2);
      const slides = await session.getSlides();
      assert.deepEqual(
        slides.items.map((slide) => [slide.key, slide.layout]),
        [
          ["sld1", "layout1"],
          ["sld2", "layout2"],
        ],
      );
      assert.equal(
        (await session.getElement("sld2:2")).item?.text,
        "Title text",
      );
      assert.equal(
        (await session.getElement("sld2:4")).item?.text,
        "On the new slide",
      );
      const edited = (await session.save()).bytes;
      await session.duplicateSlide({ pageIndex: 1 });
      assert.equal(session.state.pageCount, 3);
      await session.moveSlide({ from: 2, to: 0 });
      assert.deepEqual(
        (await session.getSlides()).items.map((slide) => slide.key),
        ["sld3", "sld1", "sld2"],
      );
      await session.deleteSlide({ pageIndex: 1 });
      assert.deepEqual(
        (await session.getSlides()).items.map((slide) => slide.key),
        ["sld3", "sld2"],
      );
      await session.undo();
      await session.undo();
      await session.undo();
      assert.deepEqual((await session.save()).bytes, edited);
      await session.undo();
      assert.deepEqual((await session.save()).bytes, bytes);
      assert.equal(session.state.pageCount, 1);
      await session.redo();
      assert.deepEqual((await session.save()).bytes, edited);
    } finally {
      await end();
    }
  });
});
