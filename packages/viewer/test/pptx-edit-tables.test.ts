import assert from "node:assert/strict";
import { describe, it } from "node:test";

import { PptxEditEngine } from "../src/edit/pptx/engine.js";
import { OoxmlPackage } from "../src/edit/ooxml/package.js";
import {
  defaultResourceLimits,
  type PptxElement,
  type PptxOperation,
} from "../src/index.js";
import { buildDeck, table, textShape } from "./fixtures/pptx-builder.js";
import { pptxSession } from "./fixtures/pptx-session.js";
import { tinyJpeg } from "./fixtures/tiny-jpeg.js";

/*
 * Task 47 of the PPTX module: a picture is stored once as a media part and
 * related from the slide; a table is a graphic frame with a grid, rows and
 * cells; a cell's text is replaced like a shape's text.
 */

const signal = new AbortController().signal;
const SLIDE = "/ppt/slides/slide1.xml";
const RELS = "/ppt/slides/_rels/slide1.xml.rels";
const PNG = new Uint8Array([
  0x89, 0x50, 0x4e, 0x47, 0x0d, 0x0a, 0x1a, 0x0a, 0, 0, 0, 13, 0x49, 0x48, 0x44,
  0x52,
]);

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

async function partText(engine: PptxEditEngine, name: string): Promise<string> {
  return new TextDecoder().decode(await (await saved(engine)).part(name));
}

async function element(
  engine: PptxEditEngine,
  id: string,
): Promise<PptxElement> {
  const found = await engine.getElement(id, signal);
  assert.ok(found, `element ${id}`);
  return found;
}

describe("PPTX images and tables (pptx-edit)", () => {
  it("inserts pictures as media parts related from the slide, storing the same bytes once", async () => {
    const engine = await open(
      buildDeck({
        slides: [
          { shapes: [textShape({ id: 2, x: 0, y: 0, cx: 10, cy: 10 })] },
          { shapes: [] },
        ],
      }),
    );
    const change = await run(engine, [
      {
        op: "insertImage",
        pageIndex: 0,
        rect: { x: 10, y: 20, width: 100, height: 50 },
        data: PNG,
        mimeType: "image/png",
      },
      {
        op: "insertImage",
        pageIndex: 1,
        rect: { x: 0, y: 0, width: 10, height: 10 },
        data: PNG,
        mimeType: "image/png",
      },
      {
        op: "insertImage",
        pageIndex: 0,
        rect: { x: 0, y: 0, width: 10, height: 10 },
        data: tinyJpeg(),
        mimeType: "image/jpeg",
      },
    ]);
    assert.deepEqual(change.createdIds, ["sld1:3", "sld2:2", "sld1:4"]);
    const picture = await element(engine, "sld1:3");
    assert.equal(picture.kind, "image");
    assert.equal(picture.name, "Picture 2");
    assert.deepEqual(picture.bounds, { x: 10, y: 20, width: 100, height: 50 });
    assert.deepEqual(picture.operations, [
      "moveElement",
      "resizeElement",
      "deleteElement",
    ]);
    const pkg = await saved(engine);
    const media = pkg.currentPartNames.filter((name) =>
      name.startsWith("/ppt/media/"),
    );
    assert.deepEqual(media, [
      "/ppt/media/image1.png",
      "/ppt/media/image1.jpeg",
    ]);
    assert.deepEqual(await pkg.part("/ppt/media/image1.png"), PNG);
    const rels = await partText(engine, RELS);
    assert.ok(rels.includes('Target="../media/image1.png"'));
    assert.ok(rels.includes('Target="../media/image1.jpeg"'));
    const xml = await partText(engine, SLIDE);
    assert.ok(
      xml.includes(
        '<p:pic><p:nvPicPr><p:cNvPr id="3" name="Picture 2"/><p:cNvPicPr><a:picLocks noChangeAspect="1"/></p:cNvPicPr><p:nvPr/></p:nvPicPr><p:blipFill><a:blip r:embed="rId2"/><a:stretch><a:fillRect/></a:stretch></p:blipFill><p:spPr><a:xfrm><a:off x="95250" y="190500"/><a:ext cx="952500" cy="476250"/></a:xfrm>',
      ),
      xml,
    );
    const types = await partText(engine, "/[Content_Types].xml");
    assert.ok(
      types.includes('Extension="png"') && types.includes('Extension="jpeg"'),
    );
    assert.ok(/<Default Extension="png"[^>]*>/.test(types));

    const issues = await check(engine, [
      {
        op: "insertImage",
        pageIndex: 0,
        rect: { x: 0, y: 0, width: 1, height: 1 },
        data: new Uint8Array([1, 2, 3, 4]),
        mimeType: "image/png",
      },
      {
        op: "insertImage",
        pageIndex: 3,
        rect: { x: 0, y: 0, width: 1, height: 1 },
        data: PNG,
        mimeType: "image/png",
      },
      {
        op: "insertImage",
        pageIndex: 0,
        rect: { x: 0, y: 0, width: 1, height: 1 },
        data: "asset:" + "0".repeat(64),
        mimeType: "image/png",
      },
    ]);
    assert.deepEqual(
      issues.map((issue) => [issue.operationIndex, issue.path, issue.code]),
      [
        [0, "/data", "invalid-value"],
        [1, "/pageIndex", "unknown-target"],
        [2, "/data", "unknown-asset"],
      ],
    );
    await engine.dispose();
  });

  it("inserts tables with a grid, the deck's default style and one paragraph per cell", async () => {
    const engine = await open(
      buildDeck({
        slides: [{ shapes: [] }],
        parts: [
          {
            name: "ppt/tableStyles.xml",
            data: '<?xml version="1.0" encoding="UTF-8" standalone="yes"?>\n<a:tblStyleLst xmlns:a="http://schemas.openxmlformats.org/drawingml/2006/main" def="{5C22544A-7EE6-4342-B048-85BDC9FD1C3A}"/>',
            contentType:
              "application/vnd.openxmlformats-officedocument.presentationml.tableStyles+xml",
          },
        ],
      }),
    );
    const change = await run(engine, [
      {
        op: "insertTable",
        pageIndex: 0,
        rect: { x: 96, y: 96, width: 300, height: 90 },
        rows: [
          ["Name", "Value"],
          ["a", "1"],
          ["b", "2"],
        ],
        columnWidths: [2, 1],
        style: { bandRow: false },
      },
    ]);
    assert.deepEqual(change.createdIds, ["sld1:2"]);
    const grid = await element(engine, "sld1:2");
    assert.equal(grid.kind, "table");
    assert.equal(grid.name, "Table 1");
    assert.deepEqual(grid.table, {
      rows: [
        ["Name", "Value"],
        ["a", "1"],
        ["b", "2"],
      ],
    });
    assert.deepEqual(grid.bounds, { x: 96, y: 96, width: 300, height: 90 });
    const xml = await partText(engine, SLIDE);
    assert.ok(
      xml.includes(
        '<a:tbl><a:tblPr firstRow="1" bandRow="0"><a:tableStyleId>{5C22544A-7EE6-4342-B048-85BDC9FD1C3A}</a:tableStyleId></a:tblPr><a:tblGrid><a:gridCol w="1905000"/><a:gridCol w="952500"/></a:tblGrid><a:tr h="285750"><a:tc><a:txBody><a:bodyPr/><a:lstStyle/><a:p><a:r><a:t>Name</a:t></a:r></a:p></a:txBody><a:tcPr/></a:tc>',
      ),
      xml,
    );
    assert.ok(
      xml.includes(
        '<a:tr h="285750"><a:tc><a:txBody><a:bodyPr/><a:lstStyle/><a:p><a:r><a:t>b</a:t></a:r></a:p></a:txBody><a:tcPr/></a:tc>',
      ),
      "the last row absorbs the rounding",
    );

    const issues = await check(engine, [
      {
        op: "insertTable",
        pageIndex: 0,
        rect: { x: 0, y: 0, width: 1, height: 1 },
        rows: [["a", "b"], ["c"]],
      },
      {
        op: "insertTable",
        pageIndex: 0,
        rect: { x: 0, y: 0, width: 1, height: 1 },
        rows: [["a"]],
        columnWidths: [1, 2],
      },
      {
        op: "insertTable",
        pageIndex: 0,
        rect: { x: 0, y: 0, width: 1, height: 1 },
        rows: [["a\u0001"]],
      },
    ]);
    assert.deepEqual(
      issues.map((issue) => [issue.operationIndex, issue.path, issue.code]),
      [
        [0, "/rows", "invalid-value"],
        [1, "/columnWidths", "invalid-value"],
        [2, "/rows/0/0", "invalid-text"],
      ],
    );
    await engine.dispose();
  });

  it("changes a cell's text keeping its properties and refuses cells outside the table", async () => {
    const engine = await open(
      buildDeck({
        slides: [
          {
            shapes: [
              table({
                id: 2,
                x: 0,
                y: 0,
                cx: 1828800,
                cy: 914400,
                rows: [
                  ["A", "B"],
                  ["C", "D"],
                ],
              }).replace(
                "<a:t>D</a:t></a:r></a:p></a:txBody><a:tcPr/>",
                '<a:t>D</a:t></a:r></a:p></a:txBody><a:tcPr><a:solidFill><a:srgbClr val="FFFF00"/></a:solidFill></a:tcPr>',
              ),
            ],
          },
        ],
      }),
    );
    await run(engine, [
      {
        op: "setTableCell",
        target: "sld1:2",
        row: 1,
        column: 1,
        text: "Changed\nTwice",
      },
      { op: "setTableCell", target: "sld1:2", row: 0, column: 0, text: "" },
    ]);
    const grid = await element(engine, "sld1:2");
    assert.deepEqual(grid.table, {
      rows: [
        ["", "B"],
        ["C", "Changed\nTwice"],
      ],
    });
    assert.equal(grid.text, "\tB\nC\tChanged\nTwice");
    const xml = await partText(engine, SLIDE);
    assert.ok(
      xml.includes(
        '<a:tc><a:txBody><a:bodyPr/><a:lstStyle/><a:p><a:r><a:rPr lang="en-US"/><a:t>Changed</a:t></a:r></a:p><a:p><a:r><a:rPr lang="en-US"/><a:t>Twice</a:t></a:r></a:p></a:txBody><a:tcPr><a:solidFill><a:srgbClr val="FFFF00"/></a:solidFill></a:tcPr></a:tc>',
      ),
      xml,
    );
    assert.ok(
      xml.includes(
        "<a:tc><a:txBody><a:bodyPr/><a:lstStyle/><a:p></a:p></a:txBody><a:tcPr/></a:tc>",
      ),
      "an emptied cell keeps one paragraph",
    );
    const issues = await check(engine, [
      { op: "setTableCell", target: "sld1:2", row: 2, column: 0, text: "x" },
      { op: "setTableCell", target: "sld1:2", row: 0, column: 5, text: "x" },
      { op: "setTableCell", target: "sld1:9", row: 0, column: 0, text: "x" },
    ]);
    assert.deepEqual(
      issues.map((issue) => [issue.operationIndex, issue.path, issue.code]),
      [
        [0, "/row", "range"],
        [1, "/column", "range"],
        [2, "/target", "unknown-target"],
      ],
    );
    await engine.dispose();
  });

  it("round-trips pictures and tables through the session with assets and undo", async () => {
    const bytes = buildDeck({ slides: [{ shapes: [] }] });
    const { session, end } = await pptxSession(bytes);
    try {
      const asset = await session.addAsset(PNG, { mimeType: "image/png" });
      const receipt = await session.apply([
        {
          op: "insertImage",
          pageIndex: 0,
          rect: { x: 0, y: 0, width: 50, height: 50 },
          data: asset,
          mimeType: "image/png",
        },
        {
          op: "insertTable",
          pageIndex: 0,
          rect: { x: 100, y: 100, width: 200, height: 60 },
          rows: [["x", "y"]],
        },
        { op: "setTableCell", target: "$1", row: 0, column: 1, text: "z" },
        {
          op: "resizeElement",
          target: "$0",
          rect: { x: 0, y: 0, width: 80, height: 40 },
        },
      ]);
      assert.deepEqual(receipt.createdIds, ["sld1:2", "sld1:3"]);
      const items = (await session.getElements()).items;
      assert.deepEqual(
        items.map((item) => [item.id, item.kind]),
        [
          ["sld1:2", "image"],
          ["sld1:3", "table"],
        ],
      );
      assert.deepEqual(items[1]!.table, { rows: [["x", "z"]] });
      assert.deepEqual(items[0]!.bounds, { x: 0, y: 0, width: 80, height: 40 });
      const edited = (await session.save()).bytes;
      await session.undo();
      assert.deepEqual((await session.save()).bytes, bytes);
      await session.redo();
      assert.deepEqual((await session.save()).bytes, edited);
    } finally {
      await end();
    }
  });
});
