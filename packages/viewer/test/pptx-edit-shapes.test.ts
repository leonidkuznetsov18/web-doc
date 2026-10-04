import assert from "node:assert/strict";
import { describe, it } from "node:test";

import { PptxEditEngine } from "../src/edit/pptx/engine.js";
import { emuToPx } from "../src/edit/pptx/geometry.js";
import { OoxmlPackage } from "../src/edit/ooxml/package.js";
import {
  defaultResourceLimits,
  type PptxElement,
  type PptxOperation,
} from "../src/index.js";
import {
  buildDeck,
  connector,
  group,
  LAYOUT2_TITLE,
  picture,
  textShape,
} from "./fixtures/pptx-builder.js";
import { pptxSession } from "./fixtures/pptx-session.js";

/*
 * Task 46 of the PPTX module: fills and lines written into p:spPr in
 * schema order, frames written to a:xfrm (created for an inheriting
 * placeholder, mapped through groups, rotation kept), deletion that takes
 * only the element's own relationships along, and text boxes appended to
 * the shape tree with the next free id.
 */

const signal = new AbortController().signal;
const SLIDE = "/ppt/slides/slide1.xml";
const RELS = "/ppt/slides/_rels/slide1.xml.rels";
const IMAGE_TYPE =
  "http://schemas.openxmlformats.org/officeDocument/2006/relationships/image";

const run = (engine: PptxEditEngine, operations: PptxOperation[]) =>
  engine.apply(operations, signal);
const check = (engine: PptxEditEngine, operations: PptxOperation[]) =>
  engine.validate(operations, signal);

async function open(bytes: Uint8Array): Promise<PptxEditEngine> {
  return PptxEditEngine.open(bytes, defaultResourceLimits, signal);
}

async function partText(engine: PptxEditEngine, name: string): Promise<string> {
  const pkg = await OoxmlPackage.open(
    await engine.materialize("save", {}, signal),
    { limits: defaultResourceLimits },
  );
  return new TextDecoder().decode(await pkg.part(name));
}

async function element(
  engine: PptxEditEngine,
  id: string,
): Promise<PptxElement> {
  const found = await engine.getElement(id, signal);
  assert.ok(found, `element ${id}`);
  return found;
}

function px(box: { x: number; y: number; cx: number; cy: number }) {
  return {
    x: emuToPx(box.x),
    y: emuToPx(box.y),
    width: emuToPx(box.cx),
    height: emuToPx(box.cy),
  };
}

function close(actual: number, expected: number, tolerance = 0.01): void {
  assert.ok(
    Math.abs(actual - expected) <= tolerance,
    `expected ${actual} to be within ${tolerance} of ${expected}`,
  );
}

function sameRect(
  actual: { x: number; y: number; width: number; height: number },
  expected: { x: number; y: number; width: number; height: number },
  tolerance = 0.01,
): void {
  close(actual.x, expected.x, tolerance);
  close(actual.y, expected.y, tolerance);
  close(actual.width, expected.width, tolerance);
  close(actual.height, expected.height, tolerance);
}

describe("PPTX shape operations (pptx-edit)", () => {
  it("rejects unrepresentable rotated group resize without changing bytes or history, including dry run", async () => {
    const bytes = buildDeck({
      slides: [
        {
          shapes: [
            group({
              id: 2,
              x: 914400,
              y: 914400,
              cx: 914400,
              cy: 914400,
              rotation: 45,
              child: { x: 0, y: 0, cx: 914400, cy: 914400 },
              children: [
                textShape({ id: 3, x: 0, y: 0, cx: 457200, cy: 457200 }),
              ],
            }),
          ],
        },
      ],
    });
    const { session, end } = await pptxSession(bytes);
    try {
      await session.moveElement({ target: "sld1:2", by: { dx: 10, dy: 0 } });
      const redoBytes = (await session.save()).bytes;
      await session.undo();
      const state = session.state;
      const before = (await session.getElement("sld1:2")).item;
      assert.ok(before);
      for (const dryRun of [false, true]) {
        const operations: PptxOperation[] = [
          {
            op: "resizeElement",
            target: before.id,
            rect: { ...before.bounds, width: before.bounds.width * 2 },
          },
        ];
        await assert.rejects(session.applyJson(operations, { dryRun }), {
          code: "invalid-operation",
        });
        assert.deepEqual(session.state, state);
        assert.deepEqual((await session.save()).bytes, bytes);
        assert.deepEqual(
          (await session.getElement(before.id)).item?.bounds,
          before.bounds,
        );
      }
      await session.redo();
      assert.deepEqual((await session.save()).bytes, redoBytes);
    } finally {
      await end();
    }
  });

  it("rejects tiny anisotropic group shear even when overflowing descendants amplify it", async () => {
    const bytes = buildDeck({
      slides: [
        {
          shapes: [
            group({
              id: 2,
              x: 914400,
              y: 914400,
              cx: 914400,
              cy: 914400,
              rotation: 45,
              child: { x: 0, y: 0, cx: 914400, cy: 914400 },
              children: [
                textShape({
                  id: 3,
                  x: 914400000000,
                  y: 0,
                  cx: 457200,
                  cy: 457200,
                }),
              ],
            }),
          ],
        },
      ],
    });
    const { session, end } = await pptxSession(bytes);
    try {
      const before = (await session.getElement("sld1:2")).item;
      const child = (await session.getElement("sld1:3")).item;
      assert.ok(before && child);
      const state = session.state;
      for (const dryRun of [false, true]) {
        const operations: PptxOperation[] = [
          {
            op: "resizeElement",
            target: before.id,
            rect: { ...before.bounds, width: before.bounds.width + 0.00001 },
          },
        ];
        await assert.rejects(session.applyJson(operations, { dryRun }), {
          code: "invalid-operation",
        });
        assert.deepEqual(session.state, state);
        assert.deepEqual((await session.save()).bytes, bytes);
        assert.deepEqual(
          (await session.getElement(child.id)).item?.bounds,
          child.bounds,
        );
      }
      await session.resizeElement({
        target: before.id,
        rect: {
          ...before.bounds,
          width: before.bounds.width * 2,
          height: before.bounds.height * 2,
        },
      });
      const resizedChild = (await session.getElement(child.id)).item;
      assert.ok(resizedChild);
      sameRect(
        resizedChild.bounds,
        {
          x: before.bounds.x + (child.bounds.x - before.bounds.x) * 2,
          y: before.bounds.y + (child.bounds.y - before.bounds.y) * 2,
          width: child.bounds.width * 2,
          height: child.bounds.height * 2,
        },
        0.001,
      );
      await session.undo();
      assert.deepEqual((await session.save()).bytes, bytes);
    } finally {
      await end();
    }
  });

  it("rejects a resize whose inherited transform yields non-finite corner bounds", async () => {
    const bytes = buildDeck({
      slides: [
        {
          shapes: [
            group({
              id: 2,
              x: 1e308,
              y: 1e308,
              cx: 914400,
              cy: 914400,
              child: { x: 0, y: 0, cx: 914400, cy: 914400 },
              children: [
                textShape({
                  id: 3,
                  x: 0,
                  y: 0,
                  cx: 457200,
                  cy: 228600,
                  rotation: 140,
                }),
              ],
            }),
          ],
        },
      ],
    });
    const { session, end } = await pptxSession(bytes);
    try {
      const state = session.state;
      const operations: PptxOperation[] = [
        {
          op: "resizeElement",
          target: "sld1:3",
          rect: { x: 0, y: 0, width: 100, height: 100 },
        },
      ];
      for (const dryRun of [false, true]) {
        await assert.rejects(session.applyJson(operations, { dryRun }), {
          code: "invalid-operation",
        });
        assert.deepEqual(session.state, state);
        assert.deepEqual((await session.save()).bytes, bytes);
      }
    } finally {
      await end();
    }
  });

  for (const childSpace of [
    "explicit",
    "missing",
    "missing-offset",
    "missing-extent",
    "zero-x",
    "zero-y",
    "zero-both",
  ] as const) {
    it(`preserves rotated group child placement with ${childSpace} child space during uniform resize, save/reopen and one undo`, async () => {
      const groupXml = group({
        id: 2,
        x: 914400,
        y: 914400,
        cx: 914400,
        cy: 914400,
        rotation: 45,
        flipH: true,
        child: {
          x: 114300,
          y: 228600,
          cx:
            childSpace === "zero-x" || childSpace === "zero-both" ? 0 : 914400,
          cy:
            childSpace === "zero-y" || childSpace === "zero-both" ? 0 : 914400,
        },
        children: [
          textShape({ id: 3, x: 0, y: 114300, cx: 457200, cy: 228600 }),
        ],
      });
      const offXml = '<a:chOff x="114300" y="228600"/>';
      const extXml = '<a:chExt cx="914400" cy="914400"/>';
      const shape =
        childSpace === "missing"
          ? groupXml.replace(offXml + extXml, "")
          : childSpace === "missing-offset"
            ? groupXml.replace(offXml, "")
            : childSpace === "missing-extent"
              ? groupXml.replace(extXml, "")
              : groupXml;
      const taggedShape =
        childSpace === "zero-x" || childSpace === "missing-extent"
          ? shape
              .replace("<a:chOff ", '<a:chOff xmlns:qa="urn:qa" qa:keep="off" ')
              .replace("<a:chExt ", '<a:chExt xmlns:qa="urn:qa" qa:keep="ext" ')
          : shape;
      const bytes = buildDeck({ slides: [{ shapes: [taggedShape] }] });
      const { session, end } = await pptxSession(bytes);
      try {
        const before = (await session.getElement("sld1:2")).item;
        const child = (await session.getElement("sld1:3")).item;
        assert.ok(before && child);
        const rect = {
          x: 30,
          y: 40,
          width: before.bounds.width * 2,
          height: before.bounds.height * 2,
        };
        const state = session.state;
        const invalidOperations: PptxOperation[] = [
          {
            op: "resizeElement",
            target: before.id,
            rect: { ...rect, height: rect.height * 1.5 },
          },
        ];
        for (const dryRun of [false, true]) {
          await assert.rejects(
            session.applyJson(invalidOperations, { dryRun }),
            { code: "invalid-operation" },
          );
          assert.deepEqual(session.state, state);
          assert.deepEqual((await session.save()).bytes, bytes);
        }
        await session.resizeElement({ target: before.id, rect });
        const expectedChild = {
          x: rect.x + (child.bounds.x - before.bounds.x) * 2,
          y: rect.y + (child.bounds.y - before.bounds.y) * 2,
          width: child.bounds.width * 2,
          height: child.bounds.height * 2,
        };
        const resizedGroup = (await session.getElement(before.id)).item;
        const resizedChild = (await session.getElement(child.id)).item;
        assert.ok(resizedGroup && resizedChild);
        sameRect(resizedGroup.bounds, rect, 0.001);
        sameRect(resizedChild.bounds, expectedChild, 0.001);
        const reopened = await open((await session.save()).bytes);
        try {
          const savedXml = await partText(reopened, SLIDE);
          const transform = savedXml.match(
            /<p:grpSp>[\s\S]*?<p:grpSpPr>([\s\S]*?)<\/p:grpSpPr>/u,
          )?.[1];
          assert.ok(transform, "saved native group transform");
          assert.ok(transform.indexOf("<a:chOff ") >= 0);
          assert.ok(
            transform.indexOf("<a:chExt ") > transform.indexOf("<a:chOff "),
          );
          if (childSpace === "zero-x" || childSpace === "missing-extent")
            assert.match(transform, /<a:chOff[^>]*qa:keep="off"/u);
          if (childSpace === "zero-x")
            assert.match(transform, /<a:chExt[^>]*qa:keep="ext"/u);
          sameRect((await element(reopened, before.id)).bounds, rect, 0.001);
          sameRect(
            (await element(reopened, child.id)).bounds,
            expectedChild,
            0.001,
          );
        } finally {
          await reopened.dispose();
        }
        await session.undo();
        assert.deepEqual((await session.save()).bytes, bytes);
        assert.equal(session.state.canUndo, false);
      } finally {
        await end();
      }
    });
  }

  it("solves single shape bounds through a nonuniform rotated parent and rejects group shear", async () => {
    const bytes = buildDeck({
      slides: [
        {
          shapes: [
            group({
              id: 2,
              x: 914400,
              y: 914400,
              cx: 1828800,
              cy: 914400,
              rotation: 20,
              flipH: true,
              child: { x: 0, y: 0, cx: 914400, cy: 914400 },
              children: [
                textShape({
                  id: 3,
                  x: 114300,
                  y: 114300,
                  cx: 457200,
                  cy: 228600,
                  rotation: 30,
                  flipV: true,
                }),
              ],
            }),
          ],
        },
      ],
    });
    const engine = await open(bytes);
    const resizedFixture = await open(
      buildDeck({
        slides: [
          {
            shapes: [
              group({
                id: 2,
                x: 914400,
                y: 914400,
                cx: 1828800,
                cy: 914400,
                rotation: 20,
                flipH: true,
                child: { x: 0, y: 0, cx: 914400, cy: 914400 },
                children: [
                  textShape({
                    id: 3,
                    x: 114300,
                    y: 114300,
                    cx: 685800,
                    cy: 285750,
                    rotation: 30,
                    flipV: true,
                  }),
                ],
              }),
            ],
          },
        ],
      }),
    );
    try {
      const expected = (await element(resizedFixture, "sld1:3")).bounds;
      await run(engine, [
        { op: "resizeElement", target: "sld1:3", rect: expected },
      ]);
      sameRect((await element(engine, "sld1:3")).bounds, expected, 0.001);
      const groupBefore = (await element(engine, "sld1:2")).bounds;
      const snapshot = await engine.materialize("save", {}, signal);
      await assert.rejects(
        run(engine, [
          {
            op: "resizeElement",
            target: "sld1:2",
            rect: { ...groupBefore, width: groupBefore.width * 1.2 },
          },
        ]),
        { code: "invalid-operation" },
      );
      assert.deepEqual(await engine.materialize("save", {}, signal), snapshot);
    } finally {
      await engine.dispose();
      await resizedFixture.dispose();
    }
  });

  it("writes fills and lines into p:spPr in schema order and removes them on null", async () => {
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
                fill: '<a:solidFill><a:srgbClr val="FF0000"/></a:solidFill>',
                line: '<a:ln w="12700"><a:solidFill><a:srgbClr val="00FF00"/></a:solidFill><a:prstDash val="dash"/></a:ln>',
              }),
              textShape({ id: 3, x: 0, y: 0, cx: 914400, cy: 914400 }),
              connector({ id: 4, x: 0, y: 0, cx: 914400, cy: 0 }),
            ],
          },
        ],
      }),
    );
    await run(engine, [
      {
        op: "setShapeStyle",
        target: "sld1:2",
        fill: { theme: "accent2", mods: { lumMod: 60000, lumOff: 40000 } },
        line: { color: "#0000FF", width: 3 },
      },
    ]);
    let shape = await element(engine, "sld1:2");
    assert.deepEqual(shape.shapeStyle, {
      fill: { theme: "accent2", mods: { lumMod: 60000, lumOff: 40000 } },
      line: { color: "#0000FF", width: 3 },
    });
    let xml = await partText(engine, SLIDE);
    assert.ok(
      xml.includes(
        '<a:prstGeom prst="rect"><a:avLst/></a:prstGeom><a:solidFill><a:schemeClr val="accent2"><a:lumMod val="60000"/><a:lumOff val="40000"/></a:schemeClr></a:solidFill><a:ln w="38100"><a:solidFill><a:srgbClr val="0000FF"/></a:solidFill><a:prstDash val="dash"/></a:ln></p:spPr>',
      ),
      "the dash of the line and the geometry are kept",
    );

    // A shape without explicit style: the fill goes after the geometry, the line after the fill.
    await run(engine, [
      { op: "setShapeStyle", target: "sld1:3", fill: "none", line: "none" },
    ]);
    shape = await element(engine, "sld1:3");
    assert.deepEqual(shape.shapeStyle, {
      fill: "none",
      line: { color: "none", width: 0.75 },
    });
    xml = await partText(engine, SLIDE);
    assert.ok(
      xml.includes("</a:prstGeom><a:noFill/><a:ln><a:noFill/></a:ln></p:spPr>"),
      xml,
    );

    await run(engine, [
      { op: "setShapeStyle", target: "sld1:2", fill: null, line: null },
    ]);
    assert.deepEqual((await element(engine, "sld1:2")).shapeStyle, {});
    xml = await partText(engine, SLIDE);
    assert.ok(xml.includes("</a:prstGeom></p:spPr>"), "fill and line removed");

    await run(engine, [
      { op: "setShapeStyle", target: "sld1:4", line: { color: "#123456" } },
    ]);
    assert.deepEqual((await element(engine, "sld1:4")).shapeStyle, {
      line: { color: "#123456", width: 1.5 },
    });
    const issues = await check(engine, [
      { op: "setShapeStyle", target: "sld1:4", fill: "#000000" },
      { op: "setShapeStyle", target: "sld1:2", fill: "blue" },
      {
        op: "setShapeStyle",
        target: "sld1:2",
        line: { color: { theme: "nope" } },
      },
      { op: "setShapeStyle", target: "sld1:9" },
    ]);
    assert.deepEqual(
      issues.map((issue) => [issue.operationIndex, issue.path, issue.code]),
      [
        [0, "/fill", "invalid-value"],
        [1, "/fill", "invalid-value"],
        [2, "/line/color", "invalid-value"],
        [3, "/target", "unknown-target"],
      ],
    );
    await engine.dispose();
  });

  it("moves and resizes frames, creating a:xfrm for an inheriting placeholder and keeping rotation", async () => {
    const engine = await open(
      buildDeck({
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
              }),
              textShape({
                id: 3,
                x: 914400,
                y: 914400,
                cx: 1828800,
                cy: 914400,
                rotation: 90,
              }),
              group({
                id: 4,
                x: 4572000,
                y: 2743200,
                cx: 1828800,
                cy: 1828800,
                child: { x: 0, y: 0, cx: 914400, cy: 914400 },
                children: [
                  textShape({ id: 5, x: 0, y: 0, cx: 457200, cy: 457200 }),
                ],
              }),
            ],
          },
        ],
      }),
    );
    // The inherited frame becomes explicit on the slide.
    const before = await partText(engine, SLIDE);
    assert.ok(
      !/<p:sp><p:nvSpPr><p:cNvPr id="2"[^]*?<a:xfrm>/.test(
        before.split('<p:cNvPr id="3"')[0]!,
      ),
    );
    await run(engine, [
      { op: "moveElement", target: "sld1:2", by: { dx: 10, dy: -5 } },
    ]);
    let title = await element(engine, "sld1:2");
    const expected = px(LAYOUT2_TITLE);
    sameRect(title.bounds, {
      ...expected,
      x: expected.x + 10,
      y: expected.y - 5,
    });
    const after = await partText(engine, SLIDE);
    assert.ok(
      after.includes(
        `<p:spPr><a:xfrm><a:off x="${LAYOUT2_TITLE.x + 95250}" y="${LAYOUT2_TITLE.y - 47625}"/><a:ext cx="${LAYOUT2_TITLE.cx}" cy="${LAYOUT2_TITLE.cy}"/></a:xfrm><a:prstGeom`,
      ),
      after,
    );
    await run(engine, [
      { op: "moveElement", target: "sld1:2", to: { x: 0, y: 0 } },
    ]);
    title = await element(engine, "sld1:2");
    sameRect(title.bounds, { ...expected, x: 0, y: 0 });

    // A rotated shape: the new bounds are the box of the rotated frame.
    await run(engine, [
      {
        op: "resizeElement",
        target: "sld1:3",
        rect: { x: 100, y: 100, width: 50, height: 300 },
      },
    ]);
    const rotated = await element(engine, "sld1:3");
    assert.equal(rotated.rotation, 90);
    sameRect(rotated.bounds, { x: 100, y: 100, width: 50, height: 300 });
    sameRect(rotated.frame!, { x: -25, y: 225, width: 300, height: 50 });

    // A group child: slide-space moves are mapped into the child space (half scale).
    await run(engine, [
      { op: "moveElement", target: "sld1:5", by: { dx: 96, dy: 0 } },
    ]);
    const child = await element(engine, "sld1:5");
    sameRect(
      child.bounds,
      px({ x: 4572000 + 914400, y: 2743200, cx: 914400, cy: 914400 }),
    );
    assert.ok(
      (await partText(engine, SLIDE)).includes(
        '<a:off x="457200" y="0"/><a:ext cx="457200" cy="457200"/>',
      ),
    );
    // Resizing the group keeps its child space, so the child scales with it.
    await run(engine, [
      {
        op: "resizeElement",
        target: "sld1:4",
        rect: { x: 480, y: 288, width: 384, height: 192 },
      },
    ]);
    sameRect((await element(engine, "sld1:4")).bounds, {
      x: 480,
      y: 288,
      width: 384,
      height: 192,
    });
    sameRect((await element(engine, "sld1:5")).bounds, {
      x: 480 + 192,
      y: 288,
      width: 192,
      height: 96,
    });

    const issues = await check(engine, [
      { op: "moveElement", target: "sld1:3" },
      {
        op: "moveElement",
        target: "sld1:3",
        to: { x: 0, y: 0 },
        by: { dx: 1, dy: 1 },
      },
      {
        op: "resizeElement",
        target: "sld1:7",
        rect: { x: 0, y: 0, width: 1, height: 1 },
      },
    ]);
    assert.deepEqual(
      issues.map((issue) => [issue.operationIndex, issue.path, issue.code]),
      [
        [0, "/to", "required"],
        [1, "/by", "conflict"],
        [2, "/target", "unknown-target"],
      ],
    );
    await engine.dispose();
  });

  it("deletes elements with the relationships only they used and reports every removed id", async () => {
    const png = new Uint8Array([0x89, 0x50, 0x4e, 0x47, 1, 2, 3, 4]);
    const engine = await open(
      buildDeck({
        parts: [{ name: "ppt/media/image1.png", data: png }],
        slides: [
          {
            relationships: [
              { id: "rId2", type: IMAGE_TYPE, target: "../media/image1.png" },
              { id: "rId3", type: IMAGE_TYPE, target: "../media/image1.png" },
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
              picture({
                id: 3,
                rId: "rId3",
                x: 0,
                y: 0,
                cx: 914400,
                cy: 914400,
              }),
              picture({
                id: 4,
                rId: "rId3",
                x: 0,
                y: 0,
                cx: 914400,
                cy: 914400,
              }),
              group({
                id: 5,
                x: 0,
                y: 0,
                cx: 914400,
                cy: 914400,
                child: { x: 0, y: 0, cx: 914400, cy: 914400 },
                children: [
                  textShape({ id: 6, x: 0, y: 0, cx: 10, cy: 10 }),
                  group({
                    id: 7,
                    x: 0,
                    y: 0,
                    cx: 10,
                    cy: 10,
                    child: { x: 0, y: 0, cx: 10, cy: 10 },
                    children: [textShape({ id: 8, x: 0, y: 0, cx: 5, cy: 5 })],
                  }),
                ],
              }),
            ],
          },
        ],
      }),
    );
    const first = await run(engine, [
      { op: "deleteElement", target: "sld1:2" },
    ]);
    assert.deepEqual(first.removedIds, ["sld1:2"]);
    assert.deepEqual(first.changedPages, [0]);
    let rels = await partText(engine, RELS);
    assert.ok(
      !rels.includes('Id="rId2"'),
      "the picture's own relationship is gone",
    );
    assert.ok(rels.includes('Id="rId3"'));
    // A relationship two pictures share stays until the last user goes.
    await run(engine, [{ op: "deleteElement", target: "sld1:3" }]);
    rels = await partText(engine, RELS);
    assert.ok(rels.includes('Id="rId3"'));
    await run(engine, [{ op: "deleteElement", target: "sld1:4" }]);
    rels = await partText(engine, RELS);
    assert.ok(!rels.includes('Id="rId3"'));
    assert.ok(rels.includes("slideLayout"), "the layout relationship stays");
    const pkg = await OoxmlPackage.open(
      await engine.materialize("save", {}, signal),
      {
        limits: defaultResourceLimits,
      },
    );
    assert.ok(
      pkg.has("/ppt/media/image1.png"),
      "the media part stays in the package",
    );

    const nested = await run(engine, [
      { op: "deleteElement", target: "sld1:5" },
    ]);
    assert.deepEqual(nested.removedIds, [
      "sld1:5",
      "sld1:6",
      "sld1:7",
      "sld1:8",
    ]);
    assert.deepEqual(
      (await engine.getElements({ pageIndex: 0 }, signal)).map(
        (item) => item.id,
      ),
      [],
    );
    assert.ok(
      (await partText(engine, SLIDE)).includes("</p:grpSpPr></p:spTree>"),
      "the tree keeps its header",
    );
    await engine.dispose();
  });

  it("inserts text boxes with the next free id, the given style and paragraphs", async () => {
    const engine = await open(
      buildDeck({
        slides: [
          { shapes: [textShape({ id: 7, x: 0, y: 0, cx: 10, cy: 10 })] },
          { shapes: [] },
        ],
      }),
    );
    const change = await run(engine, [
      {
        op: "insertTextBox",
        pageIndex: 0,
        rect: { x: 96, y: 48, width: 192, height: 96 },
        text: "Hello\u000bthere\nWorld",
        style: {
          bold: true,
          fontSize: 14,
          color: "#336699",
          align: "center",
          fontFamily: "Arial",
        },
      },
      {
        op: "insertTextBox",
        pageIndex: 1,
        rect: { x: 0, y: 0, width: 10, height: 10 },
        text: "Plain",
      },
      {
        op: "replaceText",
        target: "$0",
        text: "Hello again",
      },
    ]);
    assert.deepEqual(change.createdIds, ["sld1:8", "sld2:2"]);
    assert.deepEqual(change.changedPages, [0, 1]);
    const box = await element(engine, "sld1:8");
    assert.equal(box.kind, "shape");
    assert.equal(box.name, "TextBox 7");
    assert.equal(box.text, "Hello again");
    sameRect(box.bounds, { x: 96, y: 48, width: 192, height: 96 });
    assert.deepEqual(box.textStyle, {
      fontFamily: "Arial",
      fontSize: 14,
      bold: true,
      italic: false,
      underline: false,
      color: "#336699",
      align: "center",
    });
    assert.deepEqual(box.shapeStyle, { fill: "none" });
    assert.ok(box.operations.includes("replaceText"));
    const xml = await partText(engine, SLIDE);
    assert.ok(
      xml.includes(
        '<p:sp><p:nvSpPr><p:cNvPr id="8" name="TextBox 7"/><p:cNvSpPr txBox="1"/><p:nvPr/></p:nvSpPr><p:spPr><a:xfrm><a:off x="914400" y="457200"/><a:ext cx="1828800" cy="914400"/></a:xfrm><a:prstGeom prst="rect"><a:avLst/></a:prstGeom><a:noFill/></p:spPr><p:txBody><a:bodyPr wrap="square" rtlCol="0"><a:spAutoFit/></a:bodyPr><a:lstStyle/><a:p><a:pPr algn="ctr"/><a:r><a:rPr b="1" sz="1400"><a:solidFill><a:srgbClr val="336699"/></a:solidFill><a:latin typeface="Arial"/></a:rPr><a:t>Hello again</a:t></a:r><a:endParaRPr b="1" sz="1400"><a:solidFill><a:srgbClr val="336699"/></a:solidFill><a:latin typeface="Arial"/></a:endParaRPr></a:p></p:txBody></p:sp></p:spTree>',
      ),
      xml,
    );
    assert.equal((await element(engine, "sld2:2")).text, "Plain");

    // Ids are a function of the state: the same batch on the original gives the same ids.
    const again = await open(
      buildDeck({
        slides: [
          { shapes: [textShape({ id: 7, x: 0, y: 0, cx: 10, cy: 10 })] },
          { shapes: [] },
        ],
      }),
    );
    const repeat = await run(again, [
      {
        op: "insertTextBox",
        pageIndex: 0,
        rect: { x: 0, y: 0, width: 10, height: 10 },
        text: "x",
      },
    ]);
    assert.deepEqual(repeat.createdIds, ["sld1:8"]);
    const issues = await check(again, [
      {
        op: "insertTextBox",
        pageIndex: 5,
        rect: { x: 0, y: 0, width: 1, height: 1 },
        text: "x",
      },
      {
        op: "insertTextBox",
        pageIndex: 0,
        rect: { x: 0, y: 0, width: 1, height: 1 },
        text: "\u0000",
      },
    ]);
    assert.deepEqual(
      issues.map((issue) => [issue.operationIndex, issue.path, issue.code]),
      [
        [0, "/pageIndex", "unknown-target"],
        [1, "/text", "invalid-text"],
      ],
    );
    await again.dispose();
    await engine.dispose();
  });

  it("rolls a failing batch back and replays shape edits through the session", async () => {
    const bytes = buildDeck({
      slides: [
        { shapes: [textShape({ id: 2, x: 0, y: 0, cx: 914400, cy: 914400 })] },
      ],
    });
    const { session, end } = await pptxSession(bytes);
    try {
      await assert.rejects(
        session.apply([
          { op: "moveElement", target: "sld1:2", by: { dx: 10, dy: 10 } },
          { op: "deleteElement", target: "sld1:99" },
        ]),
        { code: "invalid-operation" },
      );
      assert.deepEqual((await session.save()).bytes, bytes);
      assert.equal(session.state.revision, 0);
      const receipt = await session.apply([
        {
          op: "insertTextBox",
          pageIndex: 0,
          rect: { x: 10, y: 10, width: 100, height: 20 },
          text: "New",
        },
        { op: "setShapeStyle", target: "$0", fill: "#ABCDEF" },
        { op: "moveElement", target: "sld1:2", to: { x: 50, y: 60 } },
      ]);
      assert.deepEqual(receipt.createdIds, ["sld1:3"]);
      const edited = (await session.save()).bytes;
      await session.undo();
      assert.deepEqual((await session.save()).bytes, bytes);
      await session.redo();
      assert.deepEqual((await session.save()).bytes, edited);
      const moved = (await session.getElement("sld1:2")).item!;
      sameRect(moved.bounds, { x: 50, y: 60, width: 96, height: 96 });
      const removed = await session.deleteElement({ target: "sld1:3" });
      assert.deepEqual(removed.removedIds, ["sld1:3"]);
    } finally {
      await end();
    }
  });
});
