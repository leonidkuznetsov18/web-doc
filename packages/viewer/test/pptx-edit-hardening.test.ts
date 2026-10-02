import assert from "node:assert/strict";
import { describe, it } from "node:test";

import { checkOperations } from "../src/edit/operations.js";
import { PptxEditEngine } from "../src/edit/pptx/engine.js";
import { pptxOperationSchemas } from "../src/edit/pptx/schemas.js";
import { OoxmlPackage } from "../src/edit/ooxml/package.js";
import {
  defaultResourceLimits,
  type PptxElement,
  type PptxOperation,
} from "../src/index.js";
import {
  buildDeck,
  group,
  syntheticDeck,
  textShape,
} from "./fixtures/pptx-builder.js";

/*
 * Regressions from the review of the PPTX module: a rolled-back batch
 * leaves no trace in the engine's view, an aborted read does not poison a
 * slide, ranges cannot split surrogate pairs, coordinates are bounded,
 * line ends are normalized, a rotated frame cannot be resized to nothing,
 * animations go with the shape they target, shape ids are never reused
 * after a deletion, a duplicate keeps a dangling relationship with a
 * warning, sections and custom shows follow slide operations, and the deck
 * index survives slide-only commits.
 */

const signal = new AbortController().signal;
const SLIDE = "/ppt/slides/slide1.xml";
const PRESENTATION = "/ppt/presentation.xml";

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

function deck(...shapes: string[]): Uint8Array {
  return buildDeck({ slides: [{ shapes }] });
}

describe("PPTX hardening (review regressions)", () => {
  it("leaves no trace of a rolled-back batch in the engine's view or its output", async () => {
    const original = deck(
      textShape({
        id: 2,
        x: 0,
        y: 0,
        cx: 914400,
        cy: 914400,
        paragraphs: [["Keep"]],
      }),
    );
    const engine = await open(original);
    await assert.rejects(
      run(engine, [
        {
          op: "insertTextBox",
          pageIndex: 0,
          rect: { x: 0, y: 0, width: 10, height: 10 },
          text: "PHANTOM",
        },
        { op: "replaceText", target: "sld1:999", text: "x" },
      ]),
      { code: "invalid-operation" },
    );
    assert.deepEqual(await engine.materialize("save", {}, signal), original);
    assert.deepEqual(
      (await engine.getElements({ pageIndex: 0 }, signal)).map(
        (item) => item.id,
      ),
      ["sld1:2"],
    );
    assert.equal(engine.pageCount, 1);
    // A later commit at the same revision the failed batch reached must not
    // revive the phantom through a stale scan.
    await run(engine, [
      { op: "replaceText", target: "sld1:2", text: "Changed" },
    ]);
    const xml = await partText(engine, SLIDE);
    assert.ok(!xml.includes("PHANTOM"), xml);
    assert.ok(xml.includes("Changed"));
    assert.deepEqual(
      (await engine.getElements({ pageIndex: 0 }, signal)).map(
        (item) => item.id,
      ),
      ["sld1:2"],
    );
    // An id the failed batch issued is issued again, as a replay would.
    const next = await run(engine, [
      {
        op: "insertTextBox",
        pageIndex: 0,
        rect: { x: 0, y: 0, width: 10, height: 10 },
        text: "Real",
      },
    ]);
    assert.deepEqual(next.createdIds, ["sld1:3"]);
    await engine.dispose();
  });

  it("does not let an aborted read stand in for a slide", async () => {
    const engine = await open(syntheticDeck(3));
    const aborted = AbortSignal.abort();
    await assert.rejects(engine.getElements({ pageIndex: 1 }, aborted), {
      code: "aborted",
    });
    const elements = await engine.getElements({ pageIndex: 1 }, signal);
    assert.deepEqual(
      elements.map((item) => item.id),
      ["sld2:2", "sld2:3"],
    );
    await run(engine, [
      { op: "replaceText", target: "sld2:3", text: "After the abort" },
    ]);
    assert.equal((await element(engine, "sld2:3")).text, "After the abort");
    await engine.dispose();
  });

  it("refuses a range that splits a surrogate pair and normalizes line ends", async () => {
    const engine = await open(
      deck(
        textShape({
          id: 2,
          x: 0,
          y: 0,
          cx: 914400,
          cy: 914400,
          paragraphs: [["a😀b"]],
        }),
      ),
    );
    const issues = await check(engine, [
      {
        op: "replaceText",
        target: "sld1:2",
        text: "X",
        range: {
          start: { elementId: "sld1:2", offset: 2 },
          end: { elementId: "sld1:2", offset: 2 },
        },
      },
      {
        op: "replaceText",
        target: "sld1:2",
        text: "X",
        range: {
          start: { elementId: "sld1:2", offset: 1 },
          end: { elementId: "sld1:2", offset: 2 },
        },
      },
      {
        op: "replaceText",
        target: "sld1:2",
        text: "X",
        range: {
          start: { elementId: "sld1:2", offset: 1 },
          end: { elementId: "sld1:2", offset: 3 },
        },
      },
      { op: "replaceText", target: "sld1:2", text: "bad\f" },
    ]);
    assert.deepEqual(
      issues.map((issue) => [issue.operationIndex, issue.code]),
      [
        [0, "invalid-range"],
        [1, "invalid-range"],
        [3, "invalid-text"],
      ],
    );
    await run(engine, [
      { op: "replaceText", target: "sld1:2", text: "one\r\ntwo\rthree" },
    ]);
    const shape = await element(engine, "sld1:2");
    assert.equal(shape.text, "one\ntwo\nthree");
    assert.ok(!(await partText(engine, SLIDE)).includes("\r"));
    await engine.dispose();
  });

  it("bounds coordinates and refuses a box that collapses a rotated frame", async () => {
    const engine = await open(
      deck(
        textShape({
          id: 2,
          x: 914400,
          y: 914400,
          cx: 1828800,
          cy: 914400,
          rotation: 30,
        }),
      ),
    );
    // Coordinates are bounded by the schemas the core checks before the engine.
    const shape = checkOperations(
      [
        {
          op: "insertTextBox",
          pageIndex: 0,
          rect: { x: 1e300, y: 0, width: 10, height: 10 },
          text: "x",
        },
        {
          op: "moveElement",
          target: "sld1:2",
          by: { dx: Number.MAX_SAFE_INTEGER, dy: 0 },
        },
        {
          op: "resizeElement",
          target: "sld1:2",
          rect: { x: 0, y: 0, width: Infinity, height: 1 },
        },
      ],
      pptxOperationSchemas,
    );
    assert.deepEqual(
      shape.map((issue) => [issue.operationIndex, issue.path, issue.code]),
      [
        [0, "/rect/x", "maximum"],
        [1, "/by/dx", "maximum"],
        [2, "/rect/width", "not-json"],
      ],
    );
    const issues = await check(engine, [
      {
        op: "resizeElement",
        target: "sld1:2",
        rect: { x: 100, y: 100, width: 10, height: 100 },
      },
      {
        op: "resizeElement",
        target: "sld1:2",
        rect: { x: 100, y: 100, width: 200, height: 150 },
      },
    ]);
    assert.deepEqual(
      issues.map((issue) => [issue.operationIndex, issue.path, issue.code]),
      [[0, "/rect", "invalid-value"]],
    );
    await engine.dispose();
  });

  it("removes the animations that target a deleted shape and says so", async () => {
    const timing =
      '<p:timing><p:tnLst><p:par><p:cTn id="1" dur="indefinite" restart="never" nodeType="tmRoot"><p:childTnLst><p:seq concurrent="1" nextAc="seek"><p:cTn id="2" dur="indefinite" nodeType="mainSeq"><p:childTnLst><p:par><p:cTn id="3" fill="hold"><p:childTnLst><p:par><p:cTn id="4" fill="hold"><p:childTnLst><p:set><p:cBhvr><p:cTn id="5" dur="1" fill="hold"/><p:tgtEl><p:spTgt spid="3"/></p:tgtEl><p:attrNameLst><p:attrName>style.visibility</p:attrName></p:attrNameLst></p:cBhvr><p:to><p:strVal val="visible"/></p:to></p:set></p:childTnLst></p:cTn></p:par></p:childTnLst></p:cTn></p:par></p:childTnLst></p:cTn></p:seq></p:childTnLst></p:cTn></p:par></p:tnLst></p:timing>';
    const bytes = deck(
      textShape({ id: 2, x: 0, y: 0, cx: 10, cy: 10 }),
      textShape({ id: 3, x: 0, y: 0, cx: 10, cy: 10 }),
    ).slice();
    // The builder has no slot for timing; splice it before the closing tag.
    const pkg = await OoxmlPackage.open(bytes, {
      limits: defaultResourceLimits,
    });
    const transaction = pkg.transaction();
    const slide = await pkg.xml(SLIDE);
    const { patches } = await import("../src/edit/ooxml/patch.js");
    transaction.patch(slide, [patches.appendChild(slide, slide.root, timing)]);
    await transaction.commit();
    const engine = await open(await pkg.save());
    const untouched = await run(engine, [
      { op: "deleteElement", target: "sld1:2" },
    ]);
    assert.deepEqual(untouched.warnings, []);
    assert.ok((await partText(engine, SLIDE)).includes("<p:timing>"));
    const removed = await run(engine, [
      { op: "deleteElement", target: "sld1:3" },
    ]);
    assert.equal(removed.warnings.length, 1);
    assert.equal(removed.warnings[0]!.details?.reason, "animations-removed");
    assert.ok(!(await partText(engine, SLIDE)).includes("<p:timing>"));
    await engine.dispose();
  });

  it("never reuses a shape id after a deletion, and reproduces ids on replay", async () => {
    const bytes = deck(
      textShape({ id: 2, x: 0, y: 0, cx: 10, cy: 10 }),
      textShape({ id: 3, x: 0, y: 0, cx: 10, cy: 10 }),
    );
    const engine = await open(bytes);
    await run(engine, [{ op: "deleteElement", target: "sld1:3" }]);
    const created = await run(engine, [
      {
        op: "insertTextBox",
        pageIndex: 0,
        rect: { x: 0, y: 0, width: 10, height: 10 },
        text: "a",
      },
    ]);
    assert.deepEqual(created.createdIds, ["sld1:4"]);
    await run(engine, [{ op: "deleteElement", target: "sld1:4" }]);
    const again = await run(engine, [
      {
        op: "insertTextBox",
        pageIndex: 0,
        rect: { x: 0, y: 0, width: 10, height: 10 },
        text: "b",
      },
    ]);
    assert.deepEqual(again.createdIds, ["sld1:5"]);
    // A replay from the original reproduces every id.
    const replay = await open(bytes);
    const results: string[][] = [];
    for (const operations of [
      [{ op: "deleteElement", target: "sld1:3" }],
      [
        {
          op: "insertTextBox",
          pageIndex: 0,
          rect: { x: 0, y: 0, width: 10, height: 10 },
          text: "a",
        },
      ],
      [{ op: "deleteElement", target: "sld1:4" }],
      [
        {
          op: "insertTextBox",
          pageIndex: 0,
          rect: { x: 0, y: 0, width: 10, height: 10 },
          text: "b",
        },
      ],
    ] as PptxOperation[][])
      results.push([...(await run(replay, operations)).createdIds]);
    assert.deepEqual(results, [[], ["sld1:4"], [], ["sld1:5"]]);
    assert.deepEqual(
      await replay.materialize("save", {}, signal),
      await engine.materialize("save", {}, signal),
    );
    await replay.dispose();
    await engine.dispose();
  });

  it("duplicates a slide whose relationship points at a missing part, with a warning", async () => {
    const engine = await open(
      buildDeck({
        slides: [
          {
            relationships: [
              {
                id: "rId2",
                type: "http://schemas.openxmlformats.org/officeDocument/2006/relationships/chart",
                target: "../charts/missing.xml",
              },
            ],
            shapes: [textShape({ id: 2, x: 0, y: 0, cx: 10, cy: 10 })],
          },
        ],
      }),
    );
    const change = await run(engine, [{ op: "duplicateSlide", pageIndex: 0 }]);
    assert.equal(change.pageCount, 2);
    assert.equal(change.warnings.length, 1);
    assert.equal(change.warnings[0]!.details?.reason, "dangling-relationship");
    assert.ok(
      (await partText(engine, "/ppt/slides/_rels/slide2.xml.rels")).includes(
        'Target="../charts/missing.xml"',
      ),
    );
    await engine.dispose();
  });

  it("keeps sections and custom shows in step with slide operations", async () => {
    const bytes = buildDeck({
      slides: [{ shapes: [] }, { shapes: [] }, { shapes: [] }],
    });
    const pkg = await OoxmlPackage.open(bytes, {
      limits: defaultResourceLimits,
    });
    const transaction = pkg.transaction();
    const presentation = await pkg.xml(PRESENTATION);
    const { patches } = await import("../src/edit/ooxml/patch.js");
    transaction.patch(presentation, [
      patches.appendChild(
        presentation,
        presentation.root,
        '<p:custShowLst><p:custShow name="Short" id="0"><p:sldLst><p:sld r:id="rId3"/><p:sld r:id="rId5"/></p:sldLst></p:custShow></p:custShowLst>',
      ),
      patches.appendChild(
        presentation,
        presentation.root,
        '<p:extLst><p:ext uri="{521415D9-36F7-43E2-AB2F-B90AF26B5E84}"><p14:sectionLst xmlns:p14="http://schemas.microsoft.com/office/powerpoint/2010/main"><p14:section name="One" id="{11111111-1111-1111-1111-111111111111}"><p14:sldIdLst><p14:sldId id="256"/><p14:sldId id="257"/></p14:sldIdLst></p14:section><p14:section name="Two" id="{22222222-2222-2222-2222-222222222222}"><p14:sldIdLst><p14:sldId id="258"/></p14:sldIdLst></p14:section></p14:sectionLst></p:ext></p:extLst>',
      ),
    ]);
    await transaction.commit();
    const engine = await open(await pkg.save());
    await run(engine, [{ op: "insertSlide", index: 1 }]);
    let xml = await partText(engine, PRESENTATION);
    assert.ok(
      xml.includes(
        '<p14:sldId id="256"/><p14:sldId id="259"/><p14:sldId id="257"/>',
      ),
      xml,
    );
    await run(engine, [{ op: "deleteSlide", pageIndex: 0 }]);
    xml = await partText(engine, PRESENTATION);
    assert.ok(!xml.includes('<p14:sldId id="256"/>'));
    assert.ok(
      !xml.includes('<p:sld r:id="rId3"/>'),
      "the custom show drops the slide",
    );
    assert.ok(xml.includes('<p:sld r:id="rId5"/>'));
    await run(engine, [{ op: "duplicateSlide", pageIndex: 0, index: 0 }]);
    xml = await partText(engine, PRESENTATION);
    assert.ok(
      xml.includes('<p14:sldIdLst><p14:sldId id="260"/><p14:sldId id="259"/>'),
      xml,
    );
    const removed = await run(engine, [{ op: "deleteSlide", pageIndex: 3 }]);
    assert.deepEqual(removed.changedPages, []);
    assert.equal(removed.pageCount, 3);
    await engine.dispose();
  });

  it("keeps the deck index across slide-only commits and lists groups without a frame", async () => {
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
                paragraphs: [["A"]],
              }),
              `<p:grpSp><p:nvGrpSpPr><p:cNvPr id="3" name="Frameless"/><p:cNvGrpSpPr/><p:nvPr/></p:nvGrpSpPr><p:grpSpPr/>${textShape({ id: 4, x: 0, y: 0, cx: 457200, cy: 457200, paragraphs: [["Inside"]] })}</p:grpSp>`,
            ],
          },
          {
            shapes: [
              textShape({
                id: 2,
                x: 0,
                y: 0,
                cx: 10,
                cy: 10,
                paragraphs: [["B"]],
              }),
            ],
          },
        ],
      }),
    );
    const before = await engine.model(signal);
    assert.deepEqual(
      (await engine.getElements({ pageIndex: 0 }, signal)).map((item) => [
        item.id,
        item.kind,
        item.text,
      ]),
      [
        ["sld1:2", "shape", "A"],
        ["sld1:3", "group", undefined],
        ["sld1:4", "shape", "Inside"],
      ],
    );
    const batch: PptxOperation[] = [];
    for (let index = 0; index < 20; index += 1)
      batch.push({
        op: "replaceText",
        target: "sld1:2",
        text: `Edit ${index}`,
      });
    await run(engine, batch);
    assert.equal(
      await engine.model(signal),
      before,
      "a slide-only batch keeps the index",
    );
    assert.equal((await element(engine, "sld1:2")).text, "Edit 19");
    assert.equal((await element(engine, "sld2:2")).text, "B");
    await run(engine, [{ op: "moveSlide", from: 0, to: 1 }]);
    assert.notEqual(
      await engine.model(signal),
      before,
      "a presentation change rebuilds it",
    );
    assert.deepEqual(
      (await engine.slides(signal)).map((slide) => slide.key),
      ["sld2", "sld1"],
    );
    await engine.dispose();
  });
});
