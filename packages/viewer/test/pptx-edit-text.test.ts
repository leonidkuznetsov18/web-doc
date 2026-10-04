import assert from "node:assert/strict";
import { describe, it } from "node:test";

import { PptxEditEngine } from "../src/edit/pptx/engine.js";
import { OoxmlPackage } from "../src/edit/ooxml/package.js";
import {
  defaultResourceLimits,
  type PptxElement,
  type PptxOperation,
} from "../src/index.js";
import { buildDeck, textShape } from "./fixtures/pptx-builder.js";
import { pptxSession } from "./fixtures/pptx-session.js";

/*
 * Task 45 of the PPTX module: replaceText (whole and ranged) and
 * setTextStyle rebuild only the paragraphs they touch, keep every untouched
 * run's bytes and the first run's properties, split paragraphs on "\n" and
 * write "\v" as a:br, keep fields whole, drop a stale autofit scale, and
 * write colours in their theme form.
 */

const signal = new AbortController().signal;
const SLIDE = "/ppt/slides/slide1.xml";

const run = (engine: PptxEditEngine, operations: PptxOperation[]) =>
  engine.apply(operations, signal);
const check = (engine: PptxEditEngine, operations: PptxOperation[]) =>
  engine.validate(operations, signal);

async function open(bytes: Uint8Array): Promise<PptxEditEngine> {
  return PptxEditEngine.open(bytes, defaultResourceLimits, signal);
}

async function slideXml(engine: PptxEditEngine): Promise<string> {
  const pkg = await OoxmlPackage.open(
    await engine.materialize("save", {}, signal),
    { limits: defaultResourceLimits },
  );
  return new TextDecoder().decode(await pkg.part(SLIDE));
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

const RANGE = (id: string, start: number, end: number) => ({
  start: { elementId: id, offset: start },
  end: { elementId: id, offset: end },
});

describe("PPTX text operations (pptx-edit)", () => {
  it("appends pending bold, italic and underline in one worker batch with isolated dry run and atomic history", async () => {
    const bytes = deck(
      textShape({
        id: 2,
        x: 0,
        y: 0,
        cx: 914400,
        cy: 914400,
        paragraphs: [["Original"]],
      }),
    );
    const { session, end } = await pptxSession(bytes);
    const operations: PptxOperation[] = [
      {
        op: "replaceText",
        target: "sld1:2",
        text: "X",
        range: RANGE("sld1:2", 8, 8),
      },
      {
        op: "setTextStyle",
        target: "sld1:2",
        style: { bold: true, italic: true, underline: true },
        range: RANGE("sld1:2", 8, 9),
      },
    ];
    try {
      const state = session.state;
      assert.equal(
        (await session.applyJson(operations, { dryRun: true })).dryRun,
        true,
      );
      assert.deepEqual(session.state, state);
      assert.deepEqual((await session.save()).bytes, bytes);
      await session.applyJson(operations);
      assert.equal(
        (await session.getElement("sld1:2")).item?.text,
        "OriginalX",
      );
      const inserted = (
        await session.getTextStyle({
          target: "sld1:2",
          range: RANGE("sld1:2", 8, 9),
        })
      ).item;
      assert.equal(inserted?.bold, true);
      assert.equal(inserted?.italic, true);
      assert.equal(inserted?.underline, true);
      const original = (
        await session.getTextStyle({
          target: "sld1:2",
          range: RANGE("sld1:2", 0, 8),
        })
      ).item;
      assert.notEqual(original?.bold, true);
      assert.notEqual(original?.italic, true);
      assert.notEqual(original?.underline, true);
      const edited = (await session.save()).bytes;
      await session.undo();
      assert.deepEqual((await session.save()).bytes, bytes);
      assert.equal(session.state.canUndo, false);
      await session.redo();
      assert.deepEqual((await session.save()).bytes, edited);
    } finally {
      await end();
    }
  });

  it("rolls back a replacement when its later style range is invalid, retaining redo history even for dry run", async () => {
    const bytes = deck(
      textShape({
        id: 2,
        x: 0,
        y: 0,
        cx: 914400,
        cy: 914400,
        paragraphs: [["Original"]],
      }),
    );
    const { session, end } = await pptxSession(bytes);
    try {
      await session.replaceText({ target: "sld1:2", text: "History" });
      const historyBytes = (await session.save()).bytes;
      await session.undo();
      const state = session.state;
      for (const dryRun of [false, true]) {
        for (const range of [RANGE("sld1:2", 7, 8), RANGE("sld1:99", 0, 1)]) {
          const operations: PptxOperation[] = [
            { op: "replaceText", target: "sld1:2", text: "A" },
            {
              op: "setTextStyle",
              target: "sld1:2",
              style: { bold: true },
              range,
            },
          ];
          await assert.rejects(session.applyJson(operations, { dryRun }), {
            code: "invalid-operation",
          });
          assert.deepEqual(session.state, state);
          assert.deepEqual((await session.save()).bytes, bytes);
          assert.equal(
            (await session.getElement("sld1:2")).item?.text,
            "Original",
          );
        }
      }
      await session.redo();
      assert.deepEqual((await session.save()).bytes, historyBytes);
    } finally {
      await end();
    }
  });

  it("replaces the whole text keeping the first run's properties, the paragraph properties and the body", async () => {
    const engine = await open(
      deck(
        textShape({
          id: 2,
          x: 0,
          y: 0,
          cx: 914400,
          cy: 914400,
          pPr: 'algn="ctr"',
          bodyPr: '<a:normAutofit fontScale="90000" lnSpcReduction="10000"/>',
          lstStyle:
            '<a:lstStyle><a:lvl1pPr><a:defRPr sz="2400"/></a:lvl1pPr></a:lstStyle>',
          paragraphs: [
            [{ text: "Bold", rPr: 'b="1" sz="2000"' }, " and plain"],
            ["Second paragraph"],
          ],
        }),
      ),
    );
    const before = await slideXml(engine);
    const change = await run(engine, [
      {
        op: "replaceText",
        target: "sld1:2",
        text: "New\u000bline\nSecond\nThird",
      },
    ]);
    assert.deepEqual(change.changedPages, [0]);
    assert.deepEqual(change.createdIds, []);
    const shape = await element(engine, "sld1:2");
    assert.equal(shape.text, "New\u000bline\nSecond\nThird");
    assert.equal(shape.textStyle?.bold, true);
    assert.equal(shape.textStyle?.fontSize, 20);
    assert.equal(shape.textStyle?.align, "center");
    const after = await slideXml(engine);
    const body = /<p:txBody>.*<\/p:txBody>/s.exec(after)![0];
    assert.ok(
      body.includes(
        '<a:bodyPr wrap="square" rtlCol="0"><a:normAutofit/></a:bodyPr>',
      ),
      "autofit scale dropped",
    );
    assert.ok(
      body.includes(
        '<a:lstStyle><a:lvl1pPr><a:defRPr sz="2400"/></a:lvl1pPr></a:lstStyle>',
      ),
      "lstStyle kept",
    );
    assert.equal(
      body.slice(body.indexOf("<a:p>")),
      '<a:p><a:pPr algn="ctr"/><a:r><a:rPr lang="en-US" b="1" sz="2000"/><a:t>New</a:t></a:r><a:br><a:rPr lang="en-US" b="1" sz="2000"/></a:br><a:r><a:rPr lang="en-US" b="1" sz="2000"/><a:t>line</a:t></a:r><a:endParaRPr lang="en-US"/></a:p>' +
        '<a:p><a:pPr algn="ctr"/><a:r><a:rPr lang="en-US" b="1" sz="2000"/><a:t>Second</a:t></a:r><a:endParaRPr lang="en-US"/></a:p>' +
        '<a:p><a:pPr algn="ctr"/><a:r><a:rPr lang="en-US" b="1" sz="2000"/><a:t>Third</a:t></a:r><a:endParaRPr lang="en-US"/></a:p></p:txBody>',
    );
    // Nothing outside the text body changed.
    assert.equal(
      before.slice(0, before.indexOf("<p:txBody>")),
      after.slice(0, after.indexOf("<p:txBody>")),
    );
    assert.equal(
      before.slice(before.indexOf("</p:txBody>")),
      after.slice(after.indexOf("</p:txBody>")),
    );
    await engine.dispose();
  });

  it("replaces a range inside one run by changing only that run's text", async () => {
    const engine = await open(
      deck(
        textShape({
          id: 2,
          x: 0,
          y: 0,
          cx: 914400,
          cy: 914400,
          paragraphs: [["Hello ", { text: "big", rPr: 'b="1"' }, " world"]],
        }),
      ),
    );
    const before = await slideXml(engine);
    await run(engine, [
      {
        op: "replaceText",
        target: "sld1:2",
        text: "small",
        range: RANGE("sld1:2", 6, 9),
      },
    ]);
    const after = await slideXml(engine);
    assert.equal((await element(engine, "sld1:2")).text, "Hello small world");
    assert.equal(after, before.replace("<a:t>big</a:t>", "<a:t>small</a:t>"));
    await engine.dispose();
  });

  it("inserts at a collapsed range with the style of the run before it, and across runs keeps the first touched run's style", async () => {
    const engine = await open(
      deck(
        textShape({
          id: 2,
          x: 0,
          y: 0,
          cx: 914400,
          cy: 914400,
          paragraphs: [
            [
              { text: "Red", rPr: 'i="1"' },
              { text: "Blue", rPr: 'u="sng"' },
            ],
          ],
        }),
      ),
    );
    await run(engine, [
      {
        op: "replaceText",
        target: "sld1:2",
        text: "+",
        range: RANGE("sld1:2", 3, 3),
      },
    ]);
    assert.equal((await element(engine, "sld1:2")).text, "Red+Blue");
    let xml = await slideXml(engine);
    assert.ok(
      xml.includes(
        '<a:r><a:rPr lang="en-US" i="1"/><a:t>Red</a:t></a:r><a:r><a:rPr lang="en-US" i="1"/><a:t>+</a:t></a:r><a:r><a:rPr lang="en-US" u="sng"/><a:t>Blue</a:t></a:r>',
      ),
      xml,
    );
    // "d+Bl" → "X": the italic run keeps "Re", takes "X", the underlined run keeps "ue".
    await run(engine, [
      {
        op: "replaceText",
        target: "sld1:2",
        text: "X",
        range: RANGE("sld1:2", 2, 6),
      },
    ]);
    assert.equal((await element(engine, "sld1:2")).text, "ReXue");
    xml = await slideXml(engine);
    assert.ok(
      xml.includes(
        '<a:r><a:rPr lang="en-US" i="1"/><a:t>Re</a:t></a:r><a:r><a:rPr lang="en-US" i="1"/><a:t>X</a:t></a:r><a:r><a:rPr lang="en-US" u="sng"/><a:t>ue</a:t></a:r>',
      ),
      xml,
    );
    // Deleting across the paragraph break merges the paragraphs under the first one's properties.
    const merged = await open(
      deck(
        textShape({
          id: 2,
          x: 0,
          y: 0,
          cx: 914400,
          cy: 914400,
          paragraphs: [["One"], ["Two"]],
        }),
      ),
    );
    await run(merged, [
      {
        op: "replaceText",
        target: "sld1:2",
        text: "",
        range: RANGE("sld1:2", 2, 5),
      },
    ]);
    assert.equal((await element(merged, "sld1:2")).text, "Onwo");
    assert.equal((await slideXml(merged)).match(/<a:p>/g)!.length, 1);
    await merged.dispose();
    await engine.dispose();
  });

  it("keeps fields whole, fills an empty body and refuses bad ranges and text", async () => {
    const engine = await open(
      deck(
        textShape({
          id: 2,
          x: 0,
          y: 0,
          cx: 914400,
          cy: 914400,
          paragraphs: [["Page "]],
        }).replace(
          "<a:endParaRPr",
          '<a:fld id="{B6F7A1F5-1C3F-4F57-9A3E-000000000001}" type="slidenum"><a:rPr lang="en-US"/><a:t>7</a:t></a:fld><a:endParaRPr',
        ),
        textShape({
          id: 3,
          x: 0,
          y: 0,
          cx: 914400,
          cy: 914400,
          paragraphs: [[]],
        }),
      ),
    );
    assert.equal((await element(engine, "sld1:2")).text, "Page 7");
    const issues = await check(engine, [
      {
        op: "replaceText",
        target: "sld1:2",
        text: "x",
        range: RANGE("sld1:2", 5, 6),
      },
      {
        op: "replaceText",
        target: "sld1:2",
        text: "x",
        range: RANGE("sld1:2", 9, 9),
      },
      {
        op: "replaceText",
        target: "sld1:2",
        text: "x",
        range: RANGE("sld1:3", 0, 0),
      },
      { op: "replaceText", target: "sld1:9", text: "x" },
      { op: "replaceText", target: "sld1:2", text: "bad\u0001" },
    ]);
    assert.deepEqual(
      issues.map((issue) => [issue.operationIndex, issue.path, issue.code]),
      [
        [1, "/range", "invalid-range"],
        [2, "/range", "invalid-range"],
        [3, "/target", "unknown-target"],
        [4, "/text", "invalid-text"],
      ],
    );
    // A range ending exactly at the field's start replaces only the text before it.
    await run(engine, [
      {
        op: "replaceText",
        target: "sld1:2",
        text: "Slide ",
        range: RANGE("sld1:2", 0, 5),
      },
    ]);
    assert.equal((await element(engine, "sld1:2")).text, "Slide 7");
    assert.ok(
      (await slideXml(engine)).includes(
        'type="slidenum"><a:rPr lang="en-US"/><a:t>7</a:t></a:fld>',
      ),
    );
    // The empty paragraph takes its text from the paragraph end properties.
    await run(engine, [
      { op: "replaceText", target: "sld1:3", text: "Filled" },
    ]);
    assert.equal((await element(engine, "sld1:3")).text, "Filled");
    assert.ok(
      (await slideXml(engine)).includes(
        '<a:r><a:rPr lang="en-US"/><a:t>Filled</a:t></a:r><a:endParaRPr lang="en-US"/>',
      ),
    );
    await engine.dispose();
  });

  it("styles a range by splitting runs, writes colours in theme form and aligns the paragraphs touched", async () => {
    const engine = await open(
      deck(
        textShape({
          id: 2,
          x: 0,
          y: 0,
          cx: 914400,
          cy: 914400,
          paragraphs: [
            [
              {
                text: "Hello world",
                rPr: 'sz="1800"><a:solidFill><a:srgbClr val="112233"/></a:solidFill><a:latin typeface="Arial"/></a:rPr><a:t>x</a:t></a:r><a:r><a:rPr lang="en-US" sz="1800"',
              },
            ],
            ["Second"],
          ],
        }),
      ),
    );
    // The fixture trick above produced two runs: "x" (coloured, Arial) and "Hello world".
    const shape = await element(engine, "sld1:2");
    assert.equal(shape.text, "xHello world\nSecond");
    assert.deepEqual(shape.textStyle?.color, "#112233");
    assert.equal(shape.textStyle?.fontFamily, "Arial");

    await run(engine, [
      {
        op: "setTextStyle",
        target: "sld1:2",
        range: RANGE("sld1:2", 0, 6),
        style: {
          bold: true,
          fontSize: 24,
          color: { theme: "accent1", mods: { lumMod: 75000 } },
          fontFamily: "+mj-lt",
          underline: true,
        },
      },
    ]);
    const styled = await element(engine, "sld1:2");
    assert.deepEqual(styled.textStyle, {
      fontFamily: "Calibri Light",
      fontSize: 24,
      bold: true,
      italic: false,
      underline: true,
      color: { theme: "accent1", mods: { lumMod: 75000 } },
      align: "left",
    });
    let xml = await slideXml(engine);
    assert.ok(
      xml.includes(
        '<a:r><a:rPr lang="en-US" sz="2400" b="1" u="sng"><a:solidFill><a:schemeClr val="accent1"><a:lumMod val="75000"/></a:schemeClr></a:solidFill><a:latin typeface="+mj-lt"/></a:rPr><a:t>x</a:t></a:r>' +
          '<a:r><a:rPr lang="en-US" sz="2400" b="1" u="sng"><a:solidFill><a:schemeClr val="accent1"><a:lumMod val="75000"/></a:schemeClr></a:solidFill><a:latin typeface="+mj-lt"/></a:rPr><a:t>Hello</a:t></a:r>' +
          '<a:r><a:rPr lang="en-US" sz="1800"/><a:t> world</a:t></a:r>',
      ),
      xml,
    );
    assert.ok(
      xml.includes('<a:p><a:r><a:rPr lang="en-US"/><a:t>Second</a:t></a:r>'),
      "the second paragraph is untouched",
    );

    await run(engine, [
      {
        op: "setTextStyle",
        target: "sld1:2",
        style: { align: "right", italic: true },
      },
    ]);
    xml = await slideXml(engine);
    assert.equal(xml.match(/<a:pPr algn="r"\/>/g)!.length, 2);
    assert.equal(
      xml.match(/ i="1"/g)!.length,
      6,
      "every run and paragraph end is italic",
    );
    assert.equal((await element(engine, "sld1:2")).textStyle?.align, "right");

    const issues = await check(engine, [
      { op: "setTextStyle", target: "sld1:2", style: { color: "red" } },
      {
        op: "setTextStyle",
        target: "sld1:2",
        style: { color: { theme: "accent9" } },
      },
      {
        op: "setTextStyle",
        target: "sld1:2",
        style: { color: { theme: "tx1", mods: { bogus: 1 } } },
      },
    ]);
    assert.deepEqual(
      issues.map((issue) => [issue.operationIndex, issue.path, issue.code]),
      [
        [0, "/style/color", "invalid-value"],
        [1, "/style/color", "invalid-value"],
        [2, "/style/color", "invalid-value"],
      ],
    );
    await engine.dispose();
  });

  it("round-trips through the session: undo returns the original bytes, redo replays, a dry run changes nothing", async () => {
    const bytes = deck(
      textShape({
        id: 2,
        x: 0,
        y: 0,
        cx: 914400,
        cy: 914400,
        paragraphs: [["Alpha"]],
      }),
    );
    const { session, end } = await pptxSession(bytes);
    try {
      const dry = await session.replaceText(
        { target: "sld1:2", text: "Dry" },
        { dryRun: true },
      );
      assert.equal(dry.dryRun, true);
      assert.deepEqual((await session.save()).bytes, bytes);
      const receipt = await session.replaceText({
        target: "sld1:2",
        text: "Beta",
      });
      assert.equal(receipt.revision, 1);
      assert.deepEqual(receipt.changedPages, [0]);
      const edited = (await session.save()).bytes;
      assert.notDeepEqual(edited, bytes);
      assert.equal((await session.getElement("sld1:2")).item?.text, "Beta");
      await session.setTextStyle({ target: "sld1:2", style: { bold: true } });
      await session.undo();
      await session.undo();
      assert.deepEqual((await session.save()).bytes, bytes);
      await session.redo();
      assert.deepEqual((await session.save()).bytes, edited);
      const found = await session.findText("beta");
      assert.equal(found.items.length, 1);
      assert.deepEqual(found.items[0]!.ranges[0]!.end, {
        elementId: "sld1:2",
        offset: 4,
      });
    } finally {
      await end();
    }
  });
});
