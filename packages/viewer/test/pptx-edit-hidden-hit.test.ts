import assert from "node:assert/strict";
import { describe, it } from "node:test";

import type { TextRun } from "../src/contracts.js";
import { buildDeck, group, textShape } from "./fixtures/pptx-builder.js";
import { pptxSession } from "./fixtures/pptx-session.js";

const EMU_PER_PX = 9525;
const px = (value: number) => value * EMU_PER_PX;
const frame = (x: number, y: number, cx: number, cy: number) => ({
  x: px(x),
  y: px(y),
  cx: px(cx),
  cy: px(cy),
});

/*
 * One slide, back to front:
 *   2  "VISIBLE_TEXT" at (40, 100), 480 × 48;
 *   3  "HIDDEN_TEXT" in the same frame, hidden;
 *   4  a hidden group at (40, 300) holding 5, a visible shape;
 *   6  a short box at (40, 450) whose text a viewer paints down to y 520;
 *   7  a box at (40, 500), drawn over that painted text.
 */
const deck = buildDeck({
  slides: [
    {
      shapes: [
        textShape({
          id: 2,
          ...frame(40, 100, 480, 48),
          paragraphs: [["VISIBLE_TEXT"]],
        }),
        textShape({
          id: 3,
          ...frame(40, 100, 480, 48),
          hidden: true,
          paragraphs: [["HIDDEN_TEXT"]],
        }),
        group({
          id: 4,
          ...frame(40, 300, 480, 100),
          child: frame(40, 300, 480, 100),
          hidden: true,
          children: [
            textShape({
              id: 5,
              ...frame(40, 300, 480, 100),
              paragraphs: [["IN_HIDDEN_GROUP"]],
            }),
          ],
        }),
        textShape({
          id: 6,
          ...frame(40, 450, 480, 32),
          paragraphs: [["A line too long for its box, wrapping below it"]],
        }),
        textShape({
          id: 7,
          ...frame(40, 500, 480, 48),
          paragraphs: [["ON_TOP"]],
        }),
      ],
    },
  ],
});

/** A run the shown deck reports, laid out in the shape whose frame starts at `origin`. */
const run = (y: number, origin: { x: number; y: number }): TextRun => ({
  text: "painted",
  x: 50,
  y,
  width: 300,
  height: 20,
  shapeOrigin: origin,
});

const ids = (items: readonly { readonly id: string }[]) =>
  items.map((item) => item.id);

// ACTION-920: a hidden shape over a visible one took the click.
describe("PPTX elementsAt: only drawn elements are hit", () => {
  it("skips hidden shapes and shapes in hidden groups, which stay listed", async () => {
    const { session, end } = await pptxSession(deck);
    try {
      assert.deepEqual(
        ids((await session.elementsAt(0, { x: 60, y: 120 })).items),
        ["sld1:2"],
      );
      assert.deepEqual(
        (await session.elementsAt(0, { x: 60, y: 350 })).items,
        [],
      );
      const listed = (await session.getElements({ pageIndex: 0 })).items;
      assert.deepEqual(
        listed.filter((element) => element.hidden).map((element) => element.id),
        ["sld1:3", "sld1:4"],
      );
      assert.ok(listed.some((element) => element.id === "sld1:5"));
    } finally {
      await end();
    }
  });

  it("still edits a hidden shape named by id", async () => {
    const { session, end } = await pptxSession(deck);
    try {
      const hidden = (await session.getElement("sld1:3")).item!;
      assert.equal(hidden.hidden, true);
      assert.equal(hidden.text, "HIDDEN_TEXT");
      await session.replaceText({ target: "sld1:3", text: "STILL_HIDDEN" });
      assert.equal(
        (await session.getElement("sld1:3")).item?.text,
        "STILL_HIDDEN",
      );
    } finally {
      await end();
    }
  });

  it("orders text painted past a frame by what is drawn over it", async () => {
    const { session, end } = await pptxSession(deck, [
      [
        run(504, { x: 40, y: 450 }),
        // The hidden shape's text is not drawn, whatever the runs say.
        run(104, { x: 40, y: 100 }),
      ],
    ]);
    try {
      // Box 7 is drawn after box 6, so it is on top of 6's painted text.
      assert.deepEqual(
        ids((await session.elementsAt(0, { x: 60, y: 510 })).items),
        ["sld1:7", "sld1:6"],
      );
      assert.deepEqual(
        ids((await session.elementsAt(0, { x: 60, y: 110 })).items),
        ["sld1:2"],
      );
    } finally {
      await end();
    }
  });
});
