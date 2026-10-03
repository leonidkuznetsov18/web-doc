import assert from "node:assert/strict";
import { describe, it } from "node:test";

import type { TextRun } from "../src/contracts.js";
import { buildDeck, textShape } from "./fixtures/pptx-builder.js";
import { pptxSession } from "./fixtures/pptx-session.js";

const EMU_PER_PX = 9525;

/** A text box 480 × 48 px at (40, 100) whose 24 pt text wraps to a second line below it. */
const deck = buildDeck({
  slides: [
    {
      shapes: [
        textShape({
          id: 2,
          x: 40 * EMU_PER_PX,
          y: 100 * EMU_PER_PX,
          cx: 480 * EMU_PER_PX,
          cy: 48 * EMU_PER_PX,
          paragraphs: [["Hello and goodbye, a title too long for its box"]],
        }),
      ],
    },
  ],
});

/** A run the shown deck reports, laid out in the shape whose frame starts at `origin`. */
const run = (text: string, y: number, origin = { x: 40, y: 100 }): TextRun => ({
  text,
  x: 50,
  y,
  width: 300,
  height: 32,
  shapeOrigin: origin,
});

// ACTION-912: a click on the second line, painted 14 px below the frame,
// selected nothing.
describe("PPTX elementsAt: text painted past a shape's frame", () => {
  it("orders overlapping overflow hits front to back without duplicates", async () => {
    const shape = (id: number) =>
      textShape({
        id,
        x: 40 * EMU_PER_PX,
        y: 100 * EMU_PER_PX,
        cx: 480 * EMU_PER_PX,
        cy: 48 * EMU_PER_PX,
        paragraphs: [[id === 2 ? "Back" : "Front"]],
      });
    const { session, end } = await pptxSession(
      buildDeck({ slides: [{ shapes: [shape(2), shape(3)] }] }),
      [[run("Back overflow", 150), run("Front overflow", 150)]],
    );
    try {
      assert.deepEqual(
        (await session.elementsAt(0, { x: 60, y: 160 })).items.map((e) => e.id),
        ["sld1:3", "sld1:2"],
      );
    } finally {
      await end();
    }
  });

  it("hits the shape whose text is painted under the point, past its frame", async () => {
    const { session, end } = await pptxSession(deck, [
      [
        run("Hello and goodbye, a title", 104),
        run("too long for its box", 136),
      ],
    ]);
    try {
      const [shape] = (await session.getElements({ pageIndex: 0 })).items;
      // Inside the frame: the frame hits, as before.
      assert.deepEqual(
        (await session.elementsAt(0, { x: 60, y: 120 })).items.map((e) => e.id),
        [shape!.id],
      );
      // On the second line, below the frame.
      assert.deepEqual(
        (await session.elementsAt(0, { x: 60, y: 160 })).items.map((e) => e.id),
        [shape!.id],
      );
      // Below the text too: nothing.
      assert.deepEqual(
        (await session.elementsAt(0, { x: 60, y: 200 })).items,
        [],
      );
    } finally {
      await end();
    }
  });

  it("ignores runs of no known shape, and a headless session hits frames only", async () => {
    const stray = await pptxSession(deck, [
      [run("stray", 150, { x: 300, y: 300 })],
    ]);
    try {
      assert.deepEqual(
        (await stray.session.elementsAt(0, { x: 60, y: 160 })).items,
        [],
      );
    } finally {
      await stray.end();
    }
    const headless = await pptxSession(deck);
    try {
      assert.deepEqual(
        (await headless.session.elementsAt(0, { x: 60, y: 160 })).items,
        [],
      );
    } finally {
      await headless.end();
    }
  });
});
