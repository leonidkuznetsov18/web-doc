import assert from "node:assert/strict";
import { describe, it } from "node:test";

import { ViewerError } from "../src/index.js";
import { buildDocx, sectPr } from "./fixtures/docx-builder.js";
import { docxSession } from "./fixtures/docx-session.js";
import { buildDeck, textShape } from "./fixtures/pptx-builder.js";
import { pptxSession } from "./fixtures/pptx-session.js";

const RUN = (text: string, rPr = ""): string =>
  `<w:r>${rPr}<w:t xml:space="preserve">${text}</w:t></w:r>`;

const RANGE = (id: string, start: number, end: number) => ({
  start: { elementId: id, offset: start },
  end: { elementId: id, offset: end },
});

// "Plain " (0–6) "bold" (6–10) " tail" (10–15): a toggle on a range needs the
// style the range shows, not the first run's.
describe("getTextStyle: the style a range of text shows", () => {
  it("reads what every DOCX run in a range shares, and the run before a caret", async () => {
    const bytes = buildDocx({
      body:
        `<w:p>${RUN("Plain ")}${RUN("bold", "<w:rPr><w:b/><w:i/></w:rPr>")}${RUN(" tail", "<w:rPr><w:i/></w:rPr>")}</w:p>` +
        `<w:p><w:pPr><w:rPr><w:b/></w:rPr></w:pPr></w:p>` +
        sectPr(),
    });
    const { session, end } = await docxSession(bytes, [[]]);
    try {
      const [mixed, empty] = (await session.getElements()).items.map(
        (element) => element.id,
      );
      const style = async (id: string, start?: number, end?: number) =>
        (
          await session.getTextStyle({
            target: id,
            ...(start === undefined ? {} : { range: RANGE(id, start, end!) }),
          })
        ).item;

      // Inside the bold run: bold and italic.
      assert.equal((await style(mixed!, 7, 9))?.bold, true);
      assert.equal((await style(mixed!, 7, 9))?.italic, true);
      // Over plain and bold: neither bold nor italic is shared.
      const across = await style(mixed!, 4, 8);
      assert.equal("bold" in across!, false);
      assert.equal("italic" in across!, false);
      // What they share stays: the size.
      assert.equal(typeof across?.fontSize, "number");
      // Bold and the tail share italic only.
      const tail = await style(mixed!, 8, 13);
      assert.equal(tail?.italic, true);
      assert.equal("bold" in tail!, false);
      // A caret reads the run before it: at 10, the bold run.
      assert.equal((await style(mixed!, 10, 10))?.bold, true);
      assert.equal((await style(mixed!, 0, 0))?.bold, false);
      // The whole paragraph, and an empty one's mark.
      assert.equal("bold" in (await style(mixed!))!, false);
      assert.equal((await style(empty!))?.bold, true);
      // Past the text, nothing; a range off the target is refused.
      assert.equal(await style(mixed!, 0, 99), undefined);
      await assert.rejects(
        session.getTextStyle({ target: mixed!, range: RANGE(empty!, 0, 0) }),
        (error: unknown) =>
          error instanceof ViewerError && error.code === "invalid-operation",
      );
    } finally {
      await end();
    }
  });

  it("reads what every PPTX run in a range shares, across paragraphs too", async () => {
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
              paragraphs: [
                [
                  "Plain ",
                  { text: "bold", rPr: 'b="1" i="1"' },
                  { text: " tail", rPr: 'i="1"' },
                ],
                [{ text: "Second", rPr: 'b="1"' }],
              ],
            }),
          ],
        },
      ],
    });
    const { session, end } = await pptxSession(bytes);
    try {
      const [shape] = (
        await session.getElements({ pageIndex: 0 })
      ).items.filter((element) => element.text !== undefined);
      const id = shape!.id;
      const style = async (start: number, end: number) =>
        (
          await session.getTextStyle({
            target: id,
            range: RANGE(id, start, end),
          })
        ).item;

      assert.equal((await style(7, 9))?.bold, true);
      assert.equal("bold" in (await style(4, 8))!, false);
      assert.equal((await style(8, 13))?.italic, true);
      // "bold" through "Second": the second paragraph is bold, the tail is not.
      assert.equal("bold" in (await style(8, 19))!, false);
      // Only the second paragraph: bold.
      assert.equal((await style(16, 22))?.bold, true);
      assert.equal((await style(10, 10))?.bold, true);
    } finally {
      await end();
    }
  });
});
