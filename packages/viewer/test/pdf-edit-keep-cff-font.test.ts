import assert from "node:assert/strict";
import { describe, it } from "node:test";

import type { PageBitmap } from "../src/index.js";
import { cffInkTextPdf } from "./fixtures/pdf-fonts.js";
import { pdfSession } from "./fixtures/pdf-session.js";

/** The RGBA rows of a band 24 px tall ending 4 px under a baseline at `y` points, from x 60 to 160. */
function band(bitmap: PageBitmap, y: number): number[] {
  const bottom = bitmap.height - y + 4;
  const pixels: number[] = [];
  for (let row = bottom - 24; row < bottom; row += 1)
    pixels.push(
      ...bitmap.data.subarray(
        (row * bitmap.width + 60) * 4,
        (row * bitmap.width + 160) * 4,
      ),
    );
  return pixels;
}

function inked(pixels: readonly number[]): boolean {
  return pixels.some((value, at) => at % 4 !== 3 && value < 128);
}

describe("editing text in an embedded CFF subset", () => {
  for (const [text, reference] of [
    ["BA", 650],
    ["A B", 600],
  ] as const) {
    it(`keeps the subset for "${text}", which its glyphs and the PDF's widths cover`, async () => {
      const { session, end } = await pdfSession(cffInkTextPdf(), {
        fallbackFont: true,
      });
      try {
        const font = (await session.getTextFont("p0:o0")).item;
        assert.equal(font?.face?.format, "opentype");
        const receipt = await session.replaceText({ target: "p0:o0", text });
        assert.deepEqual(receipt.warnings, []);
        assert.equal((await session.getElement("p0:o0")).item?.text, text);
        assert.equal((await session.getTextFont("p0:o0")).item?.key, font.key);
        // The edited line draws exactly as the line that already said it.
        const page = (await session.renderPageWithout(0, [])).item!;
        assert.equal(page.width, 612);
        const edited = band(page, 700);
        assert.ok(inked(edited));
        assert.deepEqual(edited, band(page, reference));
      } finally {
        await end();
      }
    });
  }

  for (const [text, gap] of [
    ["AZ", "a character the subset has no glyph for"],
    ["A0", "a glyph the subset lacks although the PDF gives it a width"],
    ["Aa", "a glyph the PDF gives no width to"],
  ] as const) {
    it(`substitutes a covering font for ${gap}`, async () => {
      const { session, end } = await pdfSession(cffInkTextPdf(), {
        fallbackFont: true,
      });
      try {
        const receipt = await session.replaceText({ target: "p0:o0", text });
        assert.equal(receipt.warnings[0]?.code, "font-substitution");
        assert.equal((await session.getElement("p0:o0")).item?.text, text);
      } finally {
        await end();
      }
    });
  }
});
