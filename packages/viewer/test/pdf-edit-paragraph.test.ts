import assert from "node:assert/strict";
import { describe, it } from "node:test";

import {
  buildPdf,
  extractPageText,
  fixturePdfium,
} from "./fixtures/pdf-builder.js";
import { pdfSession } from "./fixtures/pdf-session.js";

const lines = [
  "Since 2013 our independent testing has tracked",
  "software quality across teams and organisations",
  "and shared the results with practitioners.",
];
const logicalText = lines.join(" ");

async function paragraphPdf() {
  return buildPdf([
    {
      texts: [
        { text: "Testing report", x: 90, y: 745, fontSize: 20 },
        ...lines.map((text, index) => ({
          text,
          x: 90,
          y: 700 - index * 20,
          fontSize: 11,
        })),
      ],
    },
  ]);
}

describe("native imported PDF paragraphs", () => {
  it("keeps a promoted paragraph addressable but refuses reflow while its page is rotated", async () => {
    const { session, end } = await pdfSession(await paragraphPdf());
    try {
      const row = (await session.getElements({ pageIndex: 0 })).items.find(
        (item) => item.text === lines[0],
      );
      assert.ok(row);
      const paragraph = (await session.getTextParagraph(row.id)).item;
      assert.ok(paragraph);
      await session.replaceParagraphText({
        target: paragraph.id,
        text: "A retained paragraph.",
      });
      const upright = (await session.getElement(paragraph.id)).item;
      assert.ok(upright);
      await session.rotatePage({ pageIndex: 0, by: 90 });
      const rotated = (await session.getElement(paragraph.id)).item;
      assert.ok(rotated);
      assert.equal(rotated.rotation, 90);
      assert.deepEqual(rotated.operations, ["deleteElement"]);
      assert.ok(Math.abs(rotated.bounds.width - upright.bounds.height) < 0.02);
      assert.equal(
        (await session.getTextParagraph(paragraph.id)).item,
        undefined,
      );
      await assert.rejects(
        session.replaceParagraphText({
          target: paragraph.id,
          text: "Cannot reflow sideways",
        }),
      );
      const reopened = await pdfSession((await session.save()).bytes);
      try {
        assert.equal(
          (await reopened.session.getElement(paragraph.id)).item?.rotation,
          90,
        );
        await reopened.session.rotatePage({ pageIndex: 0, by: 270 });
        assert.equal(
          (await reopened.session.getTextParagraph(paragraph.id)).item?.text,
          "A retained paragraph.",
        );
      } finally {
        await reopened.end();
      }
    } finally {
      await end();
    }
  });
  it("changes paragraph color in place without replacing or reflowing imported native rows", async () => {
    const { session, end } = await pdfSession(await paragraphPdf());
    try {
      const before = (await session.getElements({ pageIndex: 0 })).items;
      const first = before.find((element) => element.text === lines[0]);
      assert.ok(first);
      const paragraph = (await session.getTextParagraph(first.id)).item;
      assert.ok(paragraph);
      await session.setTextStyle({
        target: paragraph.id,
        style: { color: "#ff0000" },
      });
      const after = (await session.getElements({ pageIndex: 0 })).items;
      assert.deepEqual(
        after.map((element) => ({
          id: element.id,
          bounds: element.bounds,
          text: element.text,
          family: element.textStyle?.fontFamily,
        })),
        before.map((element) => ({
          id: element.id,
          bounds: element.bounds,
          text: element.text,
          family: element.textStyle?.fontFamily,
        })),
      );
      for (const id of paragraph.memberIds)
        assert.equal(
          (await session.getElement(id)).item?.textStyle?.color,
          "#ff0000",
        );
      assert.equal(
        (await session.getElement(paragraph.id)).item?.textStyle?.color,
        "#ff0000",
      );
      await session.undo();
      assert.deepEqual(
        (await session.getElements({ pageIndex: 0 })).items,
        before,
      );
    } finally {
      await end();
    }
  });
  it("uses effective typography for unit-size fonts scaled by the native text matrix", async () => {
    const pdfium = await fixturePdfium();
    const original = await buildPdf([
      { texts: lines.map((text) => ({ text, x: 0, y: 0, fontSize: 1 })) },
    ]);
    const document = pdfium.openDocument(original);
    const page = pdfium.lib.FPDF_LoadPage(document.handle, 0);
    let scaled: Uint8Array;
    try {
      lines.forEach((_, index) =>
        pdfium.lib.FPDFPageObj_Transform(
          pdfium.lib.FPDFPage_GetObject(page, index),
          11,
          0,
          0,
          11,
          90,
          700 - index * 20,
        ),
      );
      pdfium.lib.FPDFPage_GenerateContent(page);
      scaled = document.save("full");
    } finally {
      pdfium.lib.FPDF_ClosePage(page);
      document.close();
    }
    const { session, end } = await pdfSession(scaled);
    try {
      const row = (await session.getElements({ pageIndex: 0 })).items[0];
      assert.ok(row);
      const paragraph = (await session.getTextParagraph(row.id)).item;
      assert.ok(paragraph);
      assert.equal(paragraph.textStyle.fontSize, 11);
      assert.ok(Math.abs(paragraph.textStyle.lineHeight - 20 / 11) < 0.01);
      assert.ok(
        (await session.getTextLayout(paragraph.id)).item?.lines.every(
          (line) => line.fontSize === 11,
        ),
      );
      await session.replaceParagraphText({
        target: paragraph.id,
        text: logicalText + " Updated.",
      });
      const layout = (await session.getTextLayout(paragraph.id)).item;
      assert.ok(layout);
      assert.ok(layout.lines.every((line) => line.fontSize === 11));
      assert.ok(Math.abs(layout.lines[0]!.baseline.y - 92) < 0.02);
    } finally {
      await end();
    }
  });
  it("resolves every visual row to the complete logical paragraph without changing the file", async () => {
    const original = await paragraphPdf();
    const { session, end } = await pdfSession(original);
    try {
      const elements = (await session.getElements({ pageIndex: 0 })).items;
      const rows = elements.filter((element) =>
        lines.includes(element.text ?? ""),
      );
      assert.equal(rows.length, 3);
      const first = (await session.getTextParagraph(rows[0]!.id)).item;
      assert.ok(first);
      assert.equal(first.text, logicalText);
      assert.equal(first.textStyle.fontSize, 11);
      assert.ok(Math.abs(first.textStyle.lineHeight - 20 / 11) < 0.01);
      assert.ok(first.bounds.height > 40);
      assert.deepEqual(
        first.memberIds,
        rows.map((row) => row.id),
      );
      for (const row of rows)
        assert.deepEqual((await session.getTextParagraph(row.id)).item, first);
      for (const member of first.members) {
        const source = rows.find((row) => row.id === member.elementId);
        assert.equal(first.text.slice(member.start, member.end), source?.text);
      }
      assert.equal(
        (await session.getElement(first.id)).item?.text,
        logicalText,
      );
      assert.equal(session.state.dirty, false);
      assert.deepEqual((await session.save()).bytes, original);
    } finally {
      await end();
    }
  });

  it("reflows a replacement across the original rows as one undoable native change", async () => {
    const { session, end } = await pdfSession(await paragraphPdf());
    try {
      const rows = (await session.getElements({ pageIndex: 0 })).items;
      const row = rows.find((element) => element.text === lines[1]);
      assert.ok(row);
      const paragraph = (await session.getTextParagraph(row.id)).item;
      assert.ok(paragraph);
      const replacement = "reliable software quality and practical testing";
      const from = logicalText.indexOf("software quality");
      const to = from + "software quality".length;
      const expected =
        logicalText.slice(0, from) + replacement + logicalText.slice(to);
      const beforeRevision = session.state.revision;
      const originalRange = {
        start: { elementId: paragraph.id, offset: to + 1 },
        end: { elementId: paragraph.id, offset: to + 5 },
      };
      const originalRects = (
        await session.rangeRects({
          start: { elementId: paragraph.id, offset: 0 },
          end: { elementId: paragraph.id, offset: logicalText.length },
        })
      ).items;
      assert.equal(originalRects.length, 3);
      const receipt = await session.replaceParagraphText({
        target: paragraph.id,
        text: replacement,
        range: {
          start: { elementId: paragraph.id, offset: from },
          end: { elementId: paragraph.id, offset: to },
        },
      });
      assert.equal(receipt.revision, beforeRevision + 1);
      assert.deepEqual(receipt.changedPages, [0]);
      assert.deepEqual(receipt.warnings, []);
      assert.equal(
        (await session.mapRange(originalRange, beforeRevision)).item?.start
          .offset,
        originalRange.start.offset + replacement.length - (to - from),
      );
      const after = (await session.getTextParagraph(paragraph.id)).item;
      assert.ok(after);
      assert.equal(after.id, paragraph.id);
      assert.equal(after.text, expected);
      assert.equal(
        (await session.getElement(paragraph.id)).item?.text,
        expected,
      );
      const layout = (await session.getTextLayout(paragraph.id)).item;
      assert.ok(layout && layout.lines.length >= 4);
      const savedText = await extractPageText((await session.save()).bytes, 0);
      assert.equal((savedText.match(/Since 2013/g) ?? []).length, 1);
      assert.ok(savedText.includes("Testing report"));
      await session.undo();
      assert.deepEqual(
        (await session.mapRange(originalRange, beforeRevision)).item,
        originalRange,
      );
      assert.deepEqual(
        (await session.getElements({ pageIndex: 0 })).items,
        rows,
      );
      assert.equal(
        (await session.getTextParagraph(row.id)).item?.text,
        logicalText,
      );
      await session.redo();
      assert.equal(
        (await session.getTextParagraph(paragraph.id)).item?.text,
        expected,
      );
    } finally {
      await end();
    }
  });

  it("leaves mixed-style, rotated and ambiguous split lines independently editable", async () => {
    const fixtures = [
      {
        texts: lines.map((text, index) => ({
          text: ` ${text} `,
          x: 72,
          y: 700 - index * 20,
          fontSize: 11,
        })),
      },
      {
        texts: [
          { text: "First regular row", x: 72, y: 700, fontSize: 11 },
          {
            text: "A different bold run",
            x: 72,
            y: 680,
            fontSize: 11,
            font: "Helvetica-Bold" as const,
          },
          { text: "Last regular row", x: 72, y: 660, fontSize: 11 },
        ],
      },
      {
        texts: lines.map((text, index) => ({
          text,
          x: 72,
          y: 700 - index * 20,
          fontSize: 11,
        })),
        rotation: 1 as const,
      },
      {
        texts: [
          { text: "One partial row", x: 72, y: 700, fontSize: 11 },
          { text: "Second partial row", x: 72, y: 680, fontSize: 11 },
          { text: "overlaid run", x: 80, y: 690, fontSize: 11 },
        ],
      },
    ];
    for (const fixture of fixtures) {
      const { session, end } = await pdfSession(await buildPdf([fixture]));
      try {
        for (const element of (await session.getElements({ pageIndex: 0 }))
          .items) {
          assert.equal(element.textEditingTarget, undefined);
          assert.equal(
            (await session.getTextParagraph(element.id)).item,
            undefined,
          );
          assert.ok(element.operations.includes("replaceText"));
        }
      } finally {
        await end();
      }
    }
  });

  it("styles all paragraph rows atomically and refuses unsupported imported-text style fields", async () => {
    const { session, end } = await pdfSession(await paragraphPdf());
    try {
      const rows = (await session.getElements({ pageIndex: 0 })).items;
      const row = rows.find((element) => element.text === lines[0]);
      assert.ok(row);
      const paragraph = (await session.getTextParagraph(row.id)).item;
      assert.ok(paragraph);
      await session.setTextStyle({
        target: paragraph.id,
        style: { color: "#ff0000", fontSize: 12 },
      });
      const after = (await session.getTextParagraph(paragraph.id)).item;
      assert.ok(after);
      assert.equal(after.textStyle.color, "#ff0000");
      assert.equal(after.textStyle.fontSize, 12);
      const layout = (await session.getTextLayout(paragraph.id)).item;
      assert.ok(layout && layout.lines.length > 1);
      assert.ok(
        layout.lines.every(
          (line) => line.color === "#ff0000" && line.fontSize === 12,
        ),
      );
      await assert.rejects(
        session.setTextStyle({
          target: paragraph.id,
          style: { fontFamily: "Courier" },
        }),
      );
      assert.equal(
        (await session.getTextParagraph(paragraph.id)).item?.textStyle
          .fontFamily,
        paragraph.textStyle.fontFamily,
      );
      await session.undo();
      assert.deepEqual(
        (await session.getElements({ pageIndex: 0 })).items,
        rows,
      );
    } finally {
      await end();
    }
  });
});
