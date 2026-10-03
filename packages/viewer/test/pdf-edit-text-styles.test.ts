import assert from "node:assert/strict";
import { createHash } from "node:crypto";
import { readFileSync } from "node:fs";
import { describe, it } from "node:test";

import type { PdfEditSession } from "../src/index.js";
import { buildPdf, fixturePdfium } from "./fixtures/pdf-builder.js";
import { pdfSession } from "./fixtures/pdf-session.js";

function underline(session: PdfEditSession, target: string) {
  const operation = { op: "setTextStyle", target, style: { underline: true } };
  return session.applyJson([operation]);
}

async function raster(session: PdfEditSession, hidden: readonly string[] = []) {
  const bitmap = (await session.renderPageWithout(0, hidden, { scale: 1 }))
    .item;
  assert.ok(bitmap);
  return createHash("sha256").update(bitmap.data).digest("hex");
}

async function nativeObjects(bytes: Uint8Array) {
  const pdfium = await fixturePdfium();
  const document = pdfium.openDocument(bytes);
  const page = pdfium.lib.FPDF_LoadPage(document.handle, 0);
  try {
    return Array.from(
      { length: pdfium.lib.FPDFPage_CountObjects(page) },
      (_, index) => {
        const object = pdfium.lib.FPDFPage_GetObject(page, index);
        return {
          type: pdfium.lib.FPDFPageObj_GetType(object),
          fill: pdfium.readNumbers(
            4,
            "i32",
            ([r, g, b, a]) =>
              r !== undefined &&
              g !== undefined &&
              b !== undefined &&
              a !== undefined &&
              pdfium.lib.FPDFPageObj_GetFillColor(object, r, g, b, a),
          ),
          matrix: pdfium.readNumbers(6, "float", ([pointer]) => {
            assert.ok(pointer);
            return pdfium.lib.FPDFPageObj_GetMatrix(object, pointer);
          }),
        };
      },
    );
  } finally {
    pdfium.lib.FPDF_ClosePage(page);
    document.close();
  }
}

const paragraphLines = [
  "Since 2013 our independent testing has tracked",
  "software quality across teams and organisations",
  "and shared the results with practitioners.",
];

async function embeddedRow(text: string): Promise<Uint8Array> {
  const pdfium = await fixturePdfium();
  const { lib } = pdfium;
  const document = pdfium.createDocument();
  const data = new Uint8Array(
    readFileSync(
      new URL("../../fonts/noto-sans-latin-cyrillic.ttf", import.meta.url),
    ),
  );
  const pointer = pdfium.writeBytes(data);
  try {
    const page = lib.FPDFPage_New(document.handle, 0, 612, 792);
    try {
      const font = lib.FPDFText_LoadFont(
        document.handle,
        pointer,
        data.length,
        1,
        true,
      );
      const object = lib.FPDFPageObj_CreateTextObj(document.handle, font, 18);
      const wide = pdfium.writeWideString(text);
      try {
        assert.ok(lib.FPDFText_SetText(object, wide));
      } finally {
        pdfium.free(wide);
      }
      lib.FPDFPageObj_SetFillColor(object, 30, 60, 90, 160);
      lib.FPDFPageObj_Transform(object, 1, 0, 0, 1, 72, 700);
      lib.FPDFPage_InsertObject(page, object);
      lib.FPDFPage_GenerateContent(page);
    } finally {
      lib.FPDF_ClosePage(page);
    }
    return document.save("full");
  } finally {
    pdfium.free(pointer);
    document.close();
  }
}

describe("native PDF basic text styles", () => {
  it("retains a true bold face when typing new Unicode into imported rows and paragraphs", async () => {
    for (const kind of ["row", "paragraph"] as const) {
      const original = await buildPdf([
        {
          texts: (kind === "row" ? ["Original title"] : paragraphLines).map(
            (text, index) => ({
              text,
              x: 90,
              y: 700 - index * 20,
              fontSize: 11,
            }),
          ),
        },
      ]);
      const { session, end } = await pdfSession(original, {
        fallbackFont: true,
      });
      try {
        const id =
          kind === "row"
            ? "p0:o0"
            : (await session.getTextParagraph("p0:o0")).item?.id;
        assert.ok(id);
        await session.setTextStyle({ target: id, style: { bold: true } });
        await session.replaceText({ target: id, text: "Привіт світе" });
        assert.equal(
          (await session.getElement(id)).item?.textStyle?.bold,
          true,
        );
        const reopened = await pdfSession((await session.save()).bytes, {
          fallbackFont: true,
        });
        try {
          await reopened.session.replaceText({
            target: id,
            text: "Привіт знову",
          });
          assert.equal(
            (await reopened.session.getElement(id)).item?.textStyle?.bold,
            true,
          );
        } finally {
          await reopened.end();
        }
      } finally {
        await end();
      }
    }
  });

  it("refuses new face and underline changes on non-filled native text atomically", async () => {
    const pdfium = await fixturePdfium();
    for (const mode of [1, 3]) {
      const document = pdfium.openDocument(await buildPdf(["Native mode"]));
      const page = pdfium.lib.FPDF_LoadPage(document.handle, 0);
      let bytes: Uint8Array;
      try {
        assert.ok(
          pdfium.lib.FPDFTextObj_SetTextRenderMode(
            pdfium.lib.FPDFPage_GetObject(page, 0),
            mode,
          ),
        );
        assert.ok(pdfium.lib.FPDFPage_GenerateContent(page));
        bytes = document.save("full");
      } finally {
        pdfium.lib.FPDF_ClosePage(page);
        document.close();
      }
      const { session, end } = await pdfSession(bytes);
      try {
        const pixels = await raster(session);
        for (const style of [
          { bold: true },
          { italic: true },
          { underline: true },
        ]) {
          await assert.rejects(
            session.setTextStyle({ target: "p0:o0", style }),
            /filled text/,
          );
          assert.deepEqual((await session.save()).bytes, bytes);
          assert.equal(await raster(session), pixels);
          assert.equal(session.state.canUndo, false);
        }
      } finally {
        await end();
      }
    }
  });

  it("preserves the existing empty-row rejection without dropping underline or history", async () => {
    const { session, end } = await pdfSession(await buildPdf(["Clear this"]));
    try {
      await underline(session, "p0:o0");
      const before = (await session.save()).bytes;
      const pixels = await raster(session);
      await assert.rejects(
        session.replaceText({ target: "p0:o0", text: "" }),
        /at least 1/,
      );
      assert.deepEqual((await session.save()).bytes, before);
      assert.equal(await raster(session), pixels);
      assert.equal(
        (await session.getElement("p0:o0")).item?.textStyle?.underline,
        true,
      );
      await session.undo();
      assert.equal(
        (await session.getElement("p0:o0")).item?.textStyle?.underline,
        undefined,
      );
    } finally {
      await end();
    }
  });

  it("rejects a styled textbox when no true face covers its text atomically", async () => {
    const original = await buildPdf(["Neighbor"]);
    const { session, end } = await pdfSession(original, { fallbackFont: true });
    try {
      await assert.rejects(
        session.insertTextBox({
          pageIndex: 0,
          rect: { x: 72, y: 180, width: 200, height: 80 },
          text: "日本語",
          style: { bold: true },
        }),
        /font|draw|cover/,
      );
      assert.deepEqual((await session.save()).bytes, original);
      assert.equal(session.state.canUndo, false);
    } finally {
      await end();
    }
  });

  it("keeps a real styled face when an inserted Unicode textbox is saved and typed into again", async () => {
    const { session, end } = await pdfSession(await buildPdf(["Neighbor"]), {
      fallbackFont: true,
    });
    try {
      const receipt = await session.insertTextBox({
        pageIndex: 0,
        rect: { x: 72, y: 180, width: 200, height: 80 },
        text: "Привіт світе",
      });
      const [id] = receipt.createdIds;
      assert.ok(id);
      await session.setTextStyle({
        target: id,
        style: { bold: true, italic: true },
      });
      const styled = (await session.getElement(id)).item;
      assert.equal(styled?.textStyle?.bold, true);
      assert.equal(styled?.textStyle?.italic, true);
      const reopened = await pdfSession((await session.save()).bytes, {
        fallbackFont: true,
      });
      try {
        await reopened.session.replaceText({
          target: id,
          text: "Привіт знову",
        });
        const typed = (await reopened.session.getElement(id)).item;
        assert.equal(typed?.textStyle?.bold, true);
        assert.equal(typed?.textStyle?.italic, true);
        assert.equal(typed?.text, "Привіт знову");
      } finally {
        await reopened.end();
      }
    } finally {
      await end();
    }
  });

  it("rejects non-axis-aligned underline geometry without changing the native text", async () => {
    const original = await buildPdf([
      { texts: [{ text: "Angled title", matrix: [0.866, 0.5, -0.5, 0.866] }] },
    ]);
    const { session, end } = await pdfSession(original);
    try {
      await assert.rejects(underline(session, "p0:o0"), /underline geometry/);
      assert.deepEqual((await session.save()).bytes, original);
      assert.equal(session.state.canUndo, false);
    } finally {
      await end();
    }
  });

  it("keeps native glyph placement and underline ownership on quarter-turned pages", async () => {
    for (const rotation of [1, 2, 3] as const) {
      const { session, end } = await pdfSession(
        await buildPdf([{ text: "Turned title", rotation }]),
      );
      try {
        const layout = (await session.getTextLayout("p0:o0")).item;
        const before = await raster(session);
        await underline(session, "p0:o0");
        assert.deepEqual((await session.getTextLayout("p0:o0")).item, layout);
        assert.equal(
          (await session.getElement("p0:o0")).item?.rotation,
          rotation * 90,
        );
        assert.notEqual(await raster(session), before);
        const reopened = await pdfSession((await session.save()).bytes);
        try {
          assert.equal(
            (await reopened.session.getElement("p0:o0")).item?.textStyle
              ?.underline,
            true,
          );
          assert.equal(await raster(reopened.session), await raster(session));
        } finally {
          await reopened.end();
        }
      } finally {
        await end();
      }
    }
  });

  it("rejects overflowing paragraph formatting atomically and preserves its native history", async () => {
    const original = await buildPdf([
      {
        texts: paragraphLines.map((text, index) => ({
          text,
          x: 90,
          y: 700 - index * 20,
          fontSize: 11,
        })),
      },
    ]);
    const { session, end } = await pdfSession(original);
    try {
      const paragraph = (await session.getTextParagraph("p0:o0")).item;
      assert.ok(paragraph);
      const before = await raster(session);
      await assert.rejects(
        session.setTextStyle({
          target: paragraph.id,
          style: { bold: true, fontSize: 500 },
        }),
        /paragraph would overlap|leave the page/,
      );
      assert.deepEqual((await session.save()).bytes, original);
      assert.equal(await raster(session), before);
      assert.equal(session.state.canUndo, false);
    } finally {
      await end();
    }
  });

  it("keeps an empty underlined paragraph addressable and restores its native paths when typing resumes", async () => {
    const { session, end } = await pdfSession(
      await buildPdf([
        {
          texts: paragraphLines.map((text, index) => ({
            text,
            x: 90,
            y: 700 - index * 20,
            fontSize: 11,
          })),
        },
      ]),
    );
    try {
      const paragraph = (await session.getTextParagraph("p0:o0")).item;
      assert.ok(paragraph);
      await underline(session, paragraph.id);
      await session.replaceParagraphText({ target: paragraph.id, text: "" });
      const reopened = await pdfSession((await session.save()).bytes);
      try {
        assert.equal(
          (await reopened.session.getTextParagraph(paragraph.id)).item?.text,
          "",
        );
        await reopened.session.replaceParagraphText({
          target: paragraph.id,
          text: "Typed again",
        });
        assert.equal(
          (await reopened.session.getTextParagraph(paragraph.id)).item
            ?.textStyle.underline,
          true,
        );
        assert.equal(
          (await nativeObjects((await reopened.session.save()).bytes)).filter(
            (object) => object.type === 2,
          ).length,
          1,
        );
        await reopened.session.undo();
        assert.equal(
          (await reopened.session.getTextParagraph(paragraph.id)).item?.text,
          "",
        );
      } finally {
        await reopened.end();
      }
    } finally {
      await end();
    }
  });
  it("uses a real covering fallback face explicitly when the imported font has no bold variant", async () => {
    const original = await embeddedRow("Hello Привіт");
    const { session, end } = await pdfSession(original, { fallbackFont: true });
    try {
      const originalNative = await nativeObjects(original);
      const baseline = (await session.getTextLayout("p0:o0")).item?.lines[0]
        ?.baseline;
      const receipt = await session.setTextStyle({
        target: "p0:o0",
        style: { bold: true, italic: true },
      });
      assert.equal(receipt.warnings[0]?.code, "font-substitution");
      const element = (await session.getElement("p0:o0")).item;
      assert.equal(element?.text, "Hello Привіт");
      assert.equal(
        (await nativeObjects((await session.save()).bytes))[0]?.fill?.[3],
        originalNative[0]?.fill?.[3],
      );
      assert.equal(element?.textStyle?.fontFamily, "Liberation Sans");
      assert.equal(element?.textStyle?.bold, true);
      assert.equal(element?.textStyle?.italic, true);
      assert.deepEqual(
        (await session.getTextLayout("p0:o0")).item?.lines[0]?.baseline,
        baseline,
      );
      const reopened = await pdfSession((await session.save()).bytes, {
        fallbackFont: true,
      });
      try {
        assert.deepEqual(
          (await reopened.session.getElement("p0:o0")).item?.textStyle,
          element?.textStyle,
        );
        assert.equal(await raster(reopened.session), await raster(session));
        await reopened.session.setTextStyle({
          target: "p0:o0",
          style: { bold: false, italic: false },
        });
        const regular = (await reopened.session.getElement("p0:o0")).item;
        assert.equal(regular?.textStyle?.bold, false);
        assert.equal(regular?.textStyle?.italic, false);
      } finally {
        await reopened.end();
      }
      await session.undo();
      assert.deepEqual((await session.save()).bytes, original);
    } finally {
      await end();
    }
  });

  it("retains an embedded font and glyph layout for underline-only changes", async () => {
    const { session, end } = await pdfSession(
      await embeddedRow("Exact Привіт"),
    );
    try {
      const font = (await session.getTextFont("p0:o0")).item;
      assert.ok(font);
      const layout = (await session.getTextLayout("p0:o0")).item;
      await underline(session, "p0:o0");
      assert.deepEqual((await session.getTextFont("p0:o0")).item, font);
      assert.deepEqual((await session.getTextLayout("p0:o0")).item, layout);
      const reopened = await pdfSession((await session.save()).bytes);
      try {
        assert.deepEqual(
          (await reopened.session.getTextFont("p0:o0")).item,
          font,
        );
      } finally {
        await reopened.end();
      }
    } finally {
      await end();
    }
  });

  it("rejects an unavailable style font and preserves bytes and history", async () => {
    const original = await embeddedRow("Hello Привіт");
    const { session, end } = await pdfSession(original);
    try {
      const before = await raster(session);
      await assert.rejects(
        session.setTextStyle({ target: "p0:o0", style: { bold: true } }),
      );
      assert.deepEqual((await session.save()).bytes, original);
      assert.equal(await raster(session), before);
      assert.equal(session.state.canUndo, false);
    } finally {
      await end();
    }
  });

  it("refuses stale underline ownership after an external PDF path edit", async () => {
    const { session, end } = await pdfSession(await buildPdf(["Owned text"]));
    let bytes: Uint8Array;
    try {
      await underline(session, "p0:o0");
      bytes = (await session.save()).bytes;
    } finally {
      await end();
    }
    const pdfium = await fixturePdfium();
    for (const mutation of ["move", "reorder"] as const) {
      const document = pdfium.openDocument(bytes);
      const page = pdfium.lib.FPDF_LoadPage(document.handle, 0);
      let altered: Uint8Array;
      try {
        const path = pdfium.lib.FPDFPage_GetObject(page, 1);
        assert.equal(pdfium.lib.FPDFPageObj_GetType(path), 2);
        if (mutation === "move")
          pdfium.lib.FPDFPageObj_Transform(path, 1, 0, 0, 1, 20, 0);
        else {
          assert.ok(pdfium.lib.FPDFPage_RemoveObject(page, path));
          pdfium.lib.FPDFPage_InsertObjectAtIndex(page, path, 0);
        }
        pdfium.lib.FPDFPage_GenerateContent(page);
        altered = document.save("full");
      } finally {
        pdfium.lib.FPDF_ClosePage(page);
        document.close();
      }
      const reopened = await pdfSession(altered);
      try {
        const rows = (await reopened.session.getElements({ pageIndex: 0 }))
          .items;
        assert.equal(
          rows.length,
          2,
          "an externally changed path is independent",
        );
        assert.equal(
          rows.find((row) => row.kind === "text")?.textStyle?.underline,
          undefined,
        );
        assert.ok(rows.some((row) => row.kind === "shape"));
      } finally {
        await reopened.end();
      }
    }
  });

  it("keeps underline with every native fragment after a ranged fallback replacement", async () => {
    const { session, end } = await pdfSession(await buildPdf(["Hello world"]), {
      fallbackFont: true,
    });
    try {
      await underline(session, "p0:o0");
      const previous = (await session.save()).bytes;
      await session.replaceText({
        target: "p0:o0",
        text: "Привіт",
        range: {
          start: { elementId: "p0:o0", offset: 1 },
          end: { elementId: "p0:o0", offset: 4 },
        },
      });
      const rows = (await session.getElements({ pageIndex: 0 })).items;
      assert.equal(rows.length, 3);
      assert.equal(
        rows.map((row) => row.text?.trimEnd()).join(""),
        "HПривітo world",
      );
      assert.ok(
        rows.every((row) => row.kind === "text" && row.textStyle?.underline),
      );
      const saved = (await session.save()).bytes;
      assert.equal(
        (await nativeObjects(saved)).filter((object) => object.type === 2)
          .length,
        3,
      );
      const reopened = await pdfSession(saved);
      try {
        assert.equal(
          (await reopened.session.getElements({ pageIndex: 0 })).items.length,
          3,
        );
        assert.equal(await raster(reopened.session), await raster(session));
      } finally {
        await reopened.end();
      }
      await session.undo();
      assert.deepEqual((await session.save()).bytes, previous);
    } finally {
      await end();
    }
  });

  it("rebuilds native textbox underlines after typing and removes only owned paths when toggled off", async () => {
    const { session, end } = await pdfSession(
      await buildPdf([{ text: "Neighbor" }]),
    );
    try {
      const receipt = await session.insertTextBox({
        pageIndex: 0,
        rect: { x: 72, y: 150, width: 180, height: 100 },
        text: "First line\nSecond line",
        style: { underline: true },
      });
      const [id] = receipt.createdIds;
      assert.ok(id);
      assert.equal(
        (await session.getElement(id)).item?.textStyle?.underline,
        true,
      );
      await session.replaceText({ target: id, text: "One line" });
      assert.equal(
        (await nativeObjects((await session.save()).bytes)).filter(
          (object) => object.type === 2,
        ).length,
        1,
      );
      await session.setTextStyle({ target: id, style: { underline: false } });
      assert.equal(
        (await nativeObjects((await session.save()).bytes)).filter(
          (object) => object.type === 2,
        ).length,
        0,
      );
      const reopened = await pdfSession((await session.save()).bytes);
      try {
        assert.equal(
          (await reopened.session.getElement(id)).item?.text,
          "One line",
        );
      } finally {
        await reopened.end();
      }
    } finally {
      await end();
    }
  });
  it("uses real bold and italic faces for an imported row through save and history", async () => {
    const original = await buildPdf([
      { texts: [{ text: "Native title", x: 72, y: 680, fontSize: 20 }] },
    ]);
    const { session, end } = await pdfSession(original);
    try {
      const before = (await session.getTextLayout("p0:o0")).item;
      assert.ok(before);
      const originalPixels = await raster(session);
      const receipt = await session.setTextStyle({
        target: "p0:o0",
        style: { bold: true, italic: true },
      });
      assert.deepEqual(receipt.warnings, []);
      const styled = (await session.getElement("p0:o0")).item;
      assert.ok(styled);
      assert.equal(styled.text, "Native title");
      assert.equal(styled.textStyle?.bold, true);
      assert.equal(styled.textStyle?.italic, true);
      assert.equal(styled.textStyle?.fontFamily, "Helvetica");
      const layout = (await session.getTextLayout(styled.id)).item;
      assert.deepEqual(layout?.lines[0]?.baseline, before.lines[0]?.baseline);
      const painted = await raster(session);
      assert.notEqual(painted, originalPixels);
      const saved = (await session.save()).bytes;
      const reopened = await pdfSession(saved);
      try {
        assert.deepEqual(
          (await reopened.session.getElement(styled.id)).item?.textStyle,
          styled.textStyle,
        );
        assert.equal(await raster(reopened.session), painted);
      } finally {
        await reopened.end();
      }
      await session.undo();
      assert.deepEqual((await session.save()).bytes, original);
      assert.equal(await raster(session), originalPixels);
      await session.redo();
      assert.equal(await raster(session), painted);
      assert.deepEqual((await session.getTextLayout(styled.id)).item, layout);
    } finally {
      await end();
    }
  });

  it("reflows an imported paragraph in a real bold face as one native history entry", async () => {
    const original = await buildPdf([
      {
        texts: paragraphLines.map((text, index) => ({
          text,
          x: 90,
          y: 700 - index * 20,
          fontSize: 11,
        })),
      },
    ]);
    const { session, end } = await pdfSession(original);
    try {
      const paragraph = (await session.getTextParagraph("p0:o0")).item;
      assert.ok(paragraph);
      const before = await raster(session);
      await session.setTextStyle({
        target: paragraph.id,
        style: { bold: true },
      });
      const styled = (await session.getTextParagraph(paragraph.id)).item;
      assert.ok(styled);
      assert.equal(styled.text, paragraph.text);
      assert.equal(styled.textStyle.bold, true);
      assert.equal(styled.bounds.width, paragraph.bounds.width);
      const painted = await raster(session);
      assert.notEqual(painted, before);
      const reopened = await pdfSession((await session.save()).bytes);
      try {
        assert.equal(
          (await reopened.session.getTextParagraph(paragraph.id)).item
            ?.textStyle.bold,
          true,
        );
        assert.equal(await raster(reopened.session), painted);
      } finally {
        await reopened.end();
      }
      await session.undo();
      assert.deepEqual((await session.save()).bytes, original);
      await session.redo();
      assert.equal(await raster(session), painted);
    } finally {
      await end();
    }
  });

  it("exports underline without rewriting native glyphs and owns it through hide, move and delete", async () => {
    const original = await buildPdf([
      {
        texts: [
          {
            text: "Underlined native title",
            x: 72,
            y: 680,
            fontSize: 1,
            matrix: [20, 0, 0, 20],
          },
          { text: "Neighbor stays", x: 72, y: 600, fontSize: 12 },
        ],
      },
    ]);
    const { session, end } = await pdfSession(original);
    try {
      const beforeLayout = (await session.getTextLayout("p0:o0")).item;
      const originalPixels = await raster(session);
      const hiddenOriginal = await raster(session, ["p0:o0"]);
      const beforeNative = await nativeObjects(original);
      await underline(session, "p0:o0");
      const underlined = (await session.getElement("p0:o0")).item;
      assert.ok(underlined?.textStyle && "underline" in underlined.textStyle);
      assert.equal(underlined.textStyle.underline, true);
      assert.deepEqual(
        (await session.getTextLayout("p0:o0")).item,
        beforeLayout,
      );
      const painted = await raster(session);
      assert.notEqual(painted, originalPixels);
      assert.equal(await raster(session, ["p0:o0"]), hiddenOriginal);
      const saved = (await session.save()).bytes;
      const objects = await nativeObjects(saved);
      assert.equal(
        objects.filter((item) => item.type === 2).length,
        1,
        "one real exported PDF path",
      );
      assert.deepEqual(
        objects.filter((item) => item.type === 1),
        beforeNative,
      );
      const reopened = await pdfSession(saved);
      try {
        assert.equal(await raster(reopened.session), painted);
        assert.equal(await raster(reopened.session, ["p0:o0"]), hiddenOriginal);
      } finally {
        await reopened.end();
      }
      await session.moveElement({ target: "p0:o0", by: { dx: 20, dy: 0 } });
      assert.equal(await raster(session, ["p0:o0"]), hiddenOriginal);
      assert.notEqual(await raster(session), painted);
      await session.deleteElement({ target: "p0:o0" });
      assert.equal(await raster(session), hiddenOriginal);
      await session.undo();
      await session.undo();
      assert.equal(await raster(session), painted);
      await session.undo();
      assert.equal(await raster(session), originalPixels);
      assert.deepEqual((await session.save()).bytes, original);
      await session.redo();
      assert.equal(await raster(session), painted);
    } finally {
      await end();
    }
  });

  it("underlines an imported paragraph without moving or replacing any native text row", async () => {
    const original = await buildPdf([
      {
        texts: paragraphLines.map((text, index) => ({
          text,
          x: 90,
          y: 700 - index * 20,
          fontSize: 11,
        })),
      },
    ]);
    const { session, end } = await pdfSession(original);
    try {
      const paragraph = (await session.getTextParagraph("p0:o0")).item;
      assert.ok(paragraph);
      const layout = (await session.getTextLayout(paragraph.id)).item;
      const before = await nativeObjects(original);
      await underline(session, paragraph.id);
      const next = (await session.getTextLayout(paragraph.id)).item;
      assert.deepEqual(next, layout);
      const saved = (await session.save()).bytes;
      const objects = await nativeObjects(saved);
      assert.deepEqual(
        objects.filter((object) => object.type === 1),
        before,
      );
      assert.equal(objects.filter((object) => object.type === 2).length, 3);
      const reopened = await pdfSession(saved);
      try {
        assert.equal(
          (await reopened.session.getTextParagraph(paragraph.id)).item?.text,
          paragraph.text,
        );
        assert.equal(await raster(reopened.session), await raster(session));
      } finally {
        await reopened.end();
      }
    } finally {
      await end();
    }
  });
});
