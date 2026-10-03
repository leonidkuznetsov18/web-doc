import assert from "node:assert/strict";
import { describe, it } from "node:test";

import { PdfEditDocument } from "../src/edit/pdf/engine/document.js";
import type { OperationIssue, PageRect, PdfOperation } from "../src/index.js";
import { extractPageText, fixturePdfium } from "./fixtures/pdf-builder.js";
import { pdfSession } from "./fixtures/pdf-session.js";
import {
  clippedTranslucentFormsPdf,
  deeplyNestedFormPdf,
  formGraphicsStatePdf,
  groupedFormPdf,
  nestedFormPdf,
  sharedFormPdf,
} from "./fixtures/pdf-forms.js";

/*
 * Text inside Form XObjects (ACTION-918). PDFium parses forms but writes
 * only page content, so an edit rewrites the forms on the way to the text,
 * for that one drawing of them; the rest of the page must look the same.
 */

const op = <T extends PdfOperation>(operation: T): T => operation;

async function open(bytes: Uint8Array): Promise<PdfEditDocument> {
  return new PdfEditDocument(await fixturePdfium(), bytes);
}

const codes = (issues: readonly OperationIssue[]) =>
  issues.map((issue) => `${issue.path}:${issue.code}`);

/** Validation as the worker runs it: form rewrites are checked first. */
async function validate(
  model: PdfEditDocument,
  operations: readonly PdfOperation[],
): Promise<OperationIssue[]> {
  await model.prepareRewrites(operations);
  return model.validate(operations);
}

async function apply(
  model: PdfEditDocument,
  operations: readonly PdfOperation[],
) {
  assert.deepEqual(await validate(model, operations), []);
  return model.apply(operations);
}

function centre(rect: PageRect): { x: number; y: number } {
  return { x: rect.x + rect.width / 2, y: rect.y + rect.height / 2 };
}

/** Pixels whose colour moved by more than antialiasing does. */
function changedPixels(before: ArrayBuffer, after: ArrayBuffer): number {
  const a = new Uint8Array(before);
  const b = new Uint8Array(after);
  assert.equal(a.length, b.length);
  let changed = 0;
  for (let offset = 0; offset < a.length; offset += 4)
    if (
      [0, 1, 2].some(
        (channel) => Math.abs(a[offset + channel]! - b[offset + channel]!) > 64,
      )
    )
      changed += 1;
  return changed;
}

describe("PDF text inside Form XObjects", () => {
  it("lists text in nested forms as text elements named by their path", async () => {
    const model = await open(nestedFormPdf());
    try {
      const elements = model.getElements({ pageIndex: 0 });
      assert.deepEqual(
        elements.map((element) => `${element.id} ${element.kind}`),
        ["p0:o0 text", "p0:o1 other", "p0:o1/0/0 text"],
      );
      const nested = elements[2]!;
      assert.equal(nested.text, "Nested Form text");
      assert.equal(nested.textStyle?.fontSize, 24);
      // Drawn at (72, 650) on a 792pt page, 10pt above the form's origin.
      assert.ok(Math.abs(nested.bounds.x - 73.8) < 1, `${nested.bounds.x}`);
      assert.ok(
        Math.abs(nested.bounds.y + nested.bounds.height - (792 - 660)) < 3,
        JSON.stringify(nested.bounds),
      );
      assert.ok(nested.operations.includes("replaceText"));
      assert.ok(!nested.operations.includes("resizeElement"));
    } finally {
      model.dispose();
    }
  });

  it("sizes and places text by every form it is drawn through", async () => {
    const model = await open(sharedFormPdf());
    try {
      const text = model
        .getElements({ pageIndex: 0, kinds: ["text"] })
        .map((element) => ({
          id: element.id,
          text: element.text?.trim(),
          size: element.textStyle?.fontSize,
          x: element.bounds.x,
        }));
      // Outer's /Matrix doubles everything inside it.
      assert.deepEqual(
        text.map(({ id, text, size }) => ({ id, text, size })),
        [
          { id: "p0:o0", text: "Top-level control text", size: 18 },
          { id: "p0:o1/1", text: "Outer own text", size: 20 },
          { id: "p0:o1/2/1", text: "Shared inner", size: 16 },
          { id: "p0:o1/3/1", text: "Shared inner", size: 16 },
        ],
      );
      // The second drawing of Inner is 120 form units, 240pt, to the right.
      assert.ok(Math.abs(text[3]!.x - text[2]!.x - 240) < 0.01);
    } finally {
      model.dispose();
    }
  });

  it("finds, hits and places the caret in nested text", async () => {
    const model = await open(nestedFormPdf());
    try {
      const [found] = model.findText("Nested Form text", {});
      assert.deepEqual(found?.elementIds, ["p0:o1/0/0"]);
      assert.deepEqual(found?.ranges, [
        {
          start: { elementId: "p0:o1/0/0", offset: 0 },
          end: { elementId: "p0:o1/0/0", offset: 16 },
        },
      ]);
      const layout = model.textLayout("p0:o1/0/0")!;
      assert.equal(layout.lines[0]?.text, "Nested Form text");
      assert.equal(layout.lines[0]?.fontSize, 24);
      // The ninth glyph, "o" of "Form": hits and the caret land on the
      // nested text, not on the page text next to it (the report's case).
      const glyph = layout.lines[0]!.glyphs[8]!;
      assert.equal(model.elementsAt(0, centre(glyph.box))[0]?.id, "p0:o1/0/0");
      assert.equal(
        model.positionAt(0, centre(glyph.box))?.elementId,
        "p0:o1/0/0",
      );
    } finally {
      model.dispose();
    }
  });

  it("edits nested text and keeps its id through save, reopen and replay", async () => {
    const model = await open(nestedFormPdf());
    const edit = op({
      op: "replaceText",
      target: "p0:o1/0/0",
      range: {
        start: { elementId: "p0:o1/0/0", offset: 7 },
        end: { elementId: "p0:o1/0/0", offset: 11 },
      },
      text: "XObject",
    });
    try {
      assert.deepEqual(await validate(model, [edit]), []);
      const change = await apply(model, [edit]);
      assert.deepEqual(change.createdIds, []);
      assert.equal(model.getElement("p0:o1/0/0")?.text, "Nested XObject text");
      const saved = model.materialize("save");
      assert.equal(
        await extractPageText(saved, 0),
        "Top-level control text\r\nNested XObject text",
      );
      const reopened = await open(saved);
      try {
        assert.equal(
          reopened.getElement("p0:o1/0/0")?.text,
          "Nested XObject text",
        );
      } finally {
        reopened.dispose();
      }
      model.restore([]);
      assert.equal(model.getElement("p0:o1/0/0")?.text, "Nested Form text");
      await model.restorePrepared({
        batches: [{ stateId: 1, operations: [edit] }],
      });
      assert.equal(model.getElement("p0:o1/0/0")?.text, "Nested XObject text");
    } finally {
      model.dispose();
    }
  });

  it("rewrites only the drawing of a shared form that it edits", async () => {
    const model = await open(sharedFormPdf());
    try {
      const before = model.renderPageWithout(0, ["p0:o1/2/1"], 1).data;
      await apply(model, [
        op({ op: "replaceText", target: "p0:o1/2/1", text: "Inner one" }),
      ]);
      assert.equal(model.getElement("p0:o1/2/1")?.text?.trim(), "Inner one");
      assert.equal(model.getElement("p0:o1/3/1")?.text, "Shared inner");
      // The bars the forms' boxes cut off stay cut off; nothing else moves.
      const after = model.renderPageWithout(0, ["p0:o1/2/1"], 1).data;
      assert.equal(changedPixels(before, after), 0);
      assert.match(
        await extractPageText(model.materialize("save"), 0),
        /Inner one\s+Shared inner/,
      );
    } finally {
      model.dispose();
    }
  });

  it("styles, moves and deletes nested text in page space", async () => {
    const model = await open(sharedFormPdf());
    try {
      const start = model.getElement("p0:o1/1")!.bounds;
      await apply(model, [
        op({
          op: "setTextStyle",
          target: "p0:o1/1",
          style: { fontSize: 30, color: "#ff0000" },
        }),
      ]);
      const styled = model.getElement("p0:o1/1")!;
      assert.equal(styled.textStyle?.fontSize, 30);
      assert.equal(styled.textStyle?.color, "#ff0000");
      assert.ok(Math.abs(styled.bounds.x - start.x) < 1);
      assert.ok(Math.abs(styled.bounds.y - start.y) < 1);

      await apply(model, [
        op({ op: "moveElement", target: "p0:o1/1", by: { dx: 10, dy: 20 } }),
      ]);
      const moved = model.getElement("p0:o1/1")!.bounds;
      assert.ok(Math.abs(moved.x - styled.bounds.x - 10) < 0.01);
      assert.ok(Math.abs(moved.y - styled.bounds.y - 20) < 0.01);

      const change = await apply(model, [
        op({ op: "deleteElement", target: "p0:o1/1" }),
      ]);
      assert.deepEqual(change.removedIds, ["p0:o1/1"]);
      assert.equal(model.getElement("p0:o1/1"), undefined);
      assert.doesNotMatch(
        await extractPageText(model.materialize("save"), 0),
        /Outer own text/,
      );
    } finally {
      model.dispose();
    }
  });

  it("keeps nested targets after sibling deletion and checkpoint Undo/Redo", async () => {
    const { session, end } = await pdfSession(sharedFormPdf());
    const first = "p0:o1/2/1";
    const second = "p0:o1/3/1";
    try {
      await session.apply([{ op: "deleteElement", target: "p0:o1/1" }]);
      await session.createCheckpoint("after sibling deletion");
      const original = (await session.getElements({ pageIndex: 0 })).items;
      const firstX = original.find((element) => element.id === first)!.bounds.x;
      const secondX = original.find((element) => element.id === second)!.bounds
        .x;
      assert.ok(Math.abs(firstX - 72.784) < 0.01);
      assert.ok(Math.abs(secondX - 312.784) < 0.01);

      await session.setTextStyle({
        target: first,
        style: { color: "#ff0000" },
      });
      await session.undo();
      await session.redo();
      const edited = (await session.getElements({ pageIndex: 0 })).items;
      assert.equal(
        edited.find((element) => element.bounds.x === firstX)?.textStyle?.color,
        "#ff0000",
      );
      assert.equal(
        edited.find((element) => element.bounds.x === secondX)?.textStyle
          ?.color,
        "#000000",
      );
      assert.equal(
        edited.find((element) => element.id === first)?.bounds.x,
        firstX,
      );
      assert.equal(
        edited.find((element) => element.id === second)?.bounds.x,
        secondX,
      );

      const saved = await session.save();
      const reopened = await open(saved.bytes);
      try {
        assert.equal(reopened.getElement(first)?.bounds.x, firstX);
        assert.equal(reopened.getElement(first)?.textStyle?.color, "#ff0000");
        assert.equal(reopened.getElement(second)?.bounds.x, secondX);
        assert.equal(reopened.getElement(second)?.textStyle?.color, "#000000");
      } finally {
        reopened.dispose();
      }
    } finally {
      await end();
    }
  });

  it("keeps fallback-created form children through checkpoint restoration", async () => {
    const { session, end } = await pdfSession(nestedFormPdf(), {
      fallbackFont: true,
    });
    const target = "p0:o1/0/0";
    try {
      const receipt = await session.replaceText({
        target,
        range: {
          start: { elementId: target, offset: 7 },
          end: { elementId: target, offset: 11 },
        },
        text: "Ж",
      });
      assert.equal(receipt.createdIds.length, 2);
      const inserted = receipt.createdIds[0]!;
      await session.createCheckpoint("after native run split");
      await session.setTextStyle({
        target: inserted,
        style: { color: "#ff0000" },
      });
      await session.undo();
      await session.redo();
      const elements = (
        await session.getElements({ pageIndex: 0, kinds: ["text"] })
      ).items;
      assert.equal(
        elements.find((element) => element.id === inserted)?.text?.trim(),
        "Ж",
      );
      assert.equal(
        elements.find((element) => element.id === inserted)?.textStyle?.color,
        "#ff0000",
      );
      const reopened = await open((await session.save()).bytes);
      try {
        assert.equal(reopened.getElement(inserted)?.text?.trim(), "Ж");
        assert.equal(
          reopened.getElement(receipt.createdIds[1]!)?.text?.trim(),
          "text",
        );
      } finally {
        reopened.dispose();
      }
    } finally {
      await end();
    }
  });

  it("ignores damaged, stale, oversized and colliding form identities", async () => {
    const inner = (id: string) => ({
      id,
      type: 5,
      children: [
        { id: `${id}/0`, type: 2 },
        { id: `${id}/1`, type: 1 },
      ],
    });
    const tree = {
      id: "p0:o1",
      type: 5,
      children: [
        { id: "p0:o1/0", type: 2 },
        { id: "p0:o1/1", type: 1 },
        inner("p0:o1/2"),
        inner("p0:o1/3"),
      ],
    };
    const invalid = [
      "{",
      JSON.stringify({ ...tree, children: tree.children.slice(1) }),
      JSON.stringify({
        ...tree,
        children: [{ id: "p0:o1/0", type: 1 }, ...tree.children.slice(1)],
      }),
      JSON.stringify({
        ...tree,
        children: [{ id: "p0:o1/1", type: 2 }, ...tree.children.slice(1)],
      }),
      JSON.stringify({
        ...tree,
        children: [{ id: "p0:o0", type: 2 }, ...tree.children.slice(1)],
      }),
      JSON.stringify({
        ...tree,
        children: [{ id: "p9:o0", type: 2 }, ...tree.children.slice(1)],
      }),
      JSON.stringify("x".repeat(600_000)),
    ];
    const pdfium = await fixturePdfium();
    const cases = [
      ...invalid.map((raw) => ({ raw, stale: false })),
      { raw: invalid[4]!, stale: true },
    ];
    for (const { raw, stale } of cases) {
      const document = pdfium.openDocument(sharedFormPdf());
      const page = pdfium.lib.FPDF_LoadPage(document.handle, 0);
      let bytes: Uint8Array;
      try {
        if (stale) {
          const control = pdfium.lib.FPDFPage_GetObject(page, 0);
          const staleMark = pdfium.lib.FPDFPageObj_AddMark(control, "WebDoc");
          assert.ok(
            pdfium.setMarkString(
              document.handle,
              control,
              staleMark,
              "webdoc",
              JSON.stringify({
                kind: "textBox",
                id: "p0:n7.0.0",
                rect: { x: 72, y: 200, width: 300, height: 40 },
                text: "Template",
                style: {
                  fontFamily: "Helvetica",
                  fontSize: 12,
                  bold: false,
                  italic: false,
                  color: "#000000",
                  align: "left",
                  lineHeight: 1.2,
                },
              }),
            ),
          );
        }
        const object = pdfium.lib.FPDFPage_GetObject(page, 1);
        const mark = pdfium.lib.FPDFPageObj_AddMark(object, "WebDocFormIds");
        assert.ok(
          pdfium.setMarkString(document.handle, object, mark, "ids", raw),
        );
        assert.ok(pdfium.lib.FPDFPage_GenerateContent(page));
        bytes = document.save("full");
      } finally {
        pdfium.lib.FPDF_ClosePage(page);
        document.close();
      }
      const model = await open(bytes);
      try {
        assert.equal(model.getElement("p0:o0")?.text, "Top-level control text");
        assert.equal(
          model.getElement("p0:o1/2/1")?.text?.trim(),
          "Shared inner",
        );
        assert.equal(
          model.getElement("p0:o1/3/1")?.text?.trim(),
          "Shared inner",
        );
      } finally {
        model.dispose();
      }
    }
  });

  it("retains a deeply nested form identity when an earlier page object is removed", async () => {
    const model = await open(deeplyNestedFormPdf());
    try {
      model.apply([{ op: "deleteElement", target: "p0:o0" }]);
      const reopened = await open(model.materialize("save"));
      try {
        assert.equal(reopened.getElement("p0:o1")?.kind, "other");
        assert.equal(reopened.getElement("p0:o0"), undefined);
      } finally {
        reopened.dispose();
      }
    } finally {
      model.dispose();
    }
  });

  it("binds saved form identities to the page a drawing moves onto", async () => {
    const { session, end } = await pdfSession(sharedFormPdf());
    try {
      await session.setTextStyle({
        target: "p0:o1/2/1",
        style: { color: "#ff0000" },
      });
      await session.apply([
        { op: "insertPage", index: 0, size: { width: 612, height: 792 } },
      ]);
      const reopened = await open((await session.save()).bytes);
      try {
        const target = "p1:p0:o1/2/1";
        assert.equal(reopened.getElement(target)?.pageIndex, 1);
        assert.equal(reopened.getElement(target)?.textStyle?.color, "#ff0000");
        await apply(reopened, [
          { op: "setTextStyle", target, style: { color: "#00ff00" } },
        ]);
        assert.equal(reopened.getElement(target)?.textStyle?.color, "#00ff00");
        assert.equal(
          reopened.getElement("p1:p0:o1/3/1")?.textStyle?.color,
          "#000000",
        );
      } finally {
        reopened.dispose();
      }
    } finally {
      await end();
    }
  });

  it("keeps nested bold and italic through save, replay and further typing", async () => {
    const { session, end } = await pdfSession(nestedFormPdf());
    try {
      const target = "p0:o1/0/0";
      await session.setTextStyle({
        target,
        style: { bold: true, italic: true },
      });
      const styled = (await session.getElement(target)).item;
      assert.equal(styled?.text, "Nested Form text");
      assert.equal(styled?.textStyle?.bold, true);
      assert.equal(styled?.textStyle?.italic, true);
      await session.undo();
      assert.equal(
        (await session.getElement(target)).item?.textStyle?.bold,
        false,
      );
      await session.redo();
      await session.replaceText({ target, text: "Nested revised text" });
      const saved = await session.save();
      const reopened = await open(saved.bytes);
      try {
        const text = reopened.getElement(target);
        assert.equal(text?.text, "Nested revised text");
        assert.equal(text?.textStyle?.bold, true);
        assert.equal(text?.textStyle?.italic, true);
        assert.equal(
          reopened.getElement("p0:o0")?.text,
          "Top-level control text",
        );
      } finally {
        reopened.dispose();
      }
    } finally {
      await end();
    }
  });

  it("refuses unsupported nested empty-frame and underline edits without changing bytes or history", async () => {
    const original = nestedFormPdf();
    const { session, end } = await pdfSession(original);
    try {
      for (const operations of [
        [op({ op: "replaceText", target: "p0:o1/0/0", text: "" })],
        [
          op({
            op: "setTextStyle",
            target: "p0:o1/0/0",
            style: { underline: true },
          }),
        ],
      ]) {
        await assert.rejects(
          session.apply(operations),
          (error: unknown) =>
            error instanceof Error &&
            "code" in error &&
            error.code === "invalid-operation",
        );
        assert.equal(session.state.canUndo, false);
        assert.equal(
          (await session.getElement("p0:o1/0/0")).item?.text,
          "Nested Form text",
        );
        assert.deepEqual((await session.save()).bytes, original);
      }
    } finally {
      await end();
    }
  });

  it("keeps the clip a form is drawn with", async () => {
    const model = await open(clippedTranslucentFormsPdf());
    try {
      const before = model.renderPageWithout(0, ["p0:o0/1"], 1).data;
      const edit = op({ op: "replaceText", target: "p0:o0/1", text: "Edited" });
      assert.deepEqual(await validate(model, [edit]), []);
      await apply(model, [edit]);
      // The red bar stays cut where the clip around the form's drawing cuts it.
      const after = model.renderPageWithout(0, ["p0:o0/1"], 1).data;
      assert.equal(changedPixels(before, after), 0);
    } finally {
      model.dispose();
    }
  });

  it("refuses text whose objects take a graphics state by name", async () => {
    // Opacity set where the form is drawn reaches its objects by name, as
    // does a state from the form's own resources; the new form holds
    // neither, so the names would dangle.
    for (const [bytes, target] of [
      [clippedTranslucentFormsPdf(), "p0:o1/1"],
      [formGraphicsStatePdf(), "p0:o0/0"],
    ] as const) {
      const model = await open(bytes);
      try {
        assert.ok(model.getElement(target)?.text);
        assert.deepEqual(
          codes(
            await validate(model, [
              op({ op: "replaceText", target, text: "Edited" }),
            ]),
          ),
          ["/target:unsupported-target"],
          target,
        );
      } finally {
        model.dispose();
      }
    }
  });

  it("refuses text whose form a rewrite would draw differently", async () => {
    const original = groupedFormPdf();
    const model = await open(original);
    try {
      // A transparency group at half opacity: rewritten without the group,
      // the overlap of its squares would come out darker.
      assert.deepEqual(
        codes(
          await validate(model, [
            op({ op: "replaceText", target: "p0:o0/2", text: "Changed" }),
          ]),
        ),
        ["/target:unsupported-target"],
      );
      assert.deepEqual(model.materialize("save"), original);
      assert.deepEqual(
        codes(
          await validate(model, [
            op({
              op: "resizeElement",
              target: "p0:o0/2",
              rect: { x: 72, y: 20, width: 200, height: 20 },
            }),
          ]),
        ),
        ["/target:unsupported-target"],
      );
    } finally {
      model.dispose();
    }
  });

  it("reports the text inside a deleted form as removed", async () => {
    const model = await open(sharedFormPdf());
    try {
      const change = await apply(model, [
        op({ op: "deleteElement", target: "p0:o1" }),
      ]);
      assert.deepEqual(change.removedIds, [
        "p0:o1",
        "p0:o1/1",
        "p0:o1/2/1",
        "p0:o1/3/1",
      ]);
      assert.deepEqual(
        model.getElements({ pageIndex: 0 }).map((element) => element.id),
        ["p0:o0"],
      );
    } finally {
      model.dispose();
    }
  });

  it("edits nested text through a session, its worker and its history", async () => {
    const { session, end } = await pdfSession(nestedFormPdf());
    try {
      const layout = await session.getTextLayout("p0:o1/0/0");
      const glyph = layout.item!.lines[0]!.glyphs[8]!;
      const position = await session.positionAt(0, centre(glyph.box));
      assert.equal(position.item?.elementId, "p0:o1/0/0");
      await session.replaceText({
        target: "p0:o1/0/0",
        text: "Nested Form revised",
      });
      const text = async () =>
        (await session.getElement("p0:o1/0/0")).item?.text;
      assert.equal(await text(), "Nested Form revised");
      await session.undo();
      assert.equal(await text(), "Nested Form text");
      await session.redo();
      assert.equal(await text(), "Nested Form revised");
    } finally {
      await end();
    }

    // The worker runs the rewrite check before validating.
    const grouped = await pdfSession(groupedFormPdf());
    try {
      await assert.rejects(
        grouped.session.replaceText({ target: "p0:o0/2", text: "Changed" }),
        (error: { code?: string; details?: { issues?: OperationIssue[] } }) =>
          error.code === "invalid-operation" &&
          error.details?.issues?.[0]?.code === "unsupported-target",
      );
    } finally {
      await grouped.end();
    }
  });
});
