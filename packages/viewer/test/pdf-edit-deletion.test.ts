import assert from "node:assert/strict";
import { readFileSync } from "node:fs";
import { describe, it } from "node:test";

import type { PdfEditSession } from "../src/index.js";
import {
  buildPdf,
  extractPageText,
  fixturePdfium,
} from "./fixtures/pdf-builder.js";
import { pdfSession } from "./fixtures/pdf-session.js";

async function textElement(session: PdfEditSession, id: string) {
  const element = (await session.getElement(id)).item;
  assert.ok(element, `editable target ${id} remains discoverable`);
  return element;
}

async function fixture(kind: "imported" | "textBox", text: string) {
  const opened = await pdfSession(
    await buildPdf([
      {
        texts: [
          { text: "Neighbor", x: 72, y: 700 },
          ...(kind === "imported"
            ? [{ text, x: 200, y: 400, matrix: [0, 1, -1, 0] as const }]
            : []),
        ],
      },
    ]),
  );
  const id =
    kind === "imported"
      ? "p0:o1"
      : (
          await opened.session.insertTextBox({
            pageIndex: 0,
            rect: { x: 100, y: 180, width: 240, height: 60 },
            text,
          })
        ).createdIds[0];
  assert.ok(id);
  return { ...opened, id };
}

function range(elementId: string, start: number, end: number) {
  return {
    start: { elementId, offset: start },
    end: { elementId, offset: end },
  };
}

describe("native PDF text deletion [ACTION-954]", () => {
  for (const kind of ["imported", "textBox"] as const) {
    it(`deletes a trailing or middle span in ${kind} text and preserves history and export`, async () => {
      const { session, end, id } = await fixture(kind, "SECOND LINEX");
      try {
        const original = await textElement(session, id);
        const receipt = await session.replaceText({
          target: id,
          text: "",
          range: range(id, 11, 12),
        });
        assert.deepEqual(receipt.createdIds, []);
        const shortened = await textElement(session, id);
        assert.equal(shortened.text, "SECOND LINE");
        assert.equal(shortened.rotation, original.rotation);
        assert.deepEqual(shortened.textStyle, original.textStyle);
        await session.undo();
        assert.equal((await textElement(session, id)).text, "SECOND LINEX");
        await session.redo();
        assert.equal((await textElement(session, id)).text, "SECOND LINE");

        await session.replaceText({
          target: id,
          text: "",
          range: range(id, 6, 7),
        });
        assert.equal((await textElement(session, id)).text, "SECONDLINE");
        const saved = await session.save();
        const extracted = await extractPageText(saved.bytes, 0);
        assert.ok(extracted.includes("Neighbor"));
        assert.ok(extracted.includes("SECONDLINE"));
        assert.ok(!extracted.includes("SECOND LINEX"));
        const reopened = await pdfSession(saved.bytes);
        try {
          assert.equal(
            (await textElement(reopened.session, id)).text,
            "SECONDLINE",
          );
          await reopened.session.replaceText({
            target: id,
            text: " again",
            range: range(id, 10, 10),
          });
          assert.equal(
            (await textElement(reopened.session, id)).text,
            "SECONDLINE again",
          );
        } finally {
          await reopened.end();
        }
      } finally {
        await end();
      }
    });

    it(`retains a cleared underlined ${kind} target for typing after save and reopen`, async () => {
      const { session, end, id } = await fixture(kind, "Clear this");
      try {
        await session.setTextStyle({ target: id, style: { underline: true } });
        const original = await textElement(session, id);
        const suppressed = (await session.renderPageWithout(0, [id])).item;
        assert.ok(suppressed);
        await session.replaceText({
          target: id,
          text: "",
          range: range(id, 0, 10),
        });
        const empty = await textElement(session, id);
        assert.equal(empty.text, "");
        assert.deepEqual(
          empty.bounds,
          kind === "imported"
            ? original.bounds
            : { x: 100, y: 180, width: 240, height: 60 },
          "cleared rows retain their native extent; text boxes retain their authored frame",
        );
        assert.deepEqual(empty.textStyle, original.textStyle);
        assert.equal(empty.rotation, original.rotation);
        assert.ok(empty.operations.includes("replaceText"));
        assert.ok(
          (await session.getElements({ pageIndex: 0 })).items.some(
            (element) => element.id === id,
          ),
        );
        assert.deepEqual(
          (await session.renderPageWithout(0, [])).item?.data,
          suppressed.data,
        );
        await session.undo();
        assert.equal((await textElement(session, id)).text, "Clear this");
        await session.redo();
        assert.equal((await textElement(session, id)).text, "");
        const saved = await session.save();
        assert.equal(
          (await extractPageText(saved.bytes, 0)).trim(),
          "Neighbor",
        );

        const reopened = await pdfSession(saved.bytes);
        try {
          assert.deepEqual(await textElement(reopened.session, id), empty);
          assert.deepEqual(
            (await reopened.session.renderPageWithout(0, [])).item?.data,
            suppressed.data,
          );
          await reopened.session.replaceText({
            target: id,
            text: "Typing resumes",
            range: range(id, 0, 0),
          });
          const resumed = await textElement(reopened.session, id);
          assert.equal(resumed.text, "Typing resumes");
          assert.deepEqual(resumed.textStyle, original.textStyle);
          assert.equal(resumed.rotation, original.rotation);
          const written = await reopened.session.save();
          assert.ok(
            (await extractPageText(written.bytes, 0)).includes(
              "Typing resumes",
            ),
          );
          await reopened.session.undo();
          assert.deepEqual(await textElement(reopened.session, id), empty);
        } finally {
          await reopened.end();
        }
      } finally {
        await end();
      }
    });
  }

  it("keeps an unstyled cleared row's transformed frame movable and its formatting editable", async () => {
    const { session, end } = await pdfSession(
      await buildPdf([
        {
          rotation: 1,
          texts: [
            {
              text: "Angled text",
              x: 160,
              y: 400,
              matrix: [0.8660254, 0.5, -0.5, 0.8660254],
            },
          ],
        },
      ]),
    );
    try {
      const original = await textElement(session, "p0:o0");
      await session.replaceText({ target: original.id, text: "" });
      assert.deepEqual(
        (await textElement(session, original.id)).bounds,
        original.bounds,
      );
      await assert.rejects(
        session.setTextStyle({
          target: original.id,
          style: { underline: true },
        }),
        /axis-aligned|quarter-turned/,
      );
      await session.setTextStyle({
        target: original.id,
        style: { color: "#c02040", bold: true, italic: true, fontSize: 18 },
      });
      await session.moveElement({ target: original.id, by: { dx: 7, dy: 11 } });
      const empty = await textElement(session, original.id);
      assert.deepEqual(empty.bounds, {
        ...original.bounds,
        x: original.bounds.x + 7,
        y: original.bounds.y + 11,
      });
      assert.equal(empty.rotation, original.rotation);
      assert.deepEqual(empty.textStyle, {
        ...original.textStyle,
        color: "#c02040",
        bold: true,
        italic: true,
        fontSize: 18,
      });
      const saved = await session.save();
      assert.equal(await extractPageText(saved.bytes, 0), "");
      const reopened = await pdfSession(saved.bytes);
      try {
        assert.deepEqual(
          await textElement(reopened.session, original.id),
          empty,
        );
        const receipt = await reopened.session.replaceText({
          target: original.id,
          text: "Again",
        });
        assert.deepEqual(
          receipt.warnings,
          [],
          "a standard matching face is not falsely reported as substituted",
        );
        const resumed = await textElement(reopened.session, original.id);
        assert.equal(resumed.text, "Again");
        assert.equal(resumed.rotation, original.rotation);
        assert.deepEqual(resumed.textStyle, empty.textStyle);
        await reopened.session.undo();
        assert.deepEqual(
          await textElement(reopened.session, original.id),
          empty,
        );
      } finally {
        await reopened.end();
      }
    } finally {
      await end();
    }
  });

  it("rejects out-of-range or foreign-target deletions without changing bytes or history", async () => {
    const bytes = await buildPdf(["Keep this"]);
    const { session, end } = await pdfSession(bytes);
    try {
      for (const invalid of [
        range("p0:o0", 4, 99),
        range("p0:o0", 5, 2),
        range("p0:o9", 0, 9),
      ]) {
        await assert.rejects(
          session.replaceText({ target: "p0:o0", text: "", range: invalid }),
          /range/i,
        );
        assert.deepEqual((await session.save()).bytes, bytes);
        assert.equal(session.state.canUndo, false);
        assert.equal((await textElement(session, "p0:o0")).text, "Keep this");
      }
    } finally {
      await end();
    }
  });

  it("clears an imported embedded font without fetching a font and reports substitution when typing resumes", async () => {
    const pdfium = await fixturePdfium();
    const document = pdfium.createDocument();
    const { lib } = pdfium;
    const bytes = new Uint8Array(
      readFileSync(
        new URL("../../fonts/noto-sans-latin-cyrillic.ttf", import.meta.url),
      ),
    );
    const data = pdfium.writeBytes(bytes);
    let original: Uint8Array;
    try {
      const page = lib.FPDFPage_New(document.handle, 0, 612, 792);
      const font = lib.FPDFText_LoadFont(
        document.handle,
        data,
        bytes.length,
        1,
        true,
      );
      const object = lib.FPDFPageObj_CreateTextObj(document.handle, font, 14);
      const text = pdfium.writeWideString("Україна");
      assert.ok(lib.FPDFText_SetText(object, text));
      pdfium.free(text);
      lib.FPDFPageObj_Transform(object, 1, 0, 0, 1, 72, 700);
      lib.FPDFPage_InsertObject(page, object);
      lib.FPDFPage_GenerateContent(page);
      lib.FPDF_ClosePage(page);
      original = document.save("full");
    } finally {
      document.close();
      pdfium.free(data);
    }
    const partial = await pdfSession(original, { fallbackFont: true });
    try {
      await partial.session.replaceText({
        target: "p0:o0",
        text: "",
        range: range("p0:o0", 3, 4),
      });
      assert.equal(
        (await extractPageText((await partial.session.save()).bytes, 0)).trim(),
        "Укрїна",
      );
      await partial.session.undo();
      assert.equal(
        (await extractPageText((await partial.session.save()).bytes, 0)).trim(),
        "Україна",
      );
    } finally {
      await partial.end();
    }
    const { session, end } = await pdfSession(original);
    try {
      await session.replaceText({ target: "p0:o0", text: "" });
      const saved = await session.save();
      assert.equal(await extractPageText(saved.bytes, 0), "");
      const reopened = await pdfSession(saved.bytes, { fallbackFont: true });
      try {
        const receipt = await reopened.session.replaceText({
          target: "p0:o0",
          text: "Вітаю",
        });
        assert.ok(
          receipt.warnings.some(
            (warning) => warning.code === "font-substitution",
          ),
        );
        assert.equal(
          (await textElement(reopened.session, "p0:o0")).text,
          "Вітаю",
        );
        assert.equal(
          (
            await extractPageText((await reopened.session.save()).bytes, 0)
          ).trim(),
          "Вітаю",
        );
      } finally {
        await reopened.end();
      }
    } finally {
      await end();
    }
  });

  it("does not trust an empty-row mark once an external editor changes its fill-only draw mode", async () => {
    const { session, end } = await pdfSession(await buildPdf(["Clear this"]));
    let saved: Uint8Array;
    try {
      await session.replaceText({ target: "p0:o0", text: "" });
      saved = (await session.save()).bytes;
    } finally {
      await end();
    }
    const pdfium = await fixturePdfium();
    const document = pdfium.openDocument(saved);
    let changed: Uint8Array;
    try {
      const page = pdfium.lib.FPDF_LoadPage(document.handle, 0);
      const object = pdfium.lib.FPDFPage_GetObject(page, 0);
      pdfium.lib.FPDFPath_SetDrawMode(object, 1, true);
      pdfium.lib.FPDFPage_GenerateContent(page);
      pdfium.lib.FPDF_ClosePage(page);
      changed = document.save("full");
    } finally {
      document.close();
    }
    const reopened = await pdfSession(changed);
    try {
      const elements = (await reopened.session.getElements({ pageIndex: 0 }))
        .items;
      assert.equal(elements.length, 1);
      assert.equal(elements[0]?.kind, "shape");
      assert.ok(!elements[0]?.operations.includes("replaceText"));
    } finally {
      await reopened.end();
    }
  });

  it("does not reopen a moved empty textbox with stale authored placement", async () => {
    const { session, end, id } = await fixture("textBox", "Clear this");
    let saved: Uint8Array;
    try {
      await session.replaceText({ target: id, text: "" });
      saved = (await session.save()).bytes;
    } finally {
      await end();
    }
    const pdfium = await fixturePdfium();
    const document = pdfium.openDocument(saved);
    let changed: Uint8Array;
    try {
      const page = pdfium.lib.FPDF_LoadPage(document.handle, 0);
      const object = pdfium.lib.FPDFPage_GetObject(page, 1);
      pdfium.lib.FPDFPageObj_Transform(object, 1, 0, 0, 1, 60, 0);
      pdfium.lib.FPDFPage_GenerateContent(page);
      pdfium.lib.FPDF_ClosePage(page);
      changed = document.save("full");
    } finally {
      document.close();
    }
    const reopened = await pdfSession(changed);
    try {
      const elements = (await reopened.session.getElements({ pageIndex: 0 }))
        .items;
      assert.equal(
        elements.some((element) => element.kind === "textBox"),
        false,
      );
      const anchor = elements.find((element) => element.kind === "shape");
      assert.equal(anchor?.bounds.x, 160);
      await assert.rejects(
        reopened.session.replaceText({
          target: id,
          text: "Must not jump back",
        }),
      );
    } finally {
      await reopened.end();
    }
  });
});
