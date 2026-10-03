import assert from "node:assert/strict";
import { describe, it } from "node:test";

import {
  ViewerError,
  type PdfEditSession,
  type PdfElement,
} from "../src/index.js";
import {
  buildPdf,
  extractPageText,
  fixturePdfium,
  type FixtureText,
} from "./fixtures/pdf-builder.js";
import { pdfSession } from "./fixtures/pdf-session.js";

const lines = [
  "The first paragraph line",
  "continues over a second line",
  "and finishes on a third line.",
] as const;
const paragraphText = lines.join(" ");
const paragraphObjects: readonly FixtureText[] = lines.map((text, index) => ({
  text,
  x: 72,
  y: 700 - index * 20,
  fontSize: 11,
}));
const rightColumn: readonly FixtureText[] = [
  { text: "Separate right column", x: 350, y: 700, fontSize: 11 },
  { text: "Its own second line", x: 350, y: 680, fontSize: 11 },
];
const nextParagraph: FixtureText = {
  text: "The next paragraph stays separate.",
  x: 72,
  y: 630,
  fontSize: 11,
};

function textElement(
  elements: readonly PdfElement[],
  text: string,
): PdfElement {
  const element = elements.find((item) => item.text?.trim() === text);
  assert.ok(element, `native text exists: ${text}`);
  return element;
}

async function paragraphFor(session: PdfEditSession, target: string) {
  const result = await session.getTextParagraph(target);
  assert.ok(result.item, `paragraph is exposed for ${target}`);
  return result.item;
}

function contentOf(element: PdfElement) {
  return {
    text: element.text,
    bounds: element.bounds,
    textStyle: element.textStyle,
    kind: element.kind,
  };
}

describe("native PDF paragraph boundaries", () => {
  for (const order of ["forward", "interleaved reverse"] as const) {
    it(`uses visual reading order without absorbing another column (${order})`, async () => {
      const objects =
        order === "forward"
          ? [...paragraphObjects, ...rightColumn]
          : [...paragraphObjects, ...rightColumn].reverse();
      // Native drawing order does not have to match the page's reading order.
      if (order === "interleaved reverse") {
        const column = objects.shift();
        assert.ok(column);
        objects.splice(3, 0, column);
      }
      const { session, end } = await pdfSession(
        await buildPdf([{ texts: objects }]),
      );
      try {
        const elements = (await session.getElements({ pageIndex: 0 })).items;
        const members = lines.map((line) => textElement(elements, line));
        const first = members[0];
        assert.ok(first);
        const paragraph = await paragraphFor(session, first.id);
        assert.equal(paragraph.text, paragraphText);
        assert.deepEqual(
          paragraph.memberIds,
          members.map((member) => member.id),
        );
        for (const member of members) {
          assert.equal(member.textEditingTarget, paragraph.id);
          assert.deepEqual(await paragraphFor(session, member.id), paragraph);
        }
        assert.deepEqual(await paragraphFor(session, paragraph.id), paragraph);
        const canonical = (await session.getElement(paragraph.id)).item;
        assert.ok(canonical);
        assert.equal(canonical.kind, "paragraph");
        assert.equal(canonical.text, paragraphText);
        assert.deepEqual(canonical.bounds, paragraph.bounds);
        assert.ok(canonical.operations.includes("deleteElement"));
        const full = (await session.renderPageWithout(0, [])).item;
        const withoutCanonical = (
          await session.renderPageWithout(0, [paragraph.id])
        ).item;
        const withoutMembers = (
          await session.renderPageWithout(0, paragraph.memberIds)
        ).item;
        assert.ok(full && withoutCanonical && withoutMembers);
        assert.deepEqual(withoutCanonical.data, withoutMembers.data);
        assert.ok(
          full.data.some(
            (byte, index) => byte !== withoutCanonical.data[index],
          ),
          "canonical suppression removes native paragraph pixels",
        );
        assert.equal(session.state.dirty, false);
        for (const column of rightColumn)
          assert.ok(
            !paragraph.memberIds.includes(
              textElement(elements, column.text).id,
            ),
          );
        assert.ok(paragraph.bounds.x + paragraph.bounds.width < 350);
      } finally {
        await end();
      }
    });
  }

  it("keeps a same-column paragraph separate after a larger vertical gap", async () => {
    const { session, end } = await pdfSession(
      await buildPdf([{ texts: [...paragraphObjects, nextParagraph] }]),
    );
    try {
      const elements = (await session.getElements({ pageIndex: 0 })).items;
      const lastLine = textElement(elements, lines[2]);
      const paragraph = await paragraphFor(session, lastLine.id);
      assert.equal(paragraph.text, paragraphText);
      assert.equal(paragraph.memberIds.length, 3);
      const neighbor = textElement(elements, nextParagraph.text);
      assert.ok(!paragraph.memberIds.includes(neighbor.id));
      assert.ok(
        paragraph.bounds.y + paragraph.bounds.height < neighbor.bounds.y,
      );
      await session.replaceParagraphText({
        target: paragraph.id,
        text: "A short replacement.",
      });
      assert.deepEqual((await session.getElement(neighbor.id)).item, neighbor);
    } finally {
      await end();
    }
  });

  it("rejects overflow without changing content or history and accepts a later valid edit", async () => {
    const { session, end } = await pdfSession(
      await buildPdf([
        { texts: [...paragraphObjects, ...rightColumn, nextParagraph] },
      ]),
    );
    try {
      const elements = (await session.getElements({ pageIndex: 0 })).items;
      const paragraph = await paragraphFor(
        session,
        textElement(elements, lines[1]).id,
      );
      const state = session.state;
      const saved = await session.save();
      const extracted = await extractPageText(saved.bytes, 0);

      await assert.rejects(
        session.replaceParagraphText({
          target: paragraph.id,
          text: "This replacement cannot fit inside the original paragraph. ".repeat(
            50,
          ),
        }),
        (error: unknown) => {
          assert.ok(error instanceof ViewerError);
          assert.equal(error.code, "invalid-operation");
          const issues = error.details?.issues;
          assert.ok(Array.isArray(issues));
          assert.ok(
            issues.some(
              (issue: unknown) =>
                typeof issue === "object" &&
                issue !== null &&
                "code" in issue &&
                issue.code === "paragraph-overflow",
            ),
          );
          return true;
        },
      );
      assert.deepEqual(session.state, state);
      assert.deepEqual(
        (await session.getElements({ pageIndex: 0 })).items,
        elements,
      );
      assert.deepEqual(await paragraphFor(session, paragraph.id), paragraph);
      const after = await session.save();
      assert.deepEqual(after.bytes, saved.bytes);
      assert.equal(await extractPageText(after.bytes, 0), extracted);

      await session.replaceParagraphText({
        target: paragraph.id,
        text: "A short replacement.",
      });
      assert.equal(
        (await paragraphFor(session, paragraph.id)).text,
        "A short replacement.",
      );
      assert.equal(session.state.revision, state.revision + 1);
      await session.undo();
      assert.equal(
        session.state.canUndo,
        false,
        "rejected edit adds no Undo entry",
      );
      assert.equal(
        (await paragraphFor(session, paragraph.id)).text,
        paragraphText,
      );
    } finally {
      await end();
    }
  });

  it("saves a complete editable replacement and keeps neighboring text unchanged after reopening", async () => {
    const original = await buildPdf([
      { texts: [...paragraphObjects, ...rightColumn, nextParagraph] },
    ]);
    const { session, end } = await pdfSession(original);
    try {
      const before = (await session.getElements({ pageIndex: 0 })).items;
      const paragraph = await paragraphFor(
        session,
        textElement(before, lines[2]).id,
      );
      const neighbors = [...rightColumn, nextParagraph].map((source) =>
        textElement(before, source.text),
      );
      const replacement = "The paragraph now has updated text.";
      await session.replaceParagraphText({
        target: paragraph.id,
        text: replacement,
      });
      for (const neighbor of neighbors)
        assert.deepEqual(
          (await session.getElement(neighbor.id)).item,
          neighbor,
        );
      const saved = await session.save();
      const text = (await extractPageText(saved.bytes, 0)).replaceAll(
        /\s+/g,
        " ",
      );
      assert.equal(
        text.split(replacement).length - 1,
        1,
        "replacement is native text exactly once",
      );
      for (const oldLine of lines)
        assert.ok(!text.includes(oldLine), "old text is removed, not hidden");

      const reopened = await pdfSession(saved.bytes);
      try {
        const current = await paragraphFor(reopened.session, paragraph.id);
        assert.equal(current.text, replacement);
        assert.equal(current.id, paragraph.id);
        for (const memberId of current.memberIds)
          assert.deepEqual(
            await paragraphFor(reopened.session, memberId),
            current,
          );
        const elements = (await reopened.session.getElements({ pageIndex: 0 }))
          .items;
        for (const neighbor of neighbors) {
          assert.ok(neighbor.text);
          assert.deepEqual(
            contentOf(textElement(elements, neighbor.text.trim())),
            contentOf(neighbor),
          );
        }
        await reopened.session.replaceParagraphText({
          target: current.id,
          text: "Edited again.",
        });
        assert.equal(
          (await paragraphFor(reopened.session, current.id)).text,
          "Edited again.",
        );
      } finally {
        await reopened.end();
      }
    } finally {
      await end();
    }
  });

  it("retains an empty non-painting paragraph through save, reopen, typing and Undo", async () => {
    const { session, end } = await pdfSession(
      await buildPdf([{ texts: [...paragraphObjects, nextParagraph] }]),
    );
    try {
      const before = (await session.getElements({ pageIndex: 0 })).items;
      const paragraph = await paragraphFor(
        session,
        textElement(before, lines[0]).id,
      );
      const suppressed = (await session.renderPageWithout(0, [paragraph.id]))
        .item;
      assert.ok(suppressed);
      await session.replaceParagraphText({ target: paragraph.id, text: "" });
      const empty = await paragraphFor(session, paragraph.id);
      assert.equal(empty.text, "");
      assert.deepEqual(empty.bounds, paragraph.bounds);
      assert.equal((await session.getElement(paragraph.id)).item?.text, "");
      const cleared = (await session.renderPageWithout(0, [])).item;
      assert.ok(cleared);
      assert.deepEqual(
        cleared.data,
        suppressed.data,
        "empty anchor paints no pixels",
      );
      const saved = await session.save();
      assert.equal(
        (await extractPageText(saved.bytes, 0)).trim(),
        nextParagraph.text,
        "no placeholder glyph is exported",
      );

      const reopened = await pdfSession(saved.bytes);
      try {
        assert.deepEqual(
          await paragraphFor(reopened.session, paragraph.id),
          empty,
        );
        assert.equal(
          (await reopened.session.getElement(paragraph.id)).item?.text,
          "",
        );
        const bitmap = (await reopened.session.renderPageWithout(0, [])).item;
        assert.ok(bitmap);
        assert.deepEqual(bitmap.data, suppressed.data);
        await reopened.session.replaceParagraphText({
          target: paragraph.id,
          text: "Typing resumes.",
        });
        assert.equal(
          (await paragraphFor(reopened.session, paragraph.id)).text,
          "Typing resumes.",
        );
        await reopened.session.undo();
        assert.equal(
          (await paragraphFor(reopened.session, paragraph.id)).text,
          "",
        );
        assert.equal(
          reopened.session.state.canUndo,
          false,
          "reopened file starts at its saved empty state",
        );
      } finally {
        await reopened.end();
      }
      await session.replaceParagraphText({
        target: paragraph.id,
        text: "Typing resumes.",
      });
      await session.undo();
      assert.equal((await paragraphFor(session, paragraph.id)).text, "");
      await session.undo();
      assert.deepEqual(
        (await session.getElements({ pageIndex: 0 })).items,
        before,
      );
      assert.equal(session.state.canUndo, false);
    } finally {
      await end();
    }
  });

  it("deletes discontiguous paragraph members together and restores them with one Undo", async () => {
    const interleaved = paragraphObjects.flatMap((row, index) => {
      const column = rightColumn[index];
      return column ? [row, column] : [row];
    });
    const { session, end } = await pdfSession(
      await buildPdf([{ texts: [...interleaved, nextParagraph] }]),
    );
    try {
      const before = (await session.getElements({ pageIndex: 0 })).items;
      const paragraph = await paragraphFor(
        session,
        textElement(before, lines[1]).id,
      );
      const neighbors = before.filter(
        (element) => !paragraph.memberIds.includes(element.id),
      );
      await session.deleteElement({ target: paragraph.id });
      assert.deepEqual(
        (await session.getElements({ pageIndex: 0 })).items,
        neighbors,
      );
      assert.equal((await session.getElement(paragraph.id)).item, undefined);
      for (const memberId of paragraph.memberIds)
        assert.equal((await session.getElement(memberId)).item, undefined);
      const saved = await session.save();
      const extracted = await extractPageText(saved.bytes, 0);
      for (const line of lines) assert.ok(!extracted.includes(line));
      for (const neighbor of neighbors) {
        assert.ok(neighbor.text);
        assert.ok(extracted.includes(neighbor.text.trim()));
      }
      await session.undo();
      assert.deepEqual(
        (await session.getElements({ pageIndex: 0 })).items,
        before,
      );
      assert.deepEqual(await paragraphFor(session, paragraph.id), paragraph);
      assert.equal(session.state.canUndo, false);
    } finally {
      await end();
    }
  });

  it("does not trust persisted logical text that disagrees with the native paragraph lines", async () => {
    const original = await pdfSession(
      await buildPdf([{ texts: paragraphObjects }]),
    );
    try {
      const elements = (await original.session.getElements({ pageIndex: 0 }))
        .items;
      const paragraph = await paragraphFor(
        original.session,
        textElement(elements, lines[0]).id,
      );
      const staleText = "Stale metadata must not replace the visible document.";
      const mark = {
        kind: "paragraph",
        id: paragraph.id,
        rect: paragraph.bounds,
        text: staleText,
        lines,
        style: paragraph.textStyle,
        baselineOffset: (92 - paragraph.bounds.y) / 11,
      };
      const bytes = await buildPdf([
        { texts: paragraphObjects.map((row) => ({ ...row, mark })) },
      ]);
      const reopened = await pdfSession(bytes);
      try {
        const current = (await reopened.session.getElements({ pageIndex: 0 }))
          .items;
        const exposedText = current
          .map((element) => element.text ?? "")
          .join(" ");
        assert.ok(
          !exposedText.includes(staleText),
          "untrusted metadata cannot invent different logical text",
        );
        for (const line of lines)
          assert.ok(
            exposedText.includes(line),
            `actual native text remains exposed: ${line}`,
          );
        assert.ok(
          !(
            await reopened.session.getTextParagraph(paragraph.id)
          ).item?.text.includes(staleText),
        );
        assert.deepEqual(
          (await reopened.session.save()).bytes,
          bytes,
          "reading stale metadata does not change the PDF",
        );
      } finally {
        await reopened.end();
      }
    } finally {
      await original.end();
    }
  });

  it("does not treat a visible path carrying an empty-paragraph mark as editable text", async () => {
    const { session, end } = await pdfSession(
      await buildPdf([{ texts: paragraphObjects }]),
    );
    try {
      const before = (await session.getElements({ pageIndex: 0 })).items;
      const paragraph = await paragraphFor(
        session,
        textElement(before, lines[0]).id,
      );
      const rect = paragraph.bounds;
      const visiblePath = await buildPdf([
        {
          rect: {
            x: rect.x,
            y: 792 - rect.y - rect.height,
            width: rect.width,
            height: rect.height,
            fill: [255, 0, 0],
          },
        },
      ]);
      const pdfium = await fixturePdfium();
      const document = pdfium.openDocument(visiblePath);
      let externallyChanged: Uint8Array;
      try {
        const page = pdfium.lib.FPDF_LoadPage(document.handle, 0);
        try {
          assert.equal(pdfium.lib.FPDFPage_CountObjects(page), 1);
          const anchor = pdfium.lib.FPDFPage_GetObject(page, 0);
          const mark = pdfium.lib.FPDFPageObj_AddMark(anchor, "WebDoc");
          pdfium.lib.FPDFPageObjMark_SetStringParam(
            document.handle,
            anchor,
            mark,
            "webdoc",
            JSON.stringify({
              kind: "paragraph",
              id: paragraph.id,
              rect,
              text: "",
              lines: [],
              style: paragraph.textStyle,
              baselineOffset: (92 - rect.y) / 11,
            }),
          );
          pdfium.lib.FPDFPage_GenerateContent(page);
          externallyChanged = document.save("full");
        } finally {
          pdfium.lib.FPDF_ClosePage(page);
        }
      } finally {
        document.close();
      }
      const reopened = await pdfSession(externallyChanged);
      try {
        assert.equal(
          (await reopened.session.getTextParagraph(paragraph.id)).item,
          undefined,
        );
        const elements = (await reopened.session.getElements({ pageIndex: 0 }))
          .items;
        assert.equal(elements.length, 1);
        assert.equal(elements[0]?.kind, "shape");
        assert.equal(elements[0]?.shapeStyle?.fill?.color, "#ff0000");
        assert.equal(elements[0]?.text, undefined);
        assert.equal(reopened.session.state.dirty, false);
      } finally {
        await reopened.end();
      }
    } finally {
      await end();
    }
  });
});
