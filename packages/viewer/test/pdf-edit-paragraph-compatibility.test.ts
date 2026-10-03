import assert from "node:assert/strict";
import { describe, it } from "node:test";

import type { PdfOperation, TextRange } from "../src/index.js";
import { buildPdf } from "./fixtures/pdf-builder.js";
import { pdfSession } from "./fixtures/pdf-session.js";

const lines = [
  "An original first line has the target content",
  "An original next line has the target content",
  "The third paragraph ends here.",
];
const text = lines.join(" ");

async function setup() {
  const fixture = await pdfSession(
    await buildPdf([
      {
        texts: lines.map((text, index) => ({
          text,
          x: 72,
          y: 700 - index * 20,
          fontSize: 11,
        })),
      },
    ]),
  );
  const rows = (await fixture.session.getElements({ pageIndex: 0 })).items;
  const first = rows.find((row) => row.text?.trim() === lines[0]);
  const second = rows.find((row) => row.text?.trim() === lines[1]);
  assert.ok(first && second);
  const paragraph = (await fixture.session.getTextParagraph(first.id)).item;
  assert.ok(paragraph);
  return { ...fixture, rows, first, second, paragraph };
}

const range = (elementId: string, start: number, end: number): TextRange => ({
  start: { elementId, offset: start },
  end: { elementId, offset: end },
});

describe("PDF paragraph compatibility with existing text contracts", () => {
  it("replaces a findText range through the existing replaceText API after promotion", async () => {
    const { session, end, paragraph } = await setup();
    try {
      assert.ok(
        (await session.getElement(paragraph.id)).item?.operations.includes(
          "replaceText",
        ),
      );
      await session.replaceParagraphText({ target: paragraph.id, text });
      assert.ok(
        (await session.getElement(paragraph.id)).item?.operations.includes(
          "replaceText",
        ),
      );
      const hit = (await session.findText("target")).items[0];
      const at = hit?.ranges[0];
      assert.ok(at);
      assert.equal(at.start.elementId, paragraph.id);
      const receipt = await session.replaceText({
        target: at.start.elementId,
        text: "result",
        range: at,
      });
      assert.deepEqual(receipt.createdIds, []);
      assert.deepEqual(receipt.removedIds, []);
      assert.equal(
        (await session.getElement(paragraph.id)).item?.text,
        text.replace("target", "result"),
      );
      await session.undo();
      assert.equal((await session.getElement(paragraph.id)).item?.text, text);
    } finally {
      await end();
    }
  });

  it("carries an original row selection through promotion, later edits, Undo, Redo and reset", async () => {
    const { session, end, second, paragraph } = await setup();
    try {
      const offset = second.text?.indexOf("target");
      assert.ok(offset !== undefined && offset >= 0);
      const selected = range(second.id, offset, offset + "target".length);
      const member = paragraph.members.find(
        (entry) => entry.elementId === second.id,
      );
      assert.ok(member);
      const replacement = "An expanded";
      const delta = replacement.length - 2;
      await session.replaceParagraphText({
        target: paragraph.id,
        text: replacement,
        range: range(paragraph.id, 0, 2),
      });
      const mapped = range(
        paragraph.id,
        member.start + offset + delta,
        member.start + offset + delta + "target".length,
      );
      assert.deepEqual((await session.mapRange(selected, 0)).item, mapped);
      assert.equal((await session.rangeRects(mapped)).items.length, 1);
      await session.undo();
      assert.deepEqual((await session.mapRange(selected, 0)).item, selected);
      await session.redo();
      assert.deepEqual((await session.mapRange(selected, 0)).item, mapped);
      await session.replaceText({
        target: paragraph.id,
        text: "A",
        range: range(paragraph.id, 0, replacement.length),
      });
      assert.deepEqual(
        (await session.mapRange(selected, 0)).item,
        range(
          paragraph.id,
          member.start + offset - 1,
          member.start + offset - 1 + "target".length,
        ),
      );
      await session.reset();
      assert.deepEqual((await session.mapRange(selected, 0)).item, selected);
    } finally {
      await end();
    }
  });

  it("leaves row anchors and history intact for dry runs and rejected paragraph replacement", async () => {
    const { session, end, second, paragraph, rows } = await setup();
    try {
      const selected = range(second.id, 0, 3);
      const state = session.state;
      const before = await session.save();
      const preview = await session.replaceParagraphText(
        { target: paragraph.id, text: "Preview only" },
        { dryRun: true },
      );
      assert.equal(preview.dryRun, true);
      assert.equal(
        preview.textAnchorMigrations?.length,
        paragraph.members.length,
      );
      assert.deepEqual(session.state, state);
      assert.deepEqual((await session.mapRange(selected, 0)).item, selected);
      await assert.rejects(
        session.replaceParagraphText({
          target: paragraph.id,
          text: "Does not fit ".repeat(1000),
        }),
      );
      assert.deepEqual(session.state, state);
      assert.deepEqual(
        (await session.getElements({ pageIndex: 0 })).items,
        rows,
      );
      assert.deepEqual((await session.save()).bytes, before.bytes);
      assert.deepEqual((await session.mapRange(selected, 0)).item, selected);
      await session.replaceText({
        target: second.id,
        text: "Row replacement remains supported",
      });
      assert.equal(
        (await session.getElement(second.id)).item?.text,
        "Row replacement remains supported",
      );
    } finally {
      await end();
    }
  });

  it("maps both row endpoints at the promotion's operation index in a JSON batch", async () => {
    const { session, end, first, second, paragraph } = await setup();
    try {
      const member = paragraph.members.find(
        (entry) => entry.elementId === second.id,
      );
      assert.ok(member);
      const selected: TextRange = {
        start: { elementId: first.id, offset: 3 },
        end: { elementId: second.id, offset: 10 },
      };
      const operations: PdfOperation[] = [
        {
          op: "setTextStyle",
          target: paragraph.id,
          style: { color: "#123456" },
        },
        {
          op: "replaceText",
          target: paragraph.id,
          range: range(paragraph.id, 0, 2),
          text: "An expanded",
        },
        {
          op: "replaceText",
          target: paragraph.id,
          range: range(paragraph.id, 0, 11),
          text: "A",
        },
      ];
      const receipt = await session.applyJson(operations);
      assert.deepEqual(receipt.createdIds, []);
      assert.deepEqual(receipt.removedIds, paragraph.memberIds);
      assert.deepEqual(
        receipt.textAnchorMigrations,
        paragraph.members.map((entry) => ({
          operationIndex: 1,
          sourceElementId: entry.elementId,
          targetElementId: paragraph.id,
          sourceLength: entry.end - entry.start,
          targetOffset: entry.start,
        })),
      );
      assert.ok(Object.isFrozen(receipt.textAnchorMigrations));
      assert.ok(receipt.textAnchorMigrations?.every(Object.isFrozen));
      const mapped = range(paragraph.id, 2, member.start + 9);
      assert.deepEqual((await session.mapRange(selected, 0)).item, mapped);
      await session.undo();
      assert.deepEqual((await session.mapRange(selected, 0)).item, selected);
      await session.redo();
      assert.deepEqual((await session.mapRange(selected, 0)).item, mapped);
      await session.reset();
      assert.deepEqual((await session.mapRange(selected, 0)).item, selected);
    } finally {
      await end();
    }
  });

  it("rolls back a batch that obstructs the paragraph before reflow without migrating its anchors", async () => {
    const { session, end, second, paragraph, rows } = await setup();
    try {
      const before = await session.save();
      const selected = range(second.id, 0, 3);
      await assert.rejects(
        session.apply([
          {
            op: "insertShape",
            pageIndex: 0,
            shape: "rectangle",
            rect: paragraph.bounds,
            fill: { color: "#ff0000" },
          },
          {
            op: "replaceText",
            target: paragraph.id,
            text: "Obstructed reflow",
          },
        ]),
      );
      assert.equal(session.state.revision, 0);
      assert.equal(session.state.canUndo, false);
      assert.deepEqual(
        (await session.getElements({ pageIndex: 0 })).items,
        rows,
      );
      assert.deepEqual((await session.save()).bytes, before.bytes);
      assert.deepEqual((await session.mapRange(selected, 0)).item, selected);
      await session.replaceText({ target: paragraph.id, text });
      const member = paragraph.members.find(
        (entry) => entry.elementId === second.id,
      );
      assert.ok(member);
      assert.deepEqual(
        (await session.mapRange(selected, 0)).item,
        range(paragraph.id, member.start, member.start + 3),
      );
    } finally {
      await end();
    }
  });

  it("keeps a paragraph with a fractional font size editable after save and reopen", async () => {
    const { session, end, paragraph } = await setup();
    try {
      await session.setTextStyle({
        target: paragraph.id,
        style: { fontSize: 11.123456 },
      });
      const saved = await session.save();
      const reopened = await pdfSession(saved.bytes);
      try {
        const current = (await reopened.session.getTextParagraph(paragraph.id))
          .item;
        assert.ok(current);
        assert.equal(current.text, text);
        assert.equal(current.textStyle.fontSize, 11.123);
        await reopened.session.replaceText({
          target: current.id,
          text: "Still editable.",
        });
        assert.equal(
          (await reopened.session.getElement(current.id)).item?.text,
          "Still editable.",
        );
      } finally {
        await reopened.end();
      }
    } finally {
      await end();
    }
  });
});
