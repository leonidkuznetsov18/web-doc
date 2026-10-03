import assert from "node:assert/strict";
import { describe, it } from "node:test";

import { DocxEditEngine } from "../src/edit/docx/engine.js";
import { loadDocxEditEngine } from "../src/edit/docx/provider.js";
import { createOoxmlEditHandler } from "../src/edit/pptx/handler.js";
import { OoxmlPackage } from "../src/edit/ooxml/package.js";
import { scanXml } from "../src/edit/ooxml/xml.js";
import { defaultResourceLimits, ViewerError } from "../src/index.js";
import { buildDocx, paragraph, sectPr } from "./fixtures/docx-builder.js";
import { docxSession } from "./fixtures/docx-session.js";
import { loopbackWorker } from "./fixtures/loopback-worker.js";

const limits = defaultResourceLimits;
const signal = new AbortController().signal;
const ORIGINAL = buildDocx({
  body:
    "<w:p><w:r><w:rPr><w:b/></w:rPr><w:t>Bold</w:t></w:r>" +
    '<w:r><w:rPr><w:i/></w:rPr><w:t xml:space="preserve"> italic</w:t></w:r></w:p>' +
    paragraph("Following paragraph") +
    sectPr(),
});

async function documentXml(bytes: Uint8Array): Promise<string> {
  const pkg = await OoxmlPackage.open(bytes, { limits });
  return new TextDecoder().decode(await pkg.part("/word/document.xml"));
}

async function paragraphId(engine: DocxEditEngine): Promise<string> {
  const first = (await engine.getElements({}, signal))[0];
  assert.ok(first);
  return first.id;
}

function range(elementId: string, start: number, end = start) {
  return {
    start: { elementId, offset: start },
    end: { elementId, offset: end },
  };
}

describe("DOCX read-only text draft preview", () => {
  it("previews pending bold, italic and underline only on inserted text without changing live runs", async () => {
    const pair = loopbackWorker(createOoxmlEditHandler());
    const engine = await loadDocxEditEngine(
      ORIGINAL,
      { format: "docx", limits, signal },
      { createWorker: () => pair.worker },
    );
    try {
      const before = (await engine.getElements({}, signal))[0];
      assert.ok(before);
      const saved = (await engine.materializeDocument("save", {}, signal))
        .bytes;
      const fields = {
        target: before.id,
        text: "🙂",
        range: range(before.id, 4),
        insertionStyle: { bold: false, italic: true, underline: true },
      };
      const draft = await engine.previewDraft(fields, signal);
      const preview = await DocxEditEngine.open(draft.bytes, limits, signal);
      try {
        assert.equal(
          (await preview.getElement(before.id, signal))?.text,
          "Bold🙂 italic",
        );
        const inserted = await preview.textStyle(
          before.id,
          { start: 4, end: 6 },
          signal,
        );
        assert.equal(inserted?.bold, false);
        assert.equal(inserted?.italic, true);
        assert.equal(inserted?.underline, true);
        assert.equal(
          (await preview.textStyle(before.id, { start: 0, end: 4 }, signal))
            ?.bold,
          true,
        );
        assert.equal(
          (await preview.textStyle(before.id, { start: 6, end: 13 }, signal))
            ?.italic,
          true,
        );
      } finally {
        await preview.dispose();
      }
      assert.deepEqual(await engine.getElement(before.id, signal), before);
      assert.deepEqual(
        (await engine.materializeDocument("save", {}, signal)).bytes,
        saved,
      );
    } finally {
      await engine.dispose();
    }
  });
  it("styles empty paragraphs without consuming live ids", async () => {
    const original = buildDocx({
      body: paragraph("") + paragraph("next") + sectPr(),
    });
    const engine = await DocxEditEngine.open(original, limits, signal);
    try {
      const target = await paragraphId(engine);
      for (const insertionStyle of [
        { bold: true },
        { bold: true, italic: true, underline: true },
      ]) {
        const draft = await engine.previewDraft(
          { target, text: "one🙂two", insertionStyle },
          signal,
        );
        const parsed = await DocxEditEngine.open(draft.bytes, limits, signal);
        try {
          assert.equal(
            (await parsed.getElement(target, signal))?.text,
            "one🙂two",
          );
          const style = await parsed.textStyle(
            target,
            { start: 0, end: 8 },
            signal,
          );
          assert.equal(style?.bold, true);
          if (insertionStyle.italic) {
            assert.equal(style?.italic, true);
            assert.equal(style?.underline, true);
          }
          assert.equal((await parsed.getElements({}, signal)).length, 2);
        } finally {
          await parsed.dispose();
        }
      }
      assert.deepEqual(
        (await engine.materializeDocument("save", {}, signal)).bytes,
        original,
      );
    } finally {
      await engine.dispose();
    }
  });

  it("rejects empty, invalid, cyclic and surrogate-splitting styled drafts without altering the session", async () => {
    const original = buildDocx({ body: paragraph("a🙂b") + sectPr() });
    const { session, end } = await docxSession(original, [[]]);
    try {
      const target = (await session.getElements()).items[0]?.id;
      assert.ok(target);
      const before = session.state;
      const cyclic = { bold: true };
      Object.assign(cyclic, { extra: cyclic });
      for (const fields of [
        { target, text: "", insertionStyle: { bold: true } },
        { target, text: "x", insertionStyle: { fontSize: 0 } },
        {
          target,
          text: "x",
          range: range(target, 2),
          insertionStyle: { bold: true },
        },
      ]) {
        await assert.rejects(
          async () => session.previewText(fields),
          (error: unknown) =>
            error instanceof ViewerError && error.code === "invalid-operation",
        );
        await assert.rejects(
          session.apply([{ op: "replaceText", ...fields }]),
          (error: unknown) =>
            error instanceof ViewerError && error.code === "invalid-operation",
        );
        assert.deepEqual(session.state, before);
        assert.deepEqual((await session.save()).bytes, original);
      }
      await assert.rejects(
        async () =>
          session.previewText({ target, text: "x", insertionStyle: cyclic }),
        (error: unknown) =>
          error instanceof ViewerError && error.code === "invalid-operation",
      );
      assert.deepEqual(session.state, before);
      assert.deepEqual((await session.save()).bytes, original);
      const aborted = new AbortController();
      aborted.abort();
      await assert.rejects(
        session.previewText(
          { target, text: "x", insertionStyle: { bold: true } },
          { signal: aborted.signal },
        ),
      );
      assert.deepEqual(session.state, before);
      assert.deepEqual((await session.save()).bytes, original);
    } finally {
      await end();
    }
  });

  it("commits styled appended text after sequential validation and keeps invalid later operations atomic", async () => {
    const { session, end } = await docxSession(ORIGINAL, [[]]);
    try {
      const target = (await session.getElements()).items[0]?.id;
      assert.ok(target);
      const styled = [
        {
          op: "replaceText" as const,
          target,
          text: "🙂",
          range: range(target, 11),
        },
        {
          op: "setTextStyle" as const,
          target,
          range: range(target, 11, 13),
          style: { bold: true, italic: true, underline: true },
        },
      ];
      const before = session.state;
      const saved = await session.save();
      const dry = await session.apply(styled, { dryRun: true });
      assert.equal(dry.dryRun, true);
      assert.deepEqual(session.state, before);
      assert.deepEqual(await session.save(), saved);
      await session.apply(styled);
      assert.equal(
        (await session.getElement(target)).item?.text,
        "Bold italic🙂",
      );
      const actualStyle = (
        await session.getTextStyle({ target, range: range(target, 11, 13) })
      ).item;
      assert.equal(actualStyle?.bold, true);
      assert.equal(actualStyle?.italic, true);
      assert.equal(actualStyle?.underline, true);
      const after = session.state;
      const afterSaved = await session.save();
      const checkpoint = await session.createCheckpoint("before failed style");
      const bad = [
        { op: "replaceText" as const, target, text: "x" },
        {
          op: "setTextStyle" as const,
          target,
          range: range(target, 11, 13),
          style: { bold: true },
        },
      ];
      await assert.rejects(
        session.apply(bad),
        (error: unknown) =>
          error instanceof ViewerError && error.code === "invalid-operation",
      );
      assert.deepEqual(session.state, after);
      assert.deepEqual(await session.save(), afterSaved);
      assert.deepEqual(session.listCheckpoints(), [checkpoint]);
      const split = [
        { op: "replaceText" as const, target, text: "a\nb" },
        {
          op: "setTextStyle" as const,
          target,
          range: range(target, 0, 3),
          style: { italic: true },
        },
      ];
      await assert.rejects(
        session.apply(split),
        (error: unknown) =>
          error instanceof ViewerError && error.code === "invalid-operation",
      );
      assert.deepEqual(session.state, after);
      assert.deepEqual(await session.save(), afterSaved);
      await session.undo();
      assert.deepEqual((await session.save()).bytes, saved.bytes);
      await session.redo();
      assert.deepEqual((await session.save()).bytes, afterSaved.bytes);
    } finally {
      await end();
    }
  });

  it("atomically styles multiple inserted paragraphs while preserving mixed tail runs through save and history", async () => {
    const { session, end } = await docxSession(ORIGINAL, [[]]);
    try {
      const target = (await session.getElements()).items[0]?.id;
      assert.ok(target);
      const fields = {
        target,
        text: "x\r\n\r\n🙂y",
        range: range(target, 4),
        insertionStyle: { italic: true, underline: true },
      };
      const operation = { op: "replaceText" as const, ...fields };
      const before = session.state;
      const saved = await session.save();
      const draft = await session.previewText(fields);
      assert.ok(draft.item);
      const preview = await DocxEditEngine.open(draft.item, limits, signal);
      try {
        const paragraphs = await preview.getElements({}, signal);
        assert.deepEqual(
          paragraphs.map((p) => p.text),
          ["Boldx", "", "🙂y italic", "Following paragraph"],
        );
        assert.equal(
          (await preview.textStyle(target, { start: 4, end: 5 }, signal))
            ?.underline,
          true,
        );
        const empty = paragraphs[1];
        const last = paragraphs[2];
        assert.ok(empty && last);
        assert.equal(
          (await preview.textStyle(empty.id, undefined, signal))?.underline,
          true,
        );
        assert.equal(
          (await preview.textStyle(empty.id, undefined, signal))?.bold,
          true,
        );
        assert.equal(
          (await preview.textStyle(last.id, { start: 0, end: 3 }, signal))
            ?.underline,
          true,
        );
        assert.equal(
          (await preview.textStyle(last.id, { start: 3, end: 10 }, signal))
            ?.underline,
          false,
        );
      } finally {
        await preview.dispose();
      }
      assert.deepEqual(session.state, before);
      assert.deepEqual(await session.save(), saved);
      const dry = await session.apply([operation], { dryRun: true });
      assert.equal(dry.createdIds.length, 2);
      assert.deepEqual(session.state, before);
      const actual = await session.apply([operation]);
      assert.deepEqual(actual.createdIds, dry.createdIds);
      const committed = await session.save();
      const reopened = await DocxEditEngine.open(
        committed.bytes,
        limits,
        signal,
      );
      try {
        assert.deepEqual(
          (await reopened.getElements({}, signal)).map((p) => p.text),
          ["Boldx", "", "🙂y italic", "Following paragraph"],
        );
      } finally {
        await reopened.dispose();
      }
      await session.undo();
      assert.deepEqual((await session.save()).bytes, saved.bytes);
      await session.redo();
      assert.deepEqual((await session.save()).bytes, committed.bytes);
    } finally {
      await end();
    }
  });

  it("styles a leading split's new empty mark and preserves the original last mark and paragraph properties", async () => {
    const originalProperties =
      '<w:pPr><w:spacing w:after="160"/><w:jc w:val="center"/><w:rPr><w:rFonts w:ascii="Courier New" w:hAnsi="Courier New"/><w:u w:val="none"/></w:rPr>' +
      "</w:pPr>";
    const original = buildDocx({
      body:
        "<w:p>" +
        originalProperties +
        "<w:r><w:t>AB</w:t></w:r></w:p>" +
        paragraph("following") +
        sectPr(),
    });
    const { session, end } = await docxSession(original, [[]]);
    try {
      const target = (await session.getElements()).items[0]?.id;
      assert.ok(target);
      const fields = {
        target,
        text: "\nX",
        range: range(target, 0),
        insertionStyle: { bold: true, italic: true, underline: true },
      };
      const before = session.state;
      const preview = await session.previewText(fields);
      assert.ok(preview.item);
      assert.deepEqual(session.state, before);
      assert.deepEqual((await session.save()).bytes, original);
      const committed = await session.replaceText(fields);
      assert.equal(committed.createdIds.length, 1);
      for (const bytes of [preview.item, (await session.save()).bytes]) {
        const engine = await DocxEditEngine.open(bytes, limits, signal);
        try {
          const elements = await engine.getElements({}, signal);
          assert.deepEqual(
            elements.map((e) => e.text),
            ["", "XAB", "following"],
          );
          assert.equal(elements[0]?.id, target);
          const mark = await engine.textStyle(target, undefined, signal);
          assert.equal(mark?.bold, true);
          assert.equal(mark?.italic, true);
          assert.equal(mark?.underline, true);
          const last = elements[1];
          assert.ok(last);
          assert.equal(
            (await engine.textStyle(last.id, { start: 0, end: 1 }, signal))
              ?.underline,
            true,
          );
          assert.equal(
            (await engine.textStyle(last.id, { start: 1, end: 3 }, signal))
              ?.underline,
            false,
          );
          const xml = scanXml("/word/document.xml", await documentXml(bytes));
          const paragraphs = xml.findAll("w:p");
          const firstProperties = paragraphs[0]?.children.find(
            (c) => c.local === "pPr",
          );
          const lastProperties = paragraphs[1]?.children.find(
            (c) => c.local === "pPr",
          );
          assert.ok(firstProperties && lastProperties);
          const firstXml = xml.text.slice(
            firstProperties.start,
            firstProperties.end,
          );
          assert.ok(firstXml.includes('<w:spacing w:after="160"/>'));
          assert.ok(firstXml.includes('<w:jc w:val="center"/>'));
          assert.ok(
            !firstProperties.children.some((c) => c.local === "sectPr"),
          );
          assert.equal(
            xml.text.slice(lastProperties.start, lastProperties.end),
            originalProperties,
          );
          assert.equal(xml.findAll("w:sectPr").length, 1);
        } finally {
          await engine.dispose();
        }
      }
      await session.undo();
      assert.deepEqual((await session.save()).bytes, original);
    } finally {
      await end();
    }
  });

  it("keeps section-break paragraphs read-only and preserves their original section bytes on a styled split", async () => {
    const original = buildDocx({
      body:
        "<w:p><w:pPr>" +
        sectPr({ width: 10000 }) +
        "</w:pPr><w:r><w:t>AB</w:t></w:r></w:p>" +
        sectPr(),
    });
    const { session, end } = await docxSession(original, [[]]);
    try {
      const target = (await session.getElements()).items[0]?.id;
      assert.ok(target);
      const before = session.state;
      const fields = {
        target,
        text: "\nX",
        range: range(target, 0),
        insertionStyle: { bold: true, italic: true, underline: true },
      };
      for (const request of [
        () => session.previewText(fields),
        () => session.replaceText(fields),
      ]) {
        await assert.rejects(
          request(),
          (error: unknown) =>
            error instanceof ViewerError &&
            error.code === "invalid-operation" &&
            error.message.includes("section-break"),
        );
        assert.deepEqual(session.state, before);
        assert.deepEqual((await session.save()).bytes, original);
      }
    } finally {
      await end();
    }
  });

  it("retains insertion styling on tracked split runs and newly inserted paragraph marks", async () => {
    const engine = await DocxEditEngine.open(ORIGINAL, limits, signal);
    try {
      const target = await paragraphId(engine);
      const operation = {
        op: "replaceText" as const,
        target,
        range: range(target, 4),
        text: "x\n\n🙂y",
        insertionStyle: { bold: false, italic: true, underline: true },
      };
      await engine.apply(
        {
          stateId: 1,
          operations: [operation],
          changeMode: "tracked",
          author: "QA",
        },
        signal,
      );
      const elements = await engine.getElements({}, signal);
      const empty = elements[1];
      const last = elements[2];
      assert.ok(empty && last);
      assert.equal(
        (await engine.textStyle(target, { start: 4, end: 5 }, signal))
          ?.underline,
        true,
      );
      assert.equal(
        (await engine.textStyle(empty.id, undefined, signal))?.underline,
        true,
      );
      assert.equal(
        (await engine.textStyle(last.id, { start: 0, end: 3 }, signal))
          ?.underline,
        true,
      );
      assert.equal(
        (await engine.textStyle(last.id, { start: 3, end: 10 }, signal))
          ?.underline,
        false,
      );
      assert.ok(
        (await engine.revisions(target, signal)).some((r) => r.kind === "ins"),
      );
    } finally {
      await engine.dispose();
    }
  });

  it("returns draft target metadata with its display bytes through the edit worker", async () => {
    const pair = loopbackWorker(createOoxmlEditHandler());
    const engine = await loadDocxEditEngine(
      ORIGINAL,
      { format: "docx", limits, signal },
      { createWorker: () => pair.worker },
    );
    try {
      const before = (await engine.getElements({}, signal))[0];
      assert.ok(before);
      const draft = await engine.previewDraft(
        { target: before.id, text: "draft" },
        signal,
      );
      assert.equal(draft.paragraph?.id, before.id);
      assert.equal(draft.paragraph?.text, "draft");
      assert.equal(draft.paragraph?.textStyle?.bold, true);
      const parsed = await DocxEditEngine.open(draft.bytes, limits, signal);
      try {
        assert.deepEqual(
          await parsed.getElement(before.id, signal),
          draft.paragraph,
        );
      } finally {
        await parsed.dispose();
      }
      assert.deepEqual(await engine.getElement(before.id, signal), before);
    } finally {
      await engine.dispose();
    }
  });
  it("preserves mixed runs in display bytes while saved bytes, ids and live styles stay unchanged", async () => {
    const engine = await DocxEditEngine.open(ORIGINAL, limits, signal);
    try {
      const target = await paragraphId(engine);
      const before = await engine.getElements({}, signal);
      const bytes = await engine.previewText(
        { target, text: "!", range: range(target, 4) },
        signal,
      );
      const xml = await documentXml(bytes);
      assert.match(xml, /<w:b\/>/);
      assert.match(xml, /<w:i\/>/);
      assert.match(xml, /<w:t>!<\/w:t>/);
      assert.match(xml, / italic<\/w:t>/);
      assert.match(xml, /Following paragraph/);
      const preview = await DocxEditEngine.open(bytes, limits, signal);
      try {
        assert.equal(
          (await preview.getElement(target, signal))?.text,
          "Bold! italic",
        );
        assert.equal(
          (await preview.textStyle(target, { start: 0, end: 5 }, signal))?.bold,
          true,
        );
        assert.equal(
          (await preview.textStyle(target, { start: 5, end: 12 }, signal))
            ?.italic,
          true,
        );
      } finally {
        await preview.dispose();
      }
      assert.deepEqual(await engine.materialize("save", {}, signal), ORIGINAL);
      assert.deepEqual(await engine.getElements({}, signal), before);
    } finally {
      await engine.dispose();
    }
  });

  it("bases every preview on the current committed document and never consumes paragraph ids", async () => {
    const engine = await DocxEditEngine.open(ORIGINAL, limits, signal);
    const control = await DocxEditEngine.open(ORIGINAL, limits, signal);
    try {
      const target = await paragraphId(engine);
      const committed = {
        op: "replaceText" as const,
        target,
        text: "+",
        range: range(target, 0),
      };
      await engine.apply([committed], signal);
      await control.apply([committed], signal);
      const saved = await engine.materialize("save", {}, signal);
      const fields = {
        target,
        text: "first\nsecond",
        range: range(target, 1, 5),
      };
      const first = await engine.previewText(fields, signal);
      const repeated = await engine.previewText(fields, signal);
      assert.deepEqual(repeated, first);
      const different = await engine.previewText(
        { target, text: "other", range: range(target, 1, 5) },
        signal,
      );
      const preview = await DocxEditEngine.open(different, limits, signal);
      try {
        assert.equal(
          (await preview.getElement(target, signal))?.text,
          "+other italic",
        );
      } finally {
        await preview.dispose();
      }
      assert.deepEqual(await engine.materialize("save", {}, signal), saved);
      const operation = { op: "replaceText" as const, ...fields };
      assert.deepEqual(
        (await engine.apply([operation], signal)).createdIds,
        (await control.apply([operation], signal)).createdIds,
      );
      assert.deepEqual(
        await engine.materialize("save", {}, signal),
        await control.materialize("save", {}, signal),
      );
    } finally {
      await engine.dispose();
      await control.dispose();
    }
  });

  it("rejects cyclic draft fields before copying without changing the live package", async () => {
    const engine = await DocxEditEngine.open(ORIGINAL, limits, signal);
    try {
      const target = await paragraphId(engine);
      const fields = { target, text: "draft" };
      Object.assign(fields, { extra: fields });
      await assert.rejects(
        engine.previewText(fields, signal),
        (error: unknown) =>
          error instanceof ViewerError && error.code === "invalid-operation",
      );
      assert.deepEqual(
        (await engine.materializeDocument("save", {}, signal)).bytes,
        ORIGINAL,
      );
    } finally {
      await engine.dispose();
    }
  });

  it("leaves live content safe after invalid or cancelled previews", async () => {
    const engine = await DocxEditEngine.open(ORIGINAL, limits, signal);
    try {
      const target = await paragraphId(engine);
      await assert.rejects(
        engine.previewText({ target, text: "bad\u0000text" }, signal),
        (error: unknown) =>
          error instanceof ViewerError && error.code === "invalid-operation",
      );
      const cancelled = new AbortController();
      cancelled.abort();
      await assert.rejects(
        engine.previewText({ target, text: "discard" }, cancelled.signal),
      );
      assert.deepEqual(await engine.materialize("save", {}, signal), ORIGINAL);
      const committed = { op: "replaceText" as const, target, text: "usable" };
      await engine.apply([committed], signal);
      assert.equal((await engine.getElement(target, signal))?.text, "usable");
    } finally {
      await engine.dispose();
    }
  });

  it("serves preview through the worker without changing revision, dirty state or Undo/Redo", async () => {
    const { session, end } = await docxSession(ORIGINAL, [[]]);
    try {
      const target = (await session.getElements()).items[0]?.id;
      assert.ok(target);
      await session.replaceText({ target, text: "+", range: range(target, 0) });
      const checkpoint = await session.createCheckpoint(
        "committed before draft",
      );
      const saved = await session.save();
      const before = session.state;
      const result = await session.previewText({
        target,
        text: "draft",
        range: range(target, 1, 5),
        insertionStyle: { bold: true, italic: true },
      });
      assert.equal(result.sessionId, before.sessionId);
      assert.equal(result.revision, before.revision);
      assert.ok(result.item);
      const preview = await DocxEditEngine.open(result.item, limits, signal);
      try {
        assert.equal(
          (await preview.getElement(target, signal))?.text,
          "+draft italic",
        );
        assert.equal(
          (await preview.textStyle(target, { start: 1, end: 6 }, signal))?.bold,
          true,
        );
        assert.equal(
          (await preview.textStyle(target, { start: 1, end: 6 }, signal))
            ?.italic,
          true,
        );
      } finally {
        await preview.dispose();
      }
      assert.deepEqual(session.state, before);
      assert.deepEqual(session.listCheckpoints(), [checkpoint]);
      assert.deepEqual(await session.save(), saved);
      await session.undo();
      assert.deepEqual((await session.save()).bytes, ORIGINAL);
      const redoState = session.state;
      await session.previewText({ target, text: "discard" });
      assert.deepEqual(session.state, redoState);
      await session.redo();
      assert.deepEqual((await session.save()).bytes, saved.bytes);
      const cancelled = new AbortController();
      cancelled.abort();
      await assert.rejects(
        session.previewText(
          { target, text: "discard" },
          { signal: cancelled.signal },
        ),
      );
      assert.deepEqual((await session.save()).bytes, saved.bytes);
    } finally {
      await end();
    }
  });
});
