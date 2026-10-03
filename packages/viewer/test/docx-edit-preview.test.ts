import assert from "node:assert/strict";
import { describe, it } from "node:test";

import { DocxEditEngine } from "../src/edit/docx/engine.js";
import { loadDocxEditEngine } from "../src/edit/docx/provider.js";
import { createOoxmlEditHandler } from "../src/edit/pptx/handler.js";
import { OoxmlPackage } from "../src/edit/ooxml/package.js";
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
