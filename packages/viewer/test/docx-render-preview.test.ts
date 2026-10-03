import assert from "node:assert/strict";
import { describe, it } from "node:test";

import type { DocumentAdapter, TextRun } from "../src/contracts.js";
import { DocxEditEngine } from "../src/edit/docx/engine.js";
import { DocxSession } from "../src/edit/docx/session.js";
import { ViewerClient, defaultResourceLimits } from "../src/index.js";
import { buildDocx, paragraph, sectPr } from "./fixtures/docx-builder.js";

/** The test adapter only records target identity; no raster output is simulated. */
class RenderTarget extends EventTarget implements OffscreenCanvas {
  width = 1;
  height = 1;
  oncontextlost = null;
  oncontextrestored = null;
  getContext(): null {
    return null;
  }
  convertToBlob(): Promise<Blob> {
    throw new Error("Not used by recording adapter");
  }
  transferToImageBitmap(): ImageBitmap {
    throw new Error("Not used by recording adapter");
  }
}

const ORIGINAL = buildDocx({ body: paragraph("original") + sectPr() });

function setup() {
  const closed: DocxEditEngine[] = [];
  const opened: DocxEditEngine[] = [];
  const painted: {
    text: string;
    target: HTMLCanvasElement | OffscreenCanvas;
  }[] = [];
  const controls: { failRender: boolean; afterRender?: () => void } = {
    failRender: false,
  };
  const adapter: DocumentAdapter<DocxEditEngine> = {
    id: "recording-docx-preview",
    formats: ["docx"],
    async open(bytes, context) {
      context.reportProgress({ phase: "loading", loaded: 1 });
      const engine = await DocxEditEngine.open(
        bytes,
        context.limits,
        context.signal,
      );
      opened.push(engine);
      return engine;
    },
    async getInfo() {
      return {
        format: "docx",
        unit: "page",
        pageCount: 1,
        pageSizes: [{ width: 816, height: 1056 }],
      };
    },
    async getTextMap(engine) {
      const element = (
        await engine.getElements({}, new AbortController().signal)
      )[0];
      assert.ok(element);
      const text = element.text ?? "";
      return [
        {
          text,
          paragraphId: element.id.slice(2),
          x: 72,
          y: 72,
          width: text.length * 8,
          height: 16,
          fontSize: 16,
          textLayer: "docx",
        },
      ] satisfies TextRun[];
    },
    async render(engine, target) {
      if (controls.failRender) throw new Error("Draft render failed");
      const element = (
        await engine.getElements({}, new AbortController().signal)
      )[0];
      painted.push({ text: element?.text ?? "", target });
      controls.afterRender?.();
    },
    async close(engine) {
      closed.push(engine);
      await engine.dispose();
    },
    edit: {
      formats: ["docx"],
      load: (bytes, context) =>
        DocxEditEngine.open(bytes, context.limits, context.signal),
      createSession: (core, access) => new DocxSession(core, access),
    },
  };
  const viewer = ViewerClient.create({
    adapters: [adapter],
    limits: defaultResourceLimits,
  }).createViewer();
  const events: string[] = [];
  for (const type of [
    "statechange",
    "progress",
    "warning",
    "documentchange",
    "layoutchange",
    "editstatechange",
  ] as const)
    viewer.on(type, () => events.push(type));
  return { viewer, opened, closed, painted, controls, events };
}

describe("DOCX renderer draft preview isolation", () => {
  it("renders draft-owned pixels and geometry without changing the shown document or emitting events", async () => {
    const test = setup();
    try {
      await test.viewer.load(ORIGINAL, { fileName: "preview.docx" });
      const session = await test.viewer.edit();
      assert.ok(session.format === "docx");
      const target = (await session.getElements()).items[0]?.id;
      assert.ok(target);
      const state = session.state;
      const viewerState = test.viewer.state;
      test.events.length = 0;
      const canvas = new RenderTarget();
      const preview = await session.previewTextPages(
        { target, text: "longer draft text" },
        {
          pages: [{ pageIndex: 0, target: canvas }],
          zoom: 1,
          devicePixelRatio: 1,
        },
      );
      assert.equal(preview.revision, state.revision);
      assert.equal(preview.item?.layout?.lines[0]?.text, "longer draft text");
      assert.equal(
        preview.item?.layout?.frame?.width,
        "longer draft text".length * 8,
      );
      assert.deepEqual(preview.item?.pageSizes, [{ width: 816, height: 1056 }]);
      assert.deepEqual(test.painted, [
        { text: "longer draft text", target: canvas },
      ]);
      assert.deepEqual(session.state, state);
      assert.deepEqual(test.viewer.state, viewerState);
      assert.deepEqual(test.events, []);
      assert.deepEqual(test.closed, [test.opened[1]]);
      assert.equal(await test.viewer.getPageText(0), "original");
      assert.deepEqual((await session.save()).bytes, ORIGINAL);
    } finally {
      await test.viewer.destroy();
    }
  });

  it("awaits temporary handle cleanup when rendering fails or a late abort invalidates the result", async () => {
    const test = setup();
    try {
      await test.viewer.load(ORIGINAL, { fileName: "preview.docx" });
      const session = await test.viewer.edit();
      assert.ok(session.format === "docx");
      const target = (await session.getElements()).items[0]?.id;
      assert.ok(target);
      const render = { pages: [{ pageIndex: 0, target: new RenderTarget() }] };
      test.controls.failRender = true;
      await assert.rejects(
        session.previewTextPages({ target, text: "failed" }, render),
        /Draft render failed/,
      );
      assert.deepEqual(test.closed, [test.opened[1]]);
      test.controls.failRender = false;
      const controller = new AbortController();
      test.controls.afterRender = () => controller.abort();
      await assert.rejects(
        session.previewTextPages({ target, text: "cancelled" }, render, {
          signal: controller.signal,
        }),
      );
      assert.deepEqual(test.closed, [test.opened[1], test.opened[2]]);
      assert.equal(session.state.revision, 0);
      assert.deepEqual((await session.save()).bytes, ORIGINAL);
      assert.equal(await test.viewer.getPageText(0), "original");
    } finally {
      await test.viewer.destroy();
    }
  });
});
