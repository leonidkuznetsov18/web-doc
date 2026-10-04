import assert from "node:assert/strict";
import { describe, it } from "node:test";

import type { DocumentAdapter, PageSize, TextRun } from "../src/contracts.js";
import { DocxEditEngine } from "../src/edit/docx/engine.js";
import { DocxSession } from "../src/edit/docx/session.js";
import {
  ViewerClient,
  defaultResourceLimits,
  type DocxTextStyle,
} from "../src/index.js";
import { buildDocx, paragraph, sectPr } from "./fixtures/docx-builder.js";

/** Records target dimensions and identity; no glyph raster output is simulated. */
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

function setup(limits = defaultResourceLimits) {
  const renderedStyle: { current: DocxTextStyle | undefined } = {
    current: undefined,
  };
  const closed: DocxEditEngine[] = [];
  const opened: DocxEditEngine[] = [];
  const painted: {
    text: string;
    target: HTMLCanvasElement | OffscreenCanvas;
  }[] = [];
  const readPages: number[] = [];
  const controls: {
    failRender: boolean;
    afterRender?: () => void;
    pageCount?: number;
    pageSizes?: readonly PageSize[] | null;
    textMap?: (
      engine: DocxEditEngine,
      pageIndex: number,
    ) => Promise<readonly TextRun[]>;
  } = {
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
        pageCount: controls.pageCount ?? 1,
        ...(controls.pageSizes === null
          ? {}
          : {
              pageSizes: controls.pageSizes ?? [{ width: 816, height: 1056 }],
            }),
      };
    },
    async getTextMap(engine, pageIndex) {
      readPages.push(pageIndex);
      if (controls.textMap) return controls.textMap(engine, pageIndex);
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
    async render(engine, target, viewport) {
      if (controls.failRender) throw new Error("Draft render failed");
      const element = (
        await engine.getElements({}, new AbortController().signal)
      )[0];
      const size = controls.pageSizes?.[viewport.pageIndex] ?? {
        width: 816,
        height: 1056,
      };
      target.width = Math.ceil(
        size.width * viewport.zoom * viewport.devicePixelRatio,
      );
      target.height = Math.ceil(
        size.height * viewport.zoom * viewport.devicePixelRatio,
      );
      renderedStyle.current = element?.textStyle;
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
    limits,
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
  return {
    viewer,
    opened,
    closed,
    painted,
    controls,
    events,
    renderedStyle,
    readPages,
  };
}

describe("DOCX renderer draft preview isolation", () => {
  it("rejects a later over-budget draft page before mutating any earlier caller canvas", async () => {
    const test = setup();
    try {
      await test.viewer.load(ORIGINAL, { fileName: "preview.docx" });
      const session = await test.viewer.edit();
      assert.ok(session.format === "docx");
      const target = (await session.getElements()).items[0]?.id;
      assert.ok(target);
      const state = session.state;
      test.events.length = 0;
      test.controls.pageCount = 2;
      test.controls.pageSizes = [
        { width: 10, height: 10 },
        { width: 20, height: 20 },
      ];
      const first = new RenderTarget();
      const second = new RenderTarget();
      first.width = 13;
      first.height = 17;
      second.width = 19;
      second.height = 23;
      await assert.rejects(
        session.previewTextPages(
          { target, text: "draft" },
          {
            pages: [
              { pageIndex: 0, target: first },
              { pageIndex: 1, target: second },
            ],
            zoom: 1,
            devicePixelRatio: 1,
            maxPixelsPerPage: 100,
          },
        ),
        { code: "resource-limit" },
      );
      assert.deepEqual(
        [first.width, first.height, second.width, second.height],
        [13, 17, 19, 23],
      );
      assert.deepEqual(test.painted, []);
      assert.deepEqual(test.closed, [test.opened[1]]);
      assert.deepEqual(session.state, state);
      assert.deepEqual(test.events, []);
      assert.deepEqual((await session.save()).bytes, ORIGINAL);
    } finally {
      await test.viewer.destroy();
    }
  });

  it("rejects invalid explicit draft budgets before opening a temporary renderer and preserves history", async () => {
    const test = setup();
    try {
      await test.viewer.load(ORIGINAL, { fileName: "preview.docx" });
      const session = await test.viewer.edit();
      assert.ok(session.format === "docx");
      const target = (await session.getElements()).items[0]?.id;
      assert.ok(target);
      const state = session.state;
      test.events.length = 0;
      for (const maxPixelsPerPage of [
        0,
        -1,
        1.5,
        NaN,
        Infinity,
        Number.MAX_SAFE_INTEGER + 1,
      ]) {
        const canvas = new RenderTarget();
        await assert.rejects(
          session.previewTextPages(
            { target, text: "draft" },
            { pages: [{ pageIndex: 0, target: canvas }], maxPixelsPerPage },
          ),
          { code: "invalid-operation" },
        );
        assert.deepEqual([canvas.width, canvas.height], [1, 1]);
      }
      assert.equal(test.opened.length, 1);
      assert.deepEqual(test.painted, []);
      assert.deepEqual(test.events, []);
      assert.deepEqual(session.state, state);
      assert.deepEqual((await session.save()).bytes, ORIGINAL);
    } finally {
      await test.viewer.destroy();
    }
  });

  it("enforces explicit draft budgets using rounded raster dimensions and keeps metadata and cleanup", async () => {
    const test = setup();
    try {
      await test.viewer.load(ORIGINAL, { fileName: "preview.docx" });
      const session = await test.viewer.edit();
      assert.ok(session.format === "docx");
      const target = (await session.getElements()).items[0]?.id;
      assert.ok(target);
      const state = session.state;
      test.events.length = 0;
      test.controls.pageSizes = [{ width: 10.25, height: 20.25 }];
      const refused = new RenderTarget();
      await assert.rejects(
        session.previewTextPages(
          { target, text: "draft" },
          {
            pages: [{ pageIndex: 0, target: refused }],
            zoom: 1,
            devicePixelRatio: 2,
            maxPixelsPerPage: 860,
          },
        ),
        { code: "resource-limit" },
      );
      assert.deepEqual([refused.width, refused.height], [1, 1]);
      const accepted = new RenderTarget();
      const result = await session.previewTextPages(
        { target, text: "draft" },
        {
          pages: [{ pageIndex: 0, target: accepted }],
          zoom: 1,
          devicePixelRatio: 2,
          maxPixelsPerPage: 861,
        },
      );
      assert.deepEqual([accepted.width, accepted.height], [21, 41]);
      assert.equal(result.item?.paragraphs[0]?.text, "draft");
      assert.deepEqual(result.item?.pageSizes, test.controls.pageSizes);
      assert.deepEqual(test.painted, [{ text: "draft", target: accepted }]);
      assert.deepEqual(test.closed, [test.opened[1], test.opened[2]]);
      assert.deepEqual(test.events, []);
      assert.deepEqual(session.state, state);
    } finally {
      await test.viewer.destroy();
    }
  });

  it("cannot enlarge the runtime raster limit with an explicit draft budget", async () => {
    const test = setup({ ...defaultResourceLimits, maxDecodedPixels: 200 });
    try {
      await test.viewer.load(ORIGINAL, { fileName: "preview.docx" });
      const session = await test.viewer.edit();
      assert.ok(session.format === "docx");
      const target = (await session.getElements()).items[0]?.id;
      assert.ok(target);
      test.controls.pageSizes = [{ width: 20, height: 20 }];
      const canvas = new RenderTarget();
      await assert.rejects(
        session.previewTextPages(
          { target, text: "draft" },
          {
            pages: [{ pageIndex: 0, target: canvas }],
            zoom: 1,
            devicePixelRatio: 1,
            maxPixelsPerPage: 500,
          },
        ),
        { code: "resource-limit" },
      );
      assert.deepEqual([canvas.width, canvas.height], [1, 1]);
      assert.deepEqual(test.painted, []);
    } finally {
      await test.viewer.destroy();
    }
  });

  it("requires authoritative finite positive draft sizes only when an explicit raster budget is requested", async () => {
    const test = setup();
    try {
      await test.viewer.load(ORIGINAL, { fileName: "preview.docx" });
      const session = await test.viewer.edit();
      assert.ok(session.format === "docx");
      const target = (await session.getElements()).items[0]?.id;
      assert.ok(target);
      for (const size of [
        null,
        [],
        [{ width: NaN, height: 10 }],
        [{ width: 10, height: Infinity }],
        [{ width: 0, height: 10 }],
        [{ width: -1, height: 10 }],
        [{ width: Number.MAX_VALUE, height: 1 }],
      ]) {
        test.controls.pageSizes = size;
        const canvas = new RenderTarget();
        await assert.rejects(
          session.previewTextPages(
            { target, text: "draft" },
            {
              pages: [{ pageIndex: 0, target: canvas }],
              zoom: 2,
              maxPixelsPerPage: 100,
            },
          ),
          { code: "resource-limit" },
        );
        assert.deepEqual([canvas.width, canvas.height], [1, 1]);
      }
      assert.deepEqual(test.painted, []);
      test.controls.pageSizes = null;
      await session.previewTextPages(
        { target, text: "default" },
        { pages: [{ pageIndex: 0, target: new RenderTarget() }] },
      );
      assert.equal(test.painted.length, 1);
    } finally {
      await test.viewer.destroy();
    }
  });

  it("preserves typed invalid-page rejection and cleanup with an explicit draft budget", async () => {
    const test = setup();
    try {
      await test.viewer.load(ORIGINAL, { fileName: "preview.docx" });
      const session = await test.viewer.edit();
      assert.ok(session.format === "docx");
      const target = (await session.getElements()).items[0]?.id;
      assert.ok(target);
      test.events.length = 0;
      const state = session.state;
      const canvas = new RenderTarget();
      await assert.rejects(
        session.previewTextPages(
          { target, text: "draft" },
          {
            pages: [{ pageIndex: 1, target: canvas }],
            maxPixelsPerPage: 8_000_000,
          },
        ),
        { code: "render-failed", message: "Page index is out of range" },
      );
      assert.deepEqual([canvas.width, canvas.height], [1, 1]);
      assert.deepEqual(test.painted, []);
      assert.deepEqual(test.closed, [test.opened[1]]);
      assert.deepEqual(session.state, state);
      assert.deepEqual(test.events, []);
    } finally {
      await test.viewer.destroy();
    }
  });

  it("preserves empty split metadata and UTF-16 LF ranges, with per-page native layouts only where available", async () => {
    const test = setup();
    test.controls.pageCount = 2;
    test.controls.textMap = async (engine, pageIndex) => {
      const paragraphs = await engine.getElements(
        { kinds: ["paragraph"] },
        new AbortController().signal,
      );
      return paragraphs.flatMap((paragraph, index) =>
        paragraph.text && Math.floor(index / 2) === pageIndex
          ? [
              {
                text: paragraph.text,
                paragraphId: paragraph.id.slice(2),
                x: 72,
                y: 72 + (index % 2) * 20,
                width: paragraph.text.length * 8,
                height: 16,
                fontSize: 16,
                textLayer: "docx" as const,
              },
            ]
          : [],
      );
    };
    try {
      await test.viewer.load(ORIGINAL, { fileName: "preview.docx" });
      const session = await test.viewer.edit();
      assert.equal(session.format, "docx");
      if (session.format !== "docx") throw new Error("DOCX session required");
      const target = (await session.getElements()).items[0]?.id;
      assert.ok(target);
      const state = session.state;
      test.events.length = 0;
      const preview = await session.previewTextPages(
        {
          target,
          text: "\nX\n🙂\n",
          range: {
            start: { elementId: target, offset: 0 },
            end: { elementId: target, offset: 0 },
          },
          insertionStyle: { bold: true, italic: true, underline: true },
        },
        {
          pages: [0, 1].map((pageIndex) => ({
            pageIndex,
            target: new RenderTarget(),
          })),
        },
      );
      assert.ok(preview.item);
      assert.deepEqual(
        preview.item.paragraphs.map((paragraph) => ({
          text: paragraph.text,
          range: paragraph.draftRange,
          pages: paragraph.layouts.map((layout) => layout.pageIndex),
        })),
        [
          { text: "", range: { start: 0, end: 0 }, pages: [] },
          { text: "X", range: { start: 1, end: 2 }, pages: [0] },
          { text: "🙂", range: { start: 3, end: 5 }, pages: [1] },
          { text: "original", range: { start: 6, end: 14 }, pages: [1] },
        ],
      );
      assert.equal(preview.item.layout, undefined);
      assert.equal(
        preview.item.paragraphs[2]?.layouts[0]?.lines[0]?.range.end.offset,
        2,
      );
      const trailing = await session.previewTextPages(
        {
          target,
          text: "X\n",
          range: {
            start: { elementId: target, offset: 8 },
            end: { elementId: target, offset: 8 },
          },
        },
        { pages: [{ pageIndex: 0, target: new RenderTarget() }] },
      );
      assert.deepEqual(
        trailing.item?.paragraphs.map((paragraph) => ({
          text: paragraph.text,
          range: paragraph.draftRange,
          hasLayout: paragraph.layouts.length > 0,
        })),
        [
          { text: "originalX", range: { start: 0, end: 9 }, hasLayout: true },
          { text: "", range: { start: 10, end: 10 }, hasLayout: false },
        ],
      );
      assert.deepEqual(session.state, state);
      assert.deepEqual(test.events, []);
      assert.deepEqual((await session.save()).bytes, ORIGINAL);
    } finally {
      await test.viewer.destroy();
    }
  });

  it("aligns repeated paragraph text on a later requested page without painting preceding pages", async () => {
    const test = setup();
    test.controls.pageCount = 2;
    test.controls.textMap = async (engine, pageIndex) => {
      const paragraph = (
        await engine.getElements({}, new AbortController().signal)
      )[0];
      assert.ok(paragraph);
      return [
        {
          text: pageIndex === 0 ? "repeat " : "repeat",
          paragraphId: paragraph.id.slice(2),
          x: 72,
          y: 72,
          width: 56,
          height: 16,
          fontSize: 16,
          textLayer: "docx",
        },
      ];
    };
    try {
      await test.viewer.load(ORIGINAL, { fileName: "preview.docx" });
      const session = await test.viewer.edit();
      if (session.format !== "docx") throw new Error("DOCX session required");
      const target = (await session.getElements()).items[0]?.id;
      assert.ok(target);
      test.readPages.length = 0;
      const canvas = new RenderTarget();
      const preview = await session.previewTextPages(
        { target, text: "repeat repeat" },
        { pages: [{ pageIndex: 1, target: canvas }] },
      );
      assert.ok(preview.item);
      assert.equal(preview.item.layout?.pageIndex, 1);
      assert.equal(preview.item.layout?.lines[0]?.range.start.offset, 7);
      assert.equal(preview.item.paragraphs[0]?.layouts[0], preview.item.layout);
      assert.deepEqual(test.readPages, [0, 1]);
      assert.deepEqual(test.painted, [
        { text: "repeat repeat", target: canvas },
      ]);
      assert.deepEqual(
        preview.item.pages.map((page) => page.pageIndex),
        [1],
      );
      assert.equal("alignmentPages" in preview.item, false);
      const both = await session.previewTextPages(
        { target, text: "repeat repeat" },
        {
          pages: [0, 1].map((pageIndex) => ({
            pageIndex,
            target: new RenderTarget(),
          })),
        },
      );
      assert.deepEqual(
        both.item?.paragraphs[0]?.layouts.map((layout) => ({
          page: layout.pageIndex,
          start: layout.lines[0]?.range.start.offset,
          end: layout.lines[0]?.range.end.offset,
        })),
        [
          { page: 0, start: 0, end: 7 },
          { page: 1, start: 7, end: 13 },
        ],
      );
    } finally {
      await test.viewer.destroy();
    }
  });

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
        {
          target,
          text: "longer draft text",
          insertionStyle: { bold: true, italic: true },
        },
        {
          pages: [{ pageIndex: 0, target: canvas }],
          zoom: 1,
          devicePixelRatio: 1,
        },
      );
      assert.equal(preview.revision, state.revision);
      assert.equal(test.renderedStyle.current?.bold, true);
      assert.equal(test.renderedStyle.current?.italic, true);
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
