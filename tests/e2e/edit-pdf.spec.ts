import { expect, test, type Page } from "@playwright/test";

import { buildPdf } from "../../packages/viewer/test/fixtures/pdf-builder.js";
import { signedPdf } from "../../packages/viewer/test/fixtures/signed-pdf.js";

/*
 * PDF editing through the public API against the real adapter: PDF.js renders,
 * the PDFium worker edits. Fixtures are built with PDFium in Node and handed
 * to the page as plain arrays.
 */

const EDIT_ASSETS = [
  "/workers/pdf-edit-worker.js",
  "/assets/pdfium/pdfium.wasm",
];

test("deletes native PDF text and resumes typing in cleared targets after reopening [ACTION-954]", async ({
  page,
}) => {
  const neighbor = { text: "Neighbor", x: 72, y: 700 };
  const original = await buildPdf([
    {
      texts: [
        neighbor,
        { text: "SECOND LINEX", x: 200, y: 400, matrix: [0, 1, -1, 0] },
      ],
    },
  ]);
  const baseline = await buildPdf([{ texts: [neighbor] }]);
  const errors: string[] = [];
  page.on("pageerror", (error) => errors.push(error.message));
  await page.goto("/");
  const result = await page.evaluate(
    async ({ original, baseline }) => {
      const { ViewerClient } =
        (await import("/main.js")) as typeof import("../../packages/viewer/src/index.js");
      const host = document.createElement("div");
      Object.assign(host.style, { width: "800px", height: "900px" });
      document.body.replaceChildren(host);
      const viewer = ViewerClient.create({
        assetBaseUrl: new URL("/", location.href).href,
      }).createViewer({ container: host, initialZoom: 1 });
      const paint = async () => {
        const canvas = document.createElement("canvas");
        await viewer.renderPage(0, canvas, { zoom: 1, devicePixelRatio: 1 });
        return canvas.toDataURL();
      };
      try {
        await viewer.load(new Uint8Array(baseline), {
          fileName: "neighbor.pdf",
        });
        const neighborPaint = await paint();
        await viewer.load(new Uint8Array(original), {
          fileName: "deletion.pdf",
        });
        const session = await viewer.edit();
        if (session.format !== "pdf")
          throw new Error("Expected native PDF session");
        const target = "p0:o1";
        await session.setTextStyle({ target, style: { underline: true } });
        const inserted = await session.insertTextBox({
          pageIndex: 0,
          rect: { x: 100, y: 200, width: 240, height: 60 },
          text: "BOX",
          style: { underline: true },
        });
        const box = inserted.createdIds[0];
        if (!box) throw new Error("No inserted text box");
        await session.replaceText({
          target,
          text: "",
          range: {
            start: { elementId: target, offset: 11 },
            end: { elementId: target, offset: 12 },
          },
        });
        const shortened = (await session.getElement(target)).item?.text;
        await session.undo();
        const undone = (await session.getElement(target)).item?.text;
        await session.redo();
        const redone = (await session.getElement(target)).item?.text;
        await session.replaceText({ target, text: "" });
        await session.replaceText({
          target: box,
          text: "",
          range: {
            start: { elementId: box, offset: 0 },
            end: { elementId: box, offset: 3 },
          },
        });
        const clearedPaint = await paint();
        const saved = await session.save();
        await viewer.load(saved.bytes, { fileName: "cleared.pdf" });
        const reopened = await viewer.edit();
        if (reopened.format !== "pdf")
          throw new Error("Expected reopened PDF session");
        const empty = await Promise.all(
          [target, box].map(async (id) => (await reopened.getElement(id)).item),
        );
        const emptyText = await viewer.getPageText(0);
        const reopenedPaint = await paint();
        await reopened.replaceText({ target, text: "Resumed row" });
        await reopened.replaceText({ target: box, text: "Resumed box" });
        await reopened.undo();
        const undoneBox = (await reopened.getElement(box)).item?.text;
        await reopened.redo();
        const written = await reopened.save();
        await viewer.load(written.bytes, { fileName: "resumed.pdf" });
        return {
          shortened,
          undone,
          redone,
          empty,
          emptyText,
          cleared: clearedPaint === neighborPaint,
          roundtrip: reopenedPaint === neighborPaint,
          undoneBox,
          resumedText: await viewer.getPageText(0),
        };
      } finally {
        await viewer.close();
        host.remove();
      }
    },
    { original: Array.from(original), baseline: Array.from(baseline) },
  );
  expect(result.shortened).toBe("SECOND LINE");
  expect(result.undone).toBe("SECOND LINEX");
  expect(result.redone).toBe("SECOND LINE");
  expect(result.empty).toHaveLength(2);
  for (const element of result.empty) {
    expect(element).toMatchObject({ text: "", textStyle: { underline: true } });
    expect(element?.operations).toContain("replaceText");
  }
  expect(result.emptyText.trim()).toBe("Neighbor");
  expect(result.cleared).toBe(true);
  expect(result.roundtrip).toBe(true);
  expect(result.undoneBox).toBe("");
  expect(result.resumedText).toContain("Resumed row");
  expect(result.resumedText).toContain("Resumed box");
  expect(errors).toEqual([]);
});

test("exports true native bold italic and underline through the PDF worker", async ({
  page,
}) => {
  const original = await buildPdf([
    {
      texts: [
        { text: "Native title", x: 72, y: 740, fontSize: 18 },
        ...[
          "Native paragraphs preserve their identity",
          "across source lines and worker messages",
          "and remain editable after saving.",
        ].map((text, index) => ({
          text,
          x: 72,
          y: 650 - index * 20,
          fontSize: 11,
        })),
      ],
    },
  ]);
  const errors: string[] = [];
  page.on("pageerror", (error) => errors.push(error.message));
  await page.goto("/");
  const result = await page.evaluate(async (data) => {
    const { ViewerClient } =
      (await import("/main.js")) as typeof import("../../packages/viewer/src/index.js");
    const host = document.createElement("div");
    Object.assign(host.style, { width: "800px", height: "900px" });
    document.body.replaceChildren(host);
    const viewer = ViewerClient.create({
      assetBaseUrl: new URL("/", location.href).href,
    }).createViewer({ container: host, initialZoom: 1 });
    const paint = async () => {
      const canvas = document.createElement("canvas");
      await viewer.renderPage(0, canvas, { zoom: 1, devicePixelRatio: 1 });
      const context = canvas.getContext("2d");
      if (!context) throw new Error("No canvas context");
      const pixels = context.getImageData(
        0,
        0,
        canvas.width,
        canvas.height,
      ).data;
      let ink = 0;
      for (let index = 0; index < pixels.length; index += 4)
        if (
          pixels[index] < 180 &&
          pixels[index + 1] < 180 &&
          pixels[index + 2] < 180
        )
          ink += 1;
      return { image: canvas.toDataURL(), ink };
    };
    try {
      await viewer.load(new Uint8Array(data), { fileName: "styles.pdf" });
      const session = await viewer.edit();
      if (session.format !== "pdf") throw new Error("Expected PDF");
      const before = await paint();
      const paragraph = (await session.getTextParagraph("p0:o1")).item;
      if (!paragraph) throw new Error("Missing native paragraph");
      await session.setTextStyle({
        target: "p0:o0",
        style: { bold: true, italic: true, underline: true },
      });
      await session.setTextStyle({
        target: paragraph.id,
        style: { bold: true, underline: true },
      });
      const receipt = await session.insertTextBox({
        pageIndex: 0,
        rect: { x: 72, y: 260, width: 250, height: 70 },
        text: "Привіт світе",
        style: { bold: true, italic: true, underline: true },
      });
      const box = receipt.createdIds[0];
      if (!box) throw new Error("Missing inserted textbox");
      const ids = ["p0:o0", paragraph.id, box];
      const styles = await Promise.all(
        ids.map(async (id) => (await session.getElement(id)).item?.textStyle),
      );
      const styled = await paint();
      const saved = await session.save();
      await viewer.load(saved.bytes, { fileName: "saved-styles.pdf" });
      const reopened = await viewer.edit();
      if (reopened.format !== "pdf") throw new Error("Expected reopened PDF");
      const roundtrip = await paint();
      const reopenedStyles = await Promise.all(
        ids.map(async (id) => (await reopened.getElement(id)).item?.textStyle),
      );
      await reopened.replaceText({ target: box, text: "Привіт знову" });
      const typed = (await reopened.getElement(box)).item;
      await reopened.undo();
      const undo = await paint();
      await reopened.redo();
      const redo = (await reopened.getElement(box)).item;
      return {
        changed: before.image !== styled.image,
        ink: styled.ink,
        roundtrip: roundtrip.image === styled.image,
        undo: undo.image === styled.image,
        styles,
        reopenedStyles,
        typed,
        redo,
        warnings: receipt.warnings.map((warning) => warning.code),
      };
    } finally {
      await viewer.close();
      host.remove();
    }
  }, Array.from(original));
  expect(result.changed).toBe(true);
  expect(result.ink).toBeGreaterThan(1000);
  expect(result.roundtrip).toBe(true);
  expect(result.undo).toBe(true);
  expect(result.styles[0]).toMatchObject({
    bold: true,
    italic: true,
    underline: true,
  });
  expect(result.styles[1]).toMatchObject({ bold: true, underline: true });
  expect(result.styles[2]).toMatchObject({
    bold: true,
    italic: true,
    underline: true,
    fontFamily: "Liberation Sans",
  });
  expect(result.reopenedStyles).toEqual(result.styles);
  expect(result.typed).toMatchObject({
    text: "Привіт знову",
    textStyle: { bold: true, italic: true, underline: true },
  });
  expect(result.redo).toEqual(result.typed);
  expect(result.warnings).toContain("font-substitution");
  expect(errors).toEqual([]);
});

test("edits an imported paragraph through the native worker and reopens its stable target", async ({
  page,
}) => {
  const lines = [
    "Native paragraphs preserve their identity",
    "across source lines and worker messages",
    "and remain editable after saving.",
  ];
  const original = await buildPdf([
    {
      texts: lines.map((text, index) => ({
        text,
        x: 72,
        y: 700 - index * 20,
        fontSize: 11,
      })),
    },
  ]);
  await page.goto("/");
  const result = await page.evaluate(
    async ({ data, first }) => {
      const { ViewerClient } =
        (await import("/main.js")) as typeof import("../../packages/viewer/src/index.js");
      const viewer = ViewerClient.create({
        assetBaseUrl: new URL("/", location.href).href,
      }).createViewer();
      try {
        await viewer.load(new Uint8Array(data), { fileName: "paragraph.pdf" });
        const session = await viewer.edit();
        if (session.format !== "pdf") throw new Error("Expected a PDF session");
        const row = (await session.getElements({ pageIndex: 0 })).items.find(
          (item) => item.text === first,
        );
        if (!row?.textEditingTarget)
          throw new Error("Imported paragraph hint is missing");
        const paragraph = (await session.getTextParagraph(row.id)).item;
        if (!paragraph) throw new Error("Native paragraph is missing");
        const initial = (await session.getElement(row.textEditingTarget)).item;
        await session.replaceParagraphText({
          target: paragraph.id,
          text: "Updated paragraph through the worker.",
        });
        await session.undo();
        const undone = (await session.getTextParagraph(paragraph.id)).item
          ?.text;
        await session.redo();
        const saved = await session.save();
        await viewer.load(saved.bytes, { fileName: "saved-paragraph.pdf" });
        const reopened = await viewer.edit();
        if (reopened.format !== "pdf")
          throw new Error("Expected a reopened PDF session");
        const current = (await reopened.getTextParagraph(paragraph.id)).item;
        const elements = (await reopened.getElements({ pageIndex: 0 })).items;
        return {
          initialText: initial?.text,
          memberCount: paragraph.memberIds.length,
          undone,
          id: current?.id,
          originalId: paragraph.id,
          text: current?.text,
          elements: elements.map((element) => element.text),
        };
      } finally {
        await viewer.close();
      }
    },
    { data: Array.from(original), first: lines[0] },
  );
  expect(result).toEqual({
    initialText: lines.join(" "),
    memberCount: 3,
    undone: lines.join(" "),
    id: result.originalId,
    originalId: result.originalId,
    text: "Updated paragraph through the worker.",
    elements: ["Updated paragraph through the worker."],
  });
});

test("publishes initial and edited PDF rasters when text layers fail", async ({
  page,
}) => {
  const original = await buildPdf(["The PDF raster remains visible"]);
  const pageErrors: string[] = [];
  page.on("pageerror", (error) => pageErrors.push(error.message));
  await page.goto("/");
  const fixture = await page.evaluateHandle(async (data) => {
    const { ViewerClient } =
      (await import("/main.js")) as typeof import("../../packages/viewer/src/index.js");
    const container = document.createElement("div");
    Object.assign(container.style, {
      position: "fixed",
      inset: "0",
      overflow: "hidden",
    });
    document.body.append(container);
    const client = ViewerClient.create({
      assetBaseUrl: new URL("/", location.href).href,
    });
    const adapter = client.registry.resolve("pdf");
    const render = adapter.render.bind(adapter);
    const getTextMap = adapter.getTextMap?.bind(adapter);
    if (!getTextMap) throw new Error("Expected native PDF text extraction");
    let failText = Promise.withResolvers<void>();
    let held = Promise.withResolvers<void>();
    let textReads = 0;
    let holdRaster = true;
    let failBuilding = false;
    let targetCanvas: HTMLCanvasElement | OffscreenCanvas | undefined;
    let renderOutcome = "pending";
    adapter.getTextMap = async (handle, pageIndex, signal) => {
      textReads += 1;
      // Font preparation reads real native text; only the parallel viewport
      // extraction fails, while its real PDF.js raster is explicitly held.
      if (holdRaster && !failBuilding && textReads === 2) {
        await failText.promise;
        throw new Error("Text layer extraction failed");
      }
      return getTextMap(handle, pageIndex, signal);
    };
    adapter.render = async (handle, target, viewport, signal) => {
      targetCanvas = target;
      try {
        await render(handle, target, viewport, signal);
        renderOutcome = "completed";
      } catch (error) {
        renderOutcome = "failed";
        throw error;
      }
    };
    const dpr = window.devicePixelRatio || 1;
    const requestFrame = window.requestAnimationFrame.bind(window);
    const cancelFrame = window.cancelAnimationFrame.bind(window);
    const append = Element.prototype.append;
    Element.prototype.append = function (...nodes) {
      append.apply(this, nodes);
      if (
        failBuilding &&
        !this.isConnected &&
        nodes.some(
          (node) =>
            node instanceof HTMLSpanElement && node.dataset.pageIndex === "0",
        )
      ) {
        failBuilding = false;
        throw new Error("Text layer construction failed");
      }
    };
    const frames = new Map<number, FrameRequestCallback>();
    let frameId = 0;
    let zoom = 1;
    window.requestAnimationFrame = (callback) => {
      if (!holdRaster || targetCanvas?.width !== Math.ceil(612 * zoom * dpr))
        return requestFrame(callback);
      const id = --frameId;
      frames.set(id, callback);
      held.resolve();
      return id;
    };
    window.cancelAnimationFrame = (id) => {
      if (!frames.delete(id)) cancelFrame(id);
    };
    const resumeRaster = () => {
      holdRaster = false;
      for (const callback of frames.values()) requestFrame(callback);
      frames.clear();
    };
    const viewer = client.createViewer({
      container,
      ui: false,
      initialZoom: zoom,
    });
    await viewer.load(new Uint8Array(data), { fileName: "text-failure.pdf" });
    return {
      dpr,
      async failTextLayer() {
        await held.promise;
        failText.resolve();
      },
      resumeRaster,
      async waitForRaster() {
        await held.promise;
      },
      recover() {
        zoom = 1.25;
        viewer.setZoom(zoom);
      },
      async highlight() {
        await viewer.search("raster");
      },
      async edit() {
        const session = await viewer.edit();
        if (session.format !== "pdf") throw new Error("Expected PDF editing");
        failText = Promise.withResolvers<void>();
        held = Promise.withResolvers<void>();
        textReads = 0;
        targetCanvas = undefined;
        renderOutcome = "pending";
        holdRaster = true;
        failBuilding = true;
        await session.apply([
          {
            op: "replaceText",
            target: "p0:o0",
            text: "The edited PDF raster is blue",
          },
          { op: "setTextStyle", target: "p0:o0", style: { color: "#0000ff" } },
        ]);
      },
      snapshot() {
        const slot = container.querySelector<HTMLElement>(
          '[data-page-index="0"]',
        );
        const canvas = slot?.querySelector("canvas");
        const pixels = canvas
          ?.getContext("2d")
          ?.getImageData(0, 0, canvas.width, canvas.height).data;
        let ink = 0;
        let blue = 0;
        if (pixels)
          for (let at = 0; at < pixels.length; at += 4) {
            if ((pixels[at + 3] ?? 0) > 0 && (pixels[at] ?? 255) < 128)
              ink += 1;
            if (
              (pixels[at + 3] ?? 0) > 0 &&
              (pixels[at] ?? 255) < 128 &&
              (pixels[at + 2] ?? 0) > 200
            )
              blue += 1;
          }
        return {
          renderOutcome,
          ink,
          blue,
          width: canvas?.width,
          height: canvas?.height,
          error: slot?.dataset.renderError ?? "",
          text:
            slot?.querySelector('[data-zrimo-layer="text"]')?.textContent ?? "",
          highlights:
            slot?.querySelector('[data-zrimo-layer="highlight"]')
              ?.childElementCount ?? 0,
        };
      },
      async close() {
        failText.resolve();
        resumeRaster();
        window.requestAnimationFrame = requestFrame;
        window.cancelAnimationFrame = cancelFrame;
        Element.prototype.append = append;
        await viewer.destroy();
        await client.destroy();
        container.remove();
      },
    };
  }, Array.from(original));
  try {
    const dpr = await fixture.evaluate((f) => f.dpr);
    await fixture.evaluate((f) => f.failTextLayer());
    await expect
      .poll(() => fixture.evaluate((f) => f.snapshot().error))
      .toBe("Text layer extraction failed");
    await fixture.evaluate((f) => f.resumeRaster());
    await expect
      .poll(() => fixture.evaluate((f) => f.snapshot()))
      .toMatchObject({
        renderOutcome: "completed",
        width: Math.ceil(612 * dpr),
        height: Math.ceil(792 * dpr),
        error: "Text layer extraction failed",
        text: "",
        highlights: 0,
      });
    await expect
      .poll(() => fixture.evaluate((f) => f.snapshot().ink))
      .toBeGreaterThan(50 * dpr * dpr);
    await fixture.evaluate((f) => f.recover());
    await expect
      .poll(() => fixture.evaluate((f) => f.snapshot()))
      .toMatchObject({
        text: "The PDF raster remains visible",
        error: "",
      });
    await fixture.evaluate((f) => f.highlight());
    await expect
      .poll(() => fixture.evaluate((f) => f.snapshot().highlights))
      .toBeGreaterThan(0);
    await fixture.evaluate((f) => f.edit());
    await fixture.evaluate((f) => f.waitForRaster());
    await fixture.evaluate((f) => f.resumeRaster());
    await expect
      .poll(() => fixture.evaluate((f) => f.snapshot()))
      .toMatchObject({
        renderOutcome: "completed",
        width: Math.ceil(612 * 1.25 * dpr),
        height: Math.ceil(792 * 1.25 * dpr),
        error: "Text layer construction failed",
        text: "",
        highlights: 0,
      });
    await expect
      .poll(() => fixture.evaluate((f) => f.snapshot().blue))
      .toBeGreaterThan(50 * dpr * dpr);
  } finally {
    await fixture.evaluate((f) => f.close());
    await fixture.dispose();
  }
  expect(pageErrors).toEqual([]);
});

test("repaints an unchanged mounted page when an edit retires its pending zoom render", async ({
  page,
}) => {
  const original = await buildPdf([
    "Original first page",
    "Unchanged second page",
  ]);
  const pageErrors: string[] = [];
  page.on("pageerror", (error) => pageErrors.push(error.message));
  await page.goto("/");
  const fixture = await page.evaluateHandle(async (data) => {
    const { ViewerClient } =
      (await import("/main.js")) as typeof import("../../packages/viewer/src/index.js");
    const container = document.createElement("div");
    Object.assign(container.style, {
      position: "fixed",
      inset: "0",
      height: "1800px",
      overflow: "hidden",
    });
    document.body.append(container);
    const client = ViewerClient.create({
      assetBaseUrl: new URL("/", location.href).href,
    });
    const adapter = client.registry.resolve("pdf");
    const render = adapter.render.bind(adapter);
    let secondPaint: HTMLCanvasElement | OffscreenCanvas | undefined;
    adapter.render = (handle, target, viewport, signal) => {
      if (viewport.pageIndex === 1 && viewport.zoom === 1.5)
        secondPaint = target;
      return render(handle, target, viewport, signal);
    };
    const viewer = client.createViewer({
      container,
      ui: false,
      initialZoom: 1,
    });
    const requestFrame = window.requestAnimationFrame.bind(window);
    const cancelFrame = window.cancelAnimationFrame.bind(window);
    const frames = new Map<number, FrameRequestCallback>();
    const dpr = Math.max(1, window.devicePixelRatio || 1);
    let frameId = 0;
    const slot = (index: number) =>
      container.querySelector<HTMLElement>(`[data-page-index="${index}"]`);
    const snapshot = (index: number) => {
      const root = slot(index);
      const canvas = root?.querySelector("canvas");
      const pixels = canvas
        ?.getContext("2d")
        ?.getImageData(0, 0, canvas.width, canvas.height).data;
      let ink = 0;
      if (pixels)
        for (let at = 0; at < pixels.length; at += 4)
          if ((pixels[at + 3] ?? 0) > 0 && (pixels[at] ?? 255) < 128) ink += 1;
      return {
        width: canvas?.width,
        height: canvas?.height,
        ink,
        error: root?.dataset.renderError ?? "",
        text:
          root?.querySelector('[data-zrimo-layer="text"]')?.textContent ?? "",
      };
    };
    const restore = () => {
      window.requestAnimationFrame = requestFrame;
      window.cancelAnimationFrame = cancelFrame;
      for (const callback of frames.values()) requestFrame(callback);
      frames.clear();
    };
    await viewer.load(new Uint8Array(data), {
      fileName: "mounted-retirement.pdf",
    });
    const session = await viewer.edit();
    if (session.format !== "pdf") throw new Error("Expected PDF editing");
    const target = (await session.getElements({ pageIndex: 0 })).items[0];
    if (!target) throw new Error("Missing first-page text");
    return {
      dpr,
      snapshot,
      async editDuringZoom() {
        const held = Promise.withResolvers<void>();
        window.requestAnimationFrame = (callback) => {
          // Hold the native PDF.js continuation only once page two has begun
          // painting its new-size canvas. Earlier viewport frames run normally.
          if (secondPaint?.width !== Math.ceil(918 * dpr))
            return requestFrame(callback);
          const id = --frameId;
          frames.set(id, callback);
          held.resolve();
          return id;
        };
        window.cancelAnimationFrame = (id) => {
          if (!frames.delete(id)) cancelFrame(id);
        };
        viewer.setZoom(1.5);
        await held.promise;
        await session.replaceText({
          target: target.id,
          text: "Updated first page",
        });
        restore();
      },
      async close() {
        restore();
        await viewer.destroy();
        container.remove();
      },
    };
  }, Array.from(original));
  try {
    const dpr = await fixture.evaluate((f) => f.dpr);
    await expect
      .poll(() => fixture.evaluate((f) => f.snapshot(1)))
      .toMatchObject({
        width: Math.ceil(612 * dpr),
        height: Math.ceil(792 * dpr),
        error: "",
        text: "Unchanged second page",
      });
    await expect
      .poll(() => fixture.evaluate((f) => f.snapshot(1).ink))
      .toBeGreaterThan(50 * dpr * dpr);
    await fixture.evaluate((f) => f.editDuringZoom());
    await expect
      .poll(() => fixture.evaluate((f) => f.snapshot(1)))
      .toMatchObject({
        width: Math.ceil(918 * dpr),
        height: Math.ceil(1188 * dpr),
        error: "",
        text: "Unchanged second page",
      });
    await expect
      .poll(() => fixture.evaluate((f) => f.snapshot(1).ink))
      .toBeGreaterThan(100 * dpr * dpr);
    await expect
      .poll(() => fixture.evaluate((f) => f.snapshot(0).text))
      .toBe("Updated first page");
  } finally {
    await fixture.evaluate((f) => f.close());
    await fixture.dispose();
  }
  await expect.poll(() => page.workers().length).toBe(0);
  expect(pageErrors).toEqual([]);
});

test("aborts real PDF renders retired during getPage and before queued execution", async ({
  page,
}) => {
  const original = await buildPdf(["Transport retirement"]);
  await page.goto("/");
  const result = await page.evaluate(async (data) => {
    const { ViewerClient } =
      (await import("/main.js")) as typeof import("../../packages/viewer/src/index.js");
    const client = ViewerClient.create({ limits: { maxConcurrentRenders: 1 } });
    const adapter = client.registry.resolve("pdf");
    const context = {
      format: "pdf" as const,
      signal: new AbortController().signal,
      limits: client.limits,
      assetBaseUrl: new URL("/", location.href),
      reportProgress() {},
      reportWarning() {},
    };
    const viewport = { pageIndex: 0, zoom: 1, devicePixelRatio: 1 };
    const outcome = (promise: Promise<void>) =>
      promise.then(
        () => "completed",
        (error: unknown) =>
          error instanceof Error && "code" in error
            ? String(error.code)
            : String(error),
      );
    const handle = await adapter.open(new Uint8Array(data), context);
    // Opening the adapter directly leaves getPage uncached: render must cross
    // the real PDF.js transport while close destroys that same transport.
    const inGetPage = outcome(
      adapter.render(handle, document.createElement("canvas"), viewport),
    );
    await adapter.close(handle);
    const other = await adapter.open(new Uint8Array(data), context);
    const scheduler = client.renderScheduler;
    const gate = Promise.withResolvers<void>();
    const controller = new AbortController();
    const active = scheduler.run(
      "visible",
      controller.signal,
      () => gate.promise,
    );
    const queued = outcome(
      scheduler.run("visible", controller.signal, () =>
        adapter.render(other, document.createElement("canvas"), viewport),
      ),
    );
    await adapter.close(other);
    gate.resolve();
    await active;
    const result = { inGetPage: await inGetPage, queued: await queued };
    await client.destroy();
    return result;
  }, Array.from(original));
  expect(result).toEqual({ inGetPage: "aborted", queued: "aborted" });
  await expect.poll(() => page.workers().length).toBe(0);
});

test("retires an actively rendering PDF after an edit and releases its workers", async ({
  page,
}) => {
  const original = await buildPdf(["Original native text"]);
  const pageErrors: string[] = [];
  page.on("pageerror", (error) => pageErrors.push(error.message));
  await page.goto("/");
  const result = await page.evaluate(async (data) => {
    const { ViewerClient } =
      (await import("/main.js")) as typeof import("../../packages/viewer/src/index.js");
    const viewer = ViewerClient.create({
      assetBaseUrl: new URL("/", location.href).href,
    }).createViewer();
    const requestFrame = window.requestAnimationFrame.bind(window);
    const cancelFrame = window.cancelAnimationFrame.bind(window);
    const heldFrames = new Map<number, FrameRequestCallback>();
    const paused = Promise.withResolvers<void>();
    const failed = Promise.withResolvers<{ readonly error: string }>();
    const controller = new AbortController();
    const canvas = document.createElement("canvas");
    const onError = (event: ErrorEvent) =>
      failed.resolve({ error: event.message });
    window.addEventListener("error", onError);
    let frameId = 0;
    const restoreFrames = () => {
      window.requestAnimationFrame = requestFrame;
      window.cancelAnimationFrame = cancelFrame;
      for (const callback of heldFrames.values()) requestFrame(callback);
      heldFrames.clear();
    };
    try {
      await viewer.load(new Uint8Array(data), { fileName: "cleanup.pdf" });
      const session = await viewer.edit();
      if (session.format !== "pdf") throw new Error("Expected PDF editing");
      const target = (await session.getElements({ pageIndex: 0 })).items[0];
      if (!target) throw new Error("Fixture text is missing");
      // Drain the viewer's own notification frame before holding the native
      // PDF.js render. No viewport is mounted in this headless public viewer.
      await new Promise<void>((resolve) => requestFrame(() => resolve()));
      window.requestAnimationFrame = (callback) => {
        const id = --frameId;
        heldFrames.set(id, callback);
        // PDF.js sizes the target before registering its active render task.
        if (canvas.width === 612) paused.resolve();
        return id;
      };
      window.cancelAnimationFrame = (id) => {
        if (!heldFrames.delete(id)) cancelFrame(id);
      };
      const rendering = viewer
        .renderPage(0, canvas, {
          signal: controller.signal,
          devicePixelRatio: 1,
        })
        .then(
          () => ({ render: "completed" }),
          (error: unknown) => ({
            render:
              error instanceof Error && "code" in error
                ? String(error.code)
                : String(error),
          }),
        );
      await paused.promise;
      // Showing the edited bytes retires the original PDF while its render is
      // still suspended, exactly as a quick formatting change can in the UI.
      await session.replaceText({
        target: target.id,
        text: "Updated native text",
      });
      const retired = await Promise.race([rendering, failed.promise]);
      restoreFrames();
      const updated = document.createElement("canvas");
      await viewer.renderPage(0, updated);
      const text = await viewer.getPageText(0);
      return {
        retired,
        text,
        painted: updated.width > 0 && updated.height > 0,
      };
    } finally {
      controller.abort();
      restoreFrames();
      window.removeEventListener("error", onError);
      await viewer.close();
    }
  }, Array.from(original));
  expect(result).toEqual({
    retired: { render: "aborted" },
    text: "Updated native text",
    painted: true,
  });
  expect(pageErrors).toEqual([]);
  await expect.poll(() => page.workers().length).toBe(0);
});

test("releases the owned PDF worker when an untouched document closes", async ({
  page,
}) => {
  const original = await buildPdf(["Close without editing"]);
  await page.goto("/");
  await page.evaluate(async (data) => {
    const { ViewerClient } =
      (await import("/main.js")) as typeof import("../../packages/viewer/src/index.js");
    const viewer = ViewerClient.create({
      assetBaseUrl: new URL("/", location.href).href,
    }).createViewer();
    try {
      await viewer.load(new Uint8Array(data), { fileName: "close.pdf" });
      await viewer.getPageText(0);
    } finally {
      await viewer.close();
    }
  }, Array.from(original));
  await expect.poll(() => page.workers().length).toBe(0);
});

async function loadPdf(page: Page, bytes: Uint8Array): Promise<void> {
  await page.goto("/");
  await page.evaluate(async (data) => {
    const { ViewerClient } = (await import("/main.js")) as {
      ViewerClient: { create(config: unknown): { createViewer(): unknown } };
    };
    const client = ViewerClient.create({
      assetBaseUrl: new URL("/", location.href),
    });
    const viewer = client.createViewer() as {
      load(bytes: Uint8Array, options: unknown): Promise<void>;
    };
    await viewer.load(new Uint8Array(data), { fileName: "fixture.pdf" });
    (window as unknown as { __pdfViewer: unknown }).__pdfViewer = viewer;
  }, Array.from(bytes));
}

test("starts the PDFium worker only on edit() and saves the untouched original", async ({
  page,
}) => {
  const original = await buildPdf(["Alpha", "Beta", "Gamma"]);
  const requests: string[] = [];
  page.on("request", (request) =>
    requests.push(new URL(request.url()).pathname),
  );
  await loadPdf(page, original);

  const info = await page.evaluate(() => {
    const viewer = (window as unknown as { __pdfViewer: any }).__pdfViewer;
    return {
      editing: viewer.getDocumentInfo().capabilities.editing,
      pageCount: viewer.state.pageCount,
    };
  });
  expect(info).toEqual({ editing: true, pageCount: 3 });
  expect(requests.filter((path) => EDIT_ASSETS.includes(path))).toEqual([]);

  const result = await page.evaluate(async (expected) => {
    const viewer = (window as unknown as { __pdfViewer: any }).__pdfViewer;
    const states: boolean[] = [];
    viewer.on("editstatechange", (state: { active: boolean }) =>
      states.push(state.active),
    );
    const session = await viewer.edit();
    const again = await viewer.edit();
    const { bytes: saved }: { bytes: Uint8Array } = await session.save();
    const identical =
      saved.length === expected.length &&
      saved.every((byte: number, index: number) => byte === expected[index]);
    await viewer.close();
    return {
      format: session.format,
      shared: session === again,
      revision: session.state.revision,
      dirty: session.state.dirty,
      identical,
      states,
      schemaFormat: session.schemas.format,
    };
  }, Array.from(original));
  expect(result).toEqual({
    format: "pdf",
    shared: true,
    revision: 0,
    dirty: false,
    identical: true,
    states: [true, false],
    schemaFormat: "pdf",
  });
  for (const asset of EDIT_ASSETS) expect(requests).toContain(asset);
});

test("inserts and edits a text box, re-renders, saves and reloads it", async ({
  page,
}) => {
  const original = await buildPdf(["Existing"]);
  await loadPdf(page, original);
  const result = await page.evaluate(async () => {
    const viewer = (window as unknown as { __pdfViewer: any }).__pdfViewer;
    const darkPixels = async () => {
      const canvas = document.createElement("canvas");
      await viewer.renderPage(0, canvas, { zoom: 1, devicePixelRatio: 1 });
      const pixels = canvas
        .getContext("2d")!
        .getImageData(0, 0, canvas.width, canvas.height).data;
      let count = 0;
      for (let offset = 0; offset < pixels.length; offset += 4)
        if (pixels[offset]! < 128) count += 1;
      return count;
    };
    const before = await darkPixels();
    const session = await viewer.edit();
    const inserted = await session.insertTextBox({
      pageIndex: 0,
      rect: { x: 72, y: 200, width: 450, height: 100 },
      text: "Hello from web-doc editing",
      style: { fontSize: 24, bold: true },
    });
    const afterInsert = await darkPixels();
    const textAfterInsert: string = await viewer.getPageText(0);
    const id: string = inserted.createdIds[0];
    await session.replaceText({ target: id, text: "Edited text box" });
    const textAfterEdit: string = await viewer.getPageText(0);
    await session.undo();
    const textAfterUndo: string = await viewer.getPageText(0);
    await session.redo();
    const { bytes: saved }: { bytes: Uint8Array } = await session.save();

    const { ViewerClient } = (await import("/main.js")) as any;
    const fresh = ViewerClient.create({
      assetBaseUrl: new URL("/", location.href),
    }).createViewer();
    await fresh.load(saved, { fileName: "saved.pdf" });
    const freshText: string = await fresh.getPageText(0);
    const freshSession = await fresh.edit();
    const { items: elements } = await freshSession.getElements({
      pageIndex: 0,
    });
    return {
      revision: session.state.revision,
      darker: afterInsert > before,
      textAfterInsert,
      textAfterEdit,
      textAfterUndo,
      freshText,
      elements: elements.map(
        (element: { kind: string; id: string; text?: string }) => [
          element.kind,
          element.id,
          element.text,
        ],
      ),
    };
  });
  expect(result.darker).toBe(true);
  expect(result.revision).toBe(4);
  // PDF.js joins text runs without separators, so check the words.
  expect(result.textAfterInsert).toContain("Hello from web-doc");
  expect(result.textAfterInsert).toContain("editing");
  expect(result.textAfterEdit).toContain("Edited text box");
  expect(result.textAfterEdit).not.toContain("Hello from");
  expect(result.textAfterUndo).toContain("Hello from web-doc");
  expect(result.freshText).toContain("Edited text box");
  expect(result.elements).toEqual([
    ["text", "p0:o0", "Existing"],
    ["textBox", "p0:n1.0.0", "Edited text box"],
  ]);
});

test("moves and deletes elements through the session", async ({ page }) => {
  const original = await buildPdf([
    {
      texts: [{ text: "Anchor", x: 72, y: 700 }],
      image: { x: 300, y: 500, width: 160, height: 80 },
    },
  ]);
  await loadPdf(page, original);
  const result = await page.evaluate(async () => {
    const viewer = (window as unknown as { __pdfViewer: any }).__pdfViewer;
    const session = await viewer.edit();
    const { item: before } = await session.getElement("p0:o1");
    await session.moveElement({ target: "p0:o1", by: { dx: 20, dy: 30 } });
    const { item: moved } = await session.getElement("p0:o1");
    const { items: hits } = await session.findText("Anchor");
    await session.deleteElement({ target: "p0:o0" });
    const { items: remaining } = await session.getElements({ pageIndex: 0 });
    const text: string = await viewer.getPageText(0);
    return {
      before: before.bounds,
      moved: moved.bounds,
      hitPage: hits[0]?.pageIndex,
      remaining: remaining.map((element: { id: string }) => element.id),
      text,
    };
  });
  expect(result.moved.x).toBeCloseTo(result.before.x + 20, 0);
  expect(result.moved.y).toBeCloseTo(result.before.y + 30, 0);
  expect(result.hitPage).toBe(0);
  expect(result.remaining).toEqual(["p0:o1"]);
  expect(result.text).not.toContain("Anchor");
});

test("changes the page structure and the viewer follows", async ({ page }) => {
  const original = await buildPdf(["One", "Two"]);
  await loadPdf(page, original);
  const result = await page.evaluate(async () => {
    const viewer = (window as unknown as { __pdfViewer: any }).__pdfViewer;
    const session = await viewer.edit();
    await session.insertPage({ index: 1, size: { width: 300, height: 200 } });
    const afterInsert = {
      count: viewer.state.pageCount,
      sizes: viewer.getDocumentInfo().pageSizes,
    };
    await session.rotatePage({ pageIndex: 0, rotation: 90 });
    const afterRotate = viewer.getDocumentInfo().pageSizes[0];
    await session.movePage({ from: 2, to: 0 });
    const firstText: string = await viewer.getPageText(0);
    await session.deletePage({ pageIndex: 1 });
    return {
      afterInsert,
      afterRotate,
      firstText,
      finalCount: viewer.state.pageCount,
      revision: session.state.revision,
    };
  });
  expect(result.afterInsert.count).toBe(3);
  expect(result.afterInsert.sizes[1]).toEqual({ width: 300, height: 200 });
  expect(result.afterRotate).toEqual({ width: 792, height: 612 });
  expect(result.firstText).toBe("Two");
  expect(result.finalCount).toBe(2);
  expect(result.revision).toBe(4);
});

test("embeds the fallback font for Cyrillic text and fetches it only then", async ({
  page,
}) => {
  const original = await buildPdf(["Latin"]);
  const requests: string[] = [];
  page.on("request", (request) =>
    requests.push(new URL(request.url()).pathname),
  );
  await loadPdf(page, original);
  const result = await page.evaluate(async () => {
    const viewer = (window as unknown as { __pdfViewer: any }).__pdfViewer;
    const session = await viewer.edit();
    await session.insertTextBox({
      pageIndex: 0,
      rect: { x: 72, y: 200, width: 450, height: 60 },
      text: "Only Latin here",
    });
    const fontRequestsBefore = performance
      .getEntriesByType("resource")
      .filter((entry) => entry.name.endsWith(".ttf")).length;
    const receipt = await session.insertTextBox({
      pageIndex: 0,
      rect: { x: 72, y: 300, width: 450, height: 60 },
      text: "Привіт, світе!",
    });
    const text: string = await viewer.getPageText(0);
    return {
      fontRequestsBefore,
      warning: receipt.warnings[0]?.code,
      text,
    };
  });
  expect(result.fontRequestsBefore).toBe(0);
  expect(result.warning).toBe("font-substitution");
  expect(result.text).toContain("Привіт, світе!");
  expect(requests).toContain("/fonts/noto-sans-latin-cyrillic.ttf");
});

test("edits text that already exists in the file", async ({ page }) => {
  const original = await buildPdf(["Existing"]);
  await loadPdf(page, original);
  const result = await page.evaluate(async () => {
    const viewer = (window as unknown as { __pdfViewer: any }).__pdfViewer;
    const session = await viewer.edit();
    const inPlace = await session.replaceText({
      target: "p0:o0",
      text: "Rewritten",
    });
    const afterReplace: string = await viewer.getPageText(0);
    await session.setTextStyle({
      target: "p0:o0",
      style: { color: "#ff0000", fontSize: 30 },
    });
    const { item: styled } = await session.getElement("p0:o0");
    return {
      warnings: inPlace.warnings.length,
      afterReplace,
      style: styled.textStyle,
    };
  });
  expect(result.warnings).toBe(0);
  expect(result.afterReplace).toContain("Rewritten");
  expect(result.style).toEqual({
    fontFamily: "Helvetica",
    fontSize: 30,
    bold: false,
    italic: false,
    color: "#ff0000",
  });
});

test("draws shapes that render on the page", async ({ page }) => {
  const original = await buildPdf([{ width: 300, height: 300 }]);
  await loadPdf(page, original);
  const result = await page.evaluate(async () => {
    const viewer = (window as unknown as { __pdfViewer: any }).__pdfViewer;
    const darkPixels = async () => {
      const canvas = document.createElement("canvas");
      await viewer.renderPage(0, canvas, { zoom: 1, devicePixelRatio: 1 });
      const pixels = canvas
        .getContext("2d")!
        .getImageData(0, 0, canvas.width, canvas.height).data;
      let count = 0;
      for (let offset = 0; offset < pixels.length; offset += 4)
        if (pixels[offset]! < 64) count += 1;
      return count;
    };
    const before = await darkPixels();
    const session = await viewer.edit();
    const receipt = await session.insertShape({
      pageIndex: 0,
      shape: "rectangle",
      rect: { x: 50, y: 50, width: 100, height: 100 },
      fill: { color: "#000000" },
    });
    const after = await darkPixels();
    const { item: element } = await session.getElement(receipt.createdIds[0]);
    return { before, after, kind: element.kind, bounds: element.bounds };
  });
  expect(result.before).toBe(0);
  expect(result.after).toBeGreaterThan(9000);
  expect(result.kind).toBe("shape");
  expect(result.bounds).toEqual({ x: 50, y: 50, width: 100, height: 100 });
});

test("inserts a PNG decoded in the worker", async ({ page }) => {
  const original = await buildPdf([{ width: 300, height: 300 }]);
  await loadPdf(page, original);
  const result = await page.evaluate(async () => {
    const viewer = (window as unknown as { __pdfViewer: any }).__pdfViewer;
    const source = document.createElement("canvas");
    source.width = 40;
    source.height = 20;
    const context = source.getContext("2d")!;
    context.fillStyle = "#000000";
    context.fillRect(0, 0, 40, 20);
    const blob: Blob = await new Promise((resolve) =>
      source.toBlob((value) => resolve(value!), "image/png"),
    );
    const data = new Uint8Array(await blob.arrayBuffer());
    const darkPixels = async () => {
      const canvas = document.createElement("canvas");
      await viewer.renderPage(0, canvas, { zoom: 1, devicePixelRatio: 1 });
      const pixels = canvas
        .getContext("2d")!
        .getImageData(0, 0, canvas.width, canvas.height).data;
      let count = 0;
      for (let offset = 0; offset < pixels.length; offset += 4)
        if (pixels[offset]! < 64) count += 1;
      return count;
    };
    const before = await darkPixels();
    const session = await viewer.edit();
    const receipt = await session.insertImage({
      pageIndex: 0,
      rect: { x: 50, y: 50, width: 120, height: 60 },
      data,
      mimeType: "image/png",
    });
    const after = await darkPixels();
    const { item: element } = await session.getElement(receipt.createdIds[0]);
    return { before, after, kind: element.kind, bounds: element.bounds };
  });
  expect(result.before).toBe(0);
  expect(result.after).toBeGreaterThan(6500);
  expect(result.kind).toBe("image");
  expect(result.bounds).toEqual({ x: 50, y: 50, width: 120, height: 60 });
});

test("inserts a table and edits a cell", async ({ page }) => {
  const original = await buildPdf([{ width: 400, height: 400 }]);
  await loadPdf(page, original);
  const result = await page.evaluate(async () => {
    const viewer = (window as unknown as { __pdfViewer: any }).__pdfViewer;
    const session = await viewer.edit();
    const receipt = await session.insertTable({
      pageIndex: 0,
      at: { x: 40, y: 40 },
      width: 300,
      rows: [
        ["Item", "Total"],
        ["Apples", "3.60"],
      ],
      style: { headerFill: "#dddddd" },
    });
    const id = receipt.createdIds[0];
    await session.setTableCell({ target: id, row: 1, column: 1, text: "4.80" });
    const { item: element } = await session.getElement(id);
    const { items: found } = await session.findText("4.80");
    const canvas = document.createElement("canvas");
    await viewer.renderPage(0, canvas, { zoom: 1, devicePixelRatio: 1 });
    const pixels = canvas
      .getContext("2d")!
      .getImageData(0, 0, canvas.width, canvas.height).data;
    // Thin grid lines render grey, so any ink counts; the page was blank.
    let dark = 0;
    for (let offset = 0; offset < pixels.length; offset += 4)
      if (pixels[offset]! < 200) dark += 1;
    return {
      kind: element.kind,
      rows: element.table.rows,
      operations: element.operations,
      bounds: element.bounds,
      matches: found.length,
      dark,
      revision: session.state.revision,
    };
  });
  expect(result.kind).toBe("table");
  expect(result.rows).toEqual([
    ["Item", "Total"],
    ["Apples", "4.80"],
  ]);
  expect(result.operations).toEqual([
    "setTableCell",
    "moveElement",
    "deleteElement",
  ]);
  // PDFium grows a stroked path's bounds by the stroke width on each side.
  expect(Math.abs(result.bounds.x - 40)).toBeLessThan(1.6);
  expect(Math.abs(result.bounds.y - 40)).toBeLessThan(1.6);
  expect(Math.abs(result.bounds.width - 300)).toBeLessThan(1.6);
  expect(result.matches).toBe(1);
  expect(result.dark).toBeGreaterThan(500);
  expect(result.revision).toBe(2);
});

test("warns once when a signed PDF is edited and keeps its bytes", async ({
  page,
}) => {
  const original = signedPdf();
  await loadPdf(page, original);
  const result = await page.evaluate(async () => {
    const viewer = (window as unknown as { __pdfViewer: any }).__pdfViewer;
    const session = await viewer.edit();
    const first = await session.insertTextBox({
      pageIndex: 0,
      rect: { x: 72, y: 72, width: 200, height: 40 },
      text: "Added",
    });
    const second = await session.replaceText({
      target: first.createdIds[0],
      text: "Changed",
    });
    const { bytes: saved }: { bytes: Uint8Array } = await session.save();
    return {
      firstCodes: first.warnings.map(
        (warning: { code: string }) => warning.code,
      ),
      secondCodes: second.warnings.map(
        (warning: { code: string }) => warning.code,
      ),
      saved: Array.from(saved.subarray(0, 400)),
      length: saved.length,
    };
  });
  expect(result.firstCodes).toEqual(["fidelity-degraded"]);
  expect(result.secondCodes).toEqual([]);
  expect(result.length).toBeGreaterThan(original.length);
  expect(result.saved.slice(0, original.length)).toEqual(
    Array.from(original.subarray(0, 400)),
  );
});

test("applies one operation on a ten-page PDF within three seconds", async ({
  page,
}) => {
  const original = await buildPdf(
    Array.from({ length: 10 }, (_, index) => `Page ${index + 1}`),
  );
  await loadPdf(page, original);
  const elapsed = await page.evaluate(async () => {
    const viewer = (window as unknown as { __pdfViewer: any }).__pdfViewer;
    const session = await viewer.edit();
    const started = performance.now();
    await session.insertTextBox({
      pageIndex: 4,
      rect: { x: 72, y: 72, width: 200, height: 40 },
      text: "Timed",
    });
    return performance.now() - started;
  });
  test.info().annotations.push({
    type: "apply-ms",
    description: elapsed.toFixed(0),
  });
  console.log(`apply latency 10 pages: ${elapsed.toFixed(0)} ms`);
  expect(elapsed).toBeLessThan(3000);
});

/*
 * Latency of one apply() — engine, incremental save, PDF.js reopen — on
 * larger files. The first apply on a document pays for PDF.js parsing it
 * again from scratch; the second shows the steady state. Numbers go into the
 * PDF spec; only a loose ceiling is asserted so the run stays stable on slow
 * machines.
 */
for (const pages of [100, 500]) {
  test(`measures apply() latency on a ${pages}-page PDF`, async ({ page }) => {
    test.setTimeout(120_000);
    const original = await buildPdf(
      Array.from({ length: pages }, (_, index) => `Page ${index + 1}`),
    );
    await loadPdf(page, original);
    const timings = await page.evaluate(async () => {
      const viewer = (window as unknown as { __pdfViewer: any }).__pdfViewer;
      const session = await viewer.edit();
      const time = async (index: number) => {
        const started = performance.now();
        await session.insertTextBox({
          pageIndex: index,
          rect: { x: 72, y: 72, width: 200, height: 40 },
          text: `Timed ${index}`,
        });
        return Math.round(performance.now() - started);
      };
      const first = await time(0);
      const second = await time(1);
      return { first, second, bytes: (await session.save()).bytes.length };
    });
    test.info().annotations.push({
      type: `apply-ms-${pages}`,
      description: `first ${timings.first}, second ${timings.second}, saved ${timings.bytes} bytes`,
    });
    console.log(
      `apply latency ${pages} pages: first ${timings.first} ms, second ${timings.second} ms`,
    );
    expect(timings.second).toBeLessThan(10_000);
  });
}

for (const transition of [
  "unchanged",
  "changed",
  "resized",
  "shifted",
  "unmounted",
  "replaced",
] as const) {
  test(`keeps a queued paint notification truthful when its page is ${transition}`, async ({
    page,
  }) => {
    const bytes = await buildPdf(["Original page A", "Original page B"]);
    const errors: string[] = [];
    page.on("pageerror", (error) => errors.push(error.message));
    await page.goto("/");
    const fixture = await page.evaluateHandle(
      async ({ data, transition }) => {
        const { ViewerClient } =
          (await import("/main.js")) as typeof import("../../packages/viewer/src/index.js");
        const container = document.createElement("div");
        Object.assign(container.style, {
          position: "fixed",
          inset: "0",
          height: "1800px",
          overflow: "hidden",
        });
        document.body.append(container);
        const client = ViewerClient.create({
          assetBaseUrl: new URL("/", location.href).href,
        });
        const adapter = client.registry.resolve("pdf");
        const nativeRender = adapter.render.bind(adapter);
        const renders = [0, 0];
        adapter.render = (handle, target, viewport, signal) => {
          renders[viewport.pageIndex] = (renders[viewport.pageIndex] ?? 0) + 1;
          return nativeRender(handle, target, viewport, signal);
        };
        const viewer = client.createViewer({
          container,
          ui: false,
          initialZoom: 1,
          layout: transition === "unmounted" ? "single" : "continuous",
        });
        const snapshotPage = (pageIndex: number) => {
          const slot = container.querySelector<HTMLElement>(
            `[data-page-index="${pageIndex}"]`,
          );
          const canvas = slot?.querySelector("canvas");
          const pixels = canvas
            ?.getContext("2d")
            ?.getImageData(0, 0, canvas.width, canvas.height).data;
          let hash = 2166136261;
          let ink = 0;
          if (pixels)
            for (let at = 0; at < pixels.length; at += 4) {
              hash = Math.imul(hash ^ (pixels[at] ?? 0), 16777619) >>> 0;
              if ((pixels[at + 3] ?? 0) > 0 && (pixels[at] ?? 255) < 128)
                ink += 1;
            }
          return {
            mounted: Boolean(slot),
            width: canvas?.width,
            height: canvas?.height,
            hash,
            ink,
            text:
              slot?.querySelector('[data-zrimo-layer="text"]')?.textContent ??
              "",
            error: slot?.dataset.renderError ?? "",
          };
        };
        const layouts: {
          sessionId: string;
          revision: number;
          pages: readonly number[];
          a: ReturnType<typeof snapshotPage>;
        }[] = [];
        viewer.on("layoutchange", (event) =>
          layouts.push({
            sessionId: event.sessionId,
            revision: event.revision,
            pages: [...event.pages],
            a: snapshotPage(0),
          }),
        );
        const requestFrame = window.requestAnimationFrame.bind(window);
        const cancelFrame = window.cancelAnimationFrame.bind(window);
        const frames = new Map<number, FrameRequestCallback>();
        let frameId = 0;
        let armed = false;
        let originalHash = 0;
        const releaseFrames = () => {
          armed = false;
          window.requestAnimationFrame = requestFrame;
          window.cancelAnimationFrame = cancelFrame;
          for (const callback of frames.values()) requestFrame(callback);
          frames.clear();
        };
        await viewer.load(new Uint8Array(data), {
          fileName: "late-paint-notification.pdf",
        });
        const session = await viewer.edit();
        if (session.format !== "pdf")
          throw new Error("Expected native PDF editing");
        const a = (await session.getElements({ pageIndex: 0 })).items[0];
        const b = (await session.getElements({ pageIndex: 1 })).items[0];
        if (!a || !b)
          throw new Error("Both native page text objects are required");
        return {
          dpr: Math.max(1, window.devicePixelRatio || 1),
          async editA() {
            originalHash = snapshotPage(0).hash;
            armed = true;
            window.requestAnimationFrame = (callback) => {
              const pageA = snapshotPage(0);
              // Hold the notification frame only after real native raster and
              // matching text have published. Raster operations remain native.
              if (
                armed &&
                pageA.text === "Updated page A" &&
                pageA.hash !== originalHash
              ) {
                const id = --frameId;
                frames.set(id, callback);
                return id;
              }
              return requestFrame(callback);
            };
            window.cancelAnimationFrame = (id) => {
              if (!frames.delete(id)) cancelFrame(id);
            };
            return session.replaceText({
              target: a.id,
              text: "Updated page A",
            });
          },
          async transition() {
            switch (transition) {
              case "unchanged":
                return session.replaceText({
                  target: b.id,
                  text: "Updated page B",
                });
              case "changed":
                return session.replaceText({
                  target: a.id,
                  text: "Final page A",
                });
              case "resized":
                return session.rotatePage({ pageIndex: 0, by: 90 });
              case "shifted":
                return session.deletePage({ pageIndex: 0 });
              case "replaced":
                await viewer.load(new Uint8Array(data), {
                  fileName: "replacement.pdf",
                });
                await viewer.edit();
                return undefined;
              case "unmounted": {
                viewer.goToPage(1);
                // The first held callback is the completed paint notification.
                // Run the later navigation frame first to retire its actual slot.
                const later = [...frames.entries()].slice(1);
                for (const [id] of later) frames.delete(id);
                await new Promise<void>((resolve) =>
                  requestFrame((time) => {
                    for (const [, callback] of later) callback(time);
                    resolve();
                  }),
                );
                if (snapshotPage(0).mounted)
                  throw new Error("Navigation did not unmount page A");
                return undefined;
              }
            }
          },
          releaseFrames,
          snapshot() {
            return {
              a: snapshotPage(0),
              b: snapshotPage(1),
              layouts,
              renders: [...renders],
              held: frames.size,
            };
          },
          settle: () =>
            new Promise<void>((resolve) =>
              requestFrame(() => requestFrame(() => resolve())),
            ),
          async close() {
            releaseFrames();
            await viewer.destroy();
            await client.destroy();
            container.remove();
          },
        };
      },
      { data: Array.from(bytes), transition },
    );
    try {
      await expect
        .poll(() => fixture.evaluate((f) => f.snapshot().a))
        .toMatchObject({ text: "Original page A", error: "" });
      if (transition !== "unmounted")
        await expect
          .poll(() => fixture.evaluate((f) => f.snapshot().b.text))
          .toBe("Original page B");
      const before = await fixture.evaluate((f) => f.snapshot());
      const receiptA = await fixture.evaluate((f) => f.editA());
      await expect
        .poll(() => fixture.evaluate((f) => f.snapshot().held))
        .toBeGreaterThan(0);
      const publishedA = await fixture.evaluate((f) => f.snapshot());
      expect(publishedA.a.hash).not.toBe(before.a.hash);
      expect(publishedA.a.ink).toBeGreaterThan(50);
      expect(publishedA.a.text).toBe("Updated page A");
      expect(
        publishedA.layouts.some(
          (event) =>
            event.revision >= receiptA.revision && event.pages.includes(0),
        ),
      ).toBe(false);
      const receiptB = await fixture.evaluate((f) => f.transition());
      await fixture.evaluate((f) => f.releaseFrames());
      const targetPage =
        transition === "unchanged" || transition === "unmounted" ? "b" : "a";
      const text =
        transition === "changed"
          ? "Final page A"
          : transition === "shifted" || transition === "unmounted"
            ? "Original page B"
            : transition === "replaced"
              ? "Original page A"
              : transition === "unchanged"
                ? "Updated page B"
                : "Updated page A";
      await expect
        .poll(() =>
          fixture.evaluate(
            (f, targetPage) => f.snapshot()[targetPage].text,
            targetPage,
          ),
        )
        .toBe(text);
      // Rotation retains the text, so its old text layer is not a readiness
      // signal. Wait for the requested page's actual public paint event.
      await expect
        .poll(() =>
          fixture.evaluate(
            (f, expected) =>
              f
                .snapshot()
                .layouts.some(
                  (event) =>
                    event.revision === expected.revision &&
                    event.pages.includes(expected.pageIndex),
                ),
            {
              revision:
                receiptB?.revision ??
                (transition === "unmounted" ? receiptA.revision : 0),
              pageIndex: targetPage === "a" ? 0 : 1,
            },
          ),
        )
        .toBe(true);
      if (transition === "unchanged") {
        await expect
          .poll(() =>
            fixture.evaluate(
              (f, revision) =>
                f
                  .snapshot()
                  .layouts.some(
                    (event) =>
                      event.revision >= revision && event.pages.includes(0),
                  ),
              receiptA.revision,
            ),
          )
          .toBe(true);
        const after = await fixture.evaluate((f) => f.snapshot());
        expect(after.a.hash).toBe(publishedA.a.hash);
        expect(after.renders[0]).toBe(publishedA.renders[0]);
      }
      await fixture.evaluate((f) => f.settle());
      const result = await fixture.evaluate((f) => f.snapshot());
      const announcements = result.layouts.filter(
        (event) =>
          event.pages.includes(0) && event.revision >= receiptA.revision,
      );
      if (transition === "unmounted" || transition === "replaced") {
        expect(announcements).toEqual([]);
      } else {
        expect(receiptB?.revision).toBe(receiptA.revision + 1);
        expect(announcements).toHaveLength(1);
        expect(announcements[0]?.revision).toBe(receiptB?.revision);
        const expectedText =
          transition === "unchanged" || transition === "resized"
            ? "Updated page A"
            : text;
        expect(announcements[0]?.a.text).toBe(expectedText);
        expect(announcements[0]?.a.error).toBe("");
        expect(announcements[0]?.a.ink).toBeGreaterThan(50);
        if (transition === "changed" || transition === "shifted") {
          expect(announcements[0]?.a.hash).not.toBe(publishedA.a.hash);
        } else if (transition === "unchanged") {
          expect(announcements[0]?.a.hash).toBe(publishedA.a.hash);
        }
        if (transition === "resized") {
          const dpr = await fixture.evaluate((f) => f.dpr);
          expect(announcements[0]?.a).toMatchObject({
            width: Math.ceil(792 * dpr),
            height: Math.ceil(612 * dpr),
          });
        }
      }
    } finally {
      await fixture.evaluate((f) => f.close());
      await fixture.dispose();
    }
    expect(errors).toEqual([]);
  });
}
