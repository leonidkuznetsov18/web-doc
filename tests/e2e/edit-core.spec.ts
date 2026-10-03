import { expect, test, type Page } from "@playwright/test";

/*
 * The editing core against a fake editable format: the adapter paints each
 * page in a colour derived from its text, and the engine edits the page list.
 * Everything a real format does differently lives behind the same contracts.
 */

interface SetupOptions {
  readonly ui: boolean;
  readonly pages: readonly string[];
  readonly zoom?: number;
}

declare global {
  interface Window {
    __editTest?: {
      readonly viewer: {
        readonly state: {
          readonly zoom: number;
          readonly fit: string;
          readonly pageCount: number;
          readonly pageIndex: number;
          readonly panY: number;
        };
        load(bytes: Uint8Array, options: { fileName: string }): Promise<void>;
        getEditSession(): unknown;
        setZoom(zoom: number): void;
        panBy(x: number, y: number): void;
        search(query: string): Promise<unknown>;
        getDocumentInfo(): { readonly pageCount: number };
        edit(): Promise<{
          apply(operations: unknown[]): Promise<{ revision: number }>;
          undo(): Promise<unknown>;
          redo(): Promise<unknown>;
        }>;
        pageToClient(
          pageIndex: number,
          rect: { x: number; y: number; width: number; height: number },
        ):
          | { left: number; top: number; width: number; height: number }
          | undefined;
        clientToPage(
          x: number,
          y: number,
        ): { pageIndex: number; point: { x: number; y: number } } | undefined;
      };
      readonly viewportRoot: HTMLElement;
      /** Paints per page since setup. */
      readonly renders: number[];
      /** Resolves with the next layoutchange event. */
      nextLayout(): Promise<{ revision: number; pages: number[] }>;
      pixel(pageIndex: number): string | undefined;
      holdPaint(text: string): void;
      holdFrames(): void;
      releaseFrames(): void;
      releasePaint(text: string): void;
      failPaint(text: string): void;
      paintStarted(text: string): boolean;
      paintFinished(text: string): boolean;
      expectedPixel(text: string): string;
      loadDocument(pages: readonly string[]): Promise<void>;
      canvasRect(pageIndex: number): DOMRect | undefined;
      highlights(pageIndex: number): number;
      settle(): Promise<void>;
    };
  }
}

async function setup(page: Page, options: SetupOptions): Promise<void> {
  await page.goto("/");
  await page.evaluate(async (options) => {
    const { ViewerClient } = (await import("/main.js")) as {
      ViewerClient: {
        create(config: unknown): {
          createViewer(
            config: unknown,
          ): Window["__editTest"] extends infer T
            ? T extends { viewer: infer V }
              ? V
              : never
            : never;
        };
      };
    };
    const SIGNATURE = "%PDF-1.7\n";
    const encode = (pages: readonly string[]) =>
      new TextEncoder().encode(SIGNATURE + JSON.stringify(pages));
    const decode = (bytes: Uint8Array) =>
      JSON.parse(
        new TextDecoder().decode(bytes).slice(SIGNATURE.length),
      ) as string[];
    const colour = (text: string) => {
      let hash = 7;
      for (const character of text)
        hash = (hash * 31 + character.charCodeAt(0)) >>> 0;
      return `hsl(${hash % 360} 70% 60%)`;
    };

    interface PaintGate {
      started: boolean;
      finished: boolean;
      pending: Promise<void>;
      release(): void;
      fail(): void;
    }
    const requestFrame = window.requestAnimationFrame.bind(window);
    const cancelFrame = window.cancelAnimationFrame.bind(window);
    const heldFrames = new Map<number, FrameRequestCallback>();
    const replayedFrames = new Map<number, number>();
    let framesHeld = false;
    window.requestAnimationFrame = (callback) => {
      const id = requestFrame((time) => {
        if (framesHeld) heldFrames.set(id, callback);
        else callback(time);
      });
      return id;
    };
    window.cancelAnimationFrame = (id) => {
      heldFrames.delete(id);
      cancelFrame(replayedFrames.get(id) ?? id);
      replayedFrames.delete(id);
    };
    const releaseFrames = () => {
      framesHeld = false;
      for (const [id, callback] of heldFrames) {
        const replayed = requestFrame((time) => {
          replayedFrames.delete(id);
          callback(time);
        });
        replayedFrames.set(id, replayed);
      }
      heldFrames.clear();
    };
    const paintGates = new Map<string, PaintGate>();
    const renders: number[] = [];
    const adapter = {
      id: "edit-fixture",
      formats: ["pdf"],
      async open(data: Uint8Array) {
        return { pages: decode(data) };
      },
      async getInfo(handle: { pages: string[] }) {
        return {
          format: "pdf",
          unit: "page",
          pageCount: handle.pages.length,
          pageSizes: handle.pages.map(() => ({ width: 300, height: 400 })),
        };
      },
      async render(
        handle: { pages: string[] },
        target: HTMLCanvasElement,
        viewport: { pageIndex: number; zoom: number; devicePixelRatio: number },
      ) {
        const text = handle.pages[viewport.pageIndex]!;
        renders[viewport.pageIndex] = (renders[viewport.pageIndex] ?? 0) + 1;
        const dpr = viewport.devicePixelRatio;
        target.width = Math.ceil(300 * viewport.zoom * dpr);
        target.height = Math.ceil(400 * viewport.zoom * dpr);
        target.style.width = `${300 * viewport.zoom}px`;
        target.style.height = `${400 * viewport.zoom}px`;
        // PDF.js also resets the target bitmap before its asynchronous paint.
        // Hold that boundary explicitly; no wall-clock delay is involved.
        const gate = paintGates.get(text);
        if (gate) {
          gate.started = true;
          try {
            await gate.pending;
          } finally {
            gate.finished = true;
          }
        }
        const context = target.getContext("2d")!;
        context.fillStyle = colour(text);
        context.fillRect(0, 0, target.width, target.height);
      },
      async getTextMap(handle: { pages: string[] }, pageIndex: number) {
        return [
          {
            text: handle.pages[pageIndex]!,
            x: 20,
            y: 20,
            width: 200,
            height: 20,
            font: "16px sans-serif",
            fontSize: 16,
            logicalStart: 0,
            logicalEnd: handle.pages[pageIndex]!.length,
          },
        ];
      },
      close() {},
      edit: {
        formats: ["pdf"],
        // The core already implements the session; a real format wraps it.
        createSession: (core: unknown) => core,
        async load(original: Uint8Array) {
          const base = decode(original);
          let pages = [...base];
          const apply = (operations: readonly Record<string, unknown>[]) => {
            const changed = new Set<number>();
            for (const operation of operations) {
              if (operation.op === "setText") {
                pages[operation.pageIndex as number] = operation.text as string;
                changed.add(operation.pageIndex as number);
              } else if (operation.op === "insertPage") {
                pages.splice(
                  operation.index as number,
                  0,
                  operation.text as string,
                );
                for (
                  let index = operation.index as number;
                  index < pages.length;
                  index += 1
                )
                  changed.add(index);
              } else if (operation.op === "deletePage") {
                pages.splice(operation.pageIndex as number, 1);
                for (
                  let index = operation.pageIndex as number;
                  index < pages.length;
                  index += 1
                )
                  changed.add(index);
              }
            }
            return {
              createdIds: [],
              changedPages: [...changed].sort((a, b) => a - b),
              pageCount: pages.length,
              warnings: [],
            };
          };
          const operation = (
            name: string,
            fields: Record<string, unknown>,
          ) => ({
            type: "object",
            required: ["op", ...Object.keys(fields)],
            additionalProperties: false,
            properties: { op: { const: name }, ...fields },
          });
          return {
            schemas: {
              format: "pdf",
              version: 1,
              operations: {
                setText: operation("setText", {
                  pageIndex: { type: "integer", minimum: 0 },
                  text: { type: "string" },
                }),
                insertPage: operation("insertPage", {
                  index: { type: "integer", minimum: 0 },
                  text: { type: "string" },
                }),
                deletePage: operation("deletePage", {
                  pageIndex: { type: "integer", minimum: 0 },
                }),
              },
            },
            async validate() {
              return [];
            },
            // Batches arrive with the state id the core assigned them.
            async apply(batch: {
              readonly operations: readonly Record<string, unknown>[];
            }) {
              return apply(batch.operations);
            },
            async materialize() {
              return encode(pages);
            },
            async restore(target: {
              readonly batches: readonly {
                readonly operations: readonly Record<string, unknown>[];
              }[];
            }) {
              pages = [...base];
              for (const batch of target.batches) apply(batch.operations);
            },
            async getElements() {
              return [];
            },
            async elementsAt() {
              return [];
            },
            async findText() {
              return [];
            },
            async dispose() {},
          };
        },
      },
    };

    const container = document.createElement("div");
    Object.assign(container.style, {
      position: "fixed",
      left: "0",
      top: "0",
      width: "640px",
      height: "480px",
      overflow: "hidden",
    });
    document.body.append(container);
    const client = ViewerClient.create({ adapters: [adapter] });
    const viewer = client.createViewer({
      container,
      ui: options.ui,
      locale: "en",
      ...(options.zoom === undefined ? {} : { initialZoom: options.zoom }),
    });
    await (
      viewer as unknown as {
        load(bytes: Uint8Array, o: unknown): Promise<void>;
      }
    ).load(encode(options.pages), { fileName: "fixture.pdf" });
    const viewportRoot = container.querySelector<HTMLElement>(
      '[data-zrimo="viewport"]',
    )!;
    const slot = (pageIndex: number) =>
      viewportRoot.querySelector<HTMLElement>(
        `[data-page-index="${pageIndex}"]`,
      );
    window.__editTest = {
      viewer,
      viewportRoot,
      renders,
      nextLayout: () =>
        new Promise((resolve) => {
          const off = (
            viewer as unknown as {
              on(type: string, listener: (event: unknown) => void): () => void;
            }
          ).on("layoutchange", (event) => {
            off();
            resolve(event as { revision: number; pages: number[] });
          });
        }),
      pixel(pageIndex) {
        const canvas = slot(pageIndex)?.querySelector("canvas");
        if (!canvas || canvas.width === 0) return undefined;
        const data = canvas
          .getContext("2d")!
          .getImageData(
            Math.floor(canvas.width / 2),
            Math.floor(canvas.height / 2),
            1,
            1,
          ).data;
        return `${data[0]},${data[1]},${data[2]}`;
      },
      holdFrames() {
        framesHeld = true;
      },
      releaseFrames,
      holdPaint(text) {
        const gate: PaintGate = {
          started: false,
          finished: false,
          pending: Promise.resolve(),
          release() {
            throw new Error("Paint gate was not initialized");
          },
          fail() {
            throw new Error("Paint gate was not initialized");
          },
        };
        gate.pending = new Promise<void>((resolve, reject) => {
          gate.release = resolve;
          gate.fail = () => reject(new Error("Fixture repaint failed"));
        });
        paintGates.set(text, gate);
      },
      releasePaint(text) {
        const gate = paintGates.get(text);
        if (!gate) throw new Error(`No held paint for ${text}`);
        gate.release();
      },
      failPaint(text) {
        const gate = paintGates.get(text);
        if (!gate) throw new Error(`No held paint for ${text}`);
        gate.fail();
      },
      paintStarted: (text) => paintGates.get(text)?.started ?? false,
      paintFinished: (text) => paintGates.get(text)?.finished ?? false,
      expectedPixel(text) {
        const canvas = document.createElement("canvas");
        const context = canvas.getContext("2d");
        if (!context) throw new Error("Fixture canvas context is unavailable");
        context.fillStyle = colour(text);
        context.fillRect(0, 0, 1, 1);
        const [red, green, blue] = context.getImageData(0, 0, 1, 1).data;
        return `${red},${green},${blue}`;
      },
      loadDocument: (pages) =>
        viewer.load(encode(pages), { fileName: "replacement.pdf" }),
      canvasRect(pageIndex) {
        return slot(pageIndex)
          ?.querySelector("canvas")
          ?.getBoundingClientRect();
      },
      highlights(pageIndex) {
        return (
          slot(pageIndex)?.querySelector('[data-zrimo-layer="highlight"]')
            ?.childElementCount ?? 0
        );
      },
      settle() {
        return new Promise((resolve) =>
          requestAnimationFrame(() => requestAnimationFrame(() => resolve())),
        );
      },
    };
  }, options);
  await page.waitForFunction((text) => {
    const fixture = window.__editTest;
    return fixture && fixture.pixel(0) === fixture.expectedPixel(text);
  }, options.pages[0] ?? "");
}

test("re-renders changed pages and keeps zoom, fit and scroll", async ({
  page,
}) => {
  await setup(page, {
    ui: false,
    pages: ["alpha", "beta", "gamma", "delta", "epsilon"],
  });
  await page.evaluate(async () => {
    const t = window.__editTest!;
    t.viewer.setZoom(1.5);
    await t.settle();
    // Searching reveals its match, which scrolls; pan only afterwards.
    await t.viewer.search("alpha");
  });
  await expect
    .poll(() => page.evaluate(() => window.__editTest!.highlights(0)))
    .toBeGreaterThan(0);
  const before = await page.evaluate(async () => {
    const t = window.__editTest!;
    t.viewer.panBy(0, 150);
    await t.settle();
    return {
      pixel: t.pixel(0),
      zoom: t.viewer.state.zoom,
      scrollTop: t.viewportRoot.scrollTop,
    };
  });
  expect(before.scrollTop).toBe(150);

  const layout = await page.evaluate(async () => {
    const t = window.__editTest!;
    const rendersBefore = [...t.renders];
    const session = await t.viewer.edit();
    const next = t.nextLayout();
    const receipt = await session.apply([
      { op: "setText", pageIndex: 0, text: "ALPHA edited" },
    ]);
    const event = await next;
    const canvas = t.canvasRect(0)!;
    const client = t.viewer.pageToClient(0, {
      x: 0,
      y: 0,
      width: 300,
      height: 400,
    })!;
    return {
      event,
      receipt,
      rendersBefore,
      rendersAfter: [...t.renders],
      geometryExact:
        Math.abs(client.left - canvas.left) < 1 &&
        Math.abs(client.top - canvas.top) < 1 &&
        Math.abs(client.width - canvas.width) < 1,
    };
  });
  // layoutchange names the applied revision and only the changed page repainted.
  expect(layout.event.revision).toBe(layout.receipt.revision);
  expect(layout.event.pages).toContain(0);
  expect(layout.geometryExact).toBe(true);
  expect(layout.rendersAfter[0]).toBe((layout.rendersBefore[0] ?? 0) + 1);
  expect(layout.rendersAfter[1] ?? 0).toBe(layout.rendersBefore[1] ?? 0);
  await expect
    .poll(() => page.evaluate(() => window.__editTest!.pixel(0)))
    .not.toBe(before.pixel);
  const after = await page.evaluate(() => {
    const t = window.__editTest!;
    return {
      zoom: t.viewer.state.zoom,
      fit: t.viewer.state.fit,
      scrollTop: t.viewportRoot.scrollTop,
      highlights: t.highlights(0),
      pageCount: t.viewer.state.pageCount,
    };
  });
  expect(after).toEqual({
    zoom: 1.5,
    fit: "none",
    scrollTop: 150,
    highlights: 0,
    pageCount: 5,
  });

  await page.evaluate(async () => {
    const session = await window.__editTest!.viewer.edit();
    await session.undo();
  });
  await expect
    .poll(() => page.evaluate(() => window.__editTest!.pixel(0)))
    .toBe(before.pixel);
  await page.evaluate(async () => {
    const session = await window.__editTest!.viewer.edit();
    await session.redo();
  });
  await expect
    .poll(() => page.evaluate(() => window.__editTest!.pixel(0)))
    .not.toBe(before.pixel);
});

test("a changed page count updates the layout and the built-in page counter", async ({
  page,
}) => {
  await setup(page, { ui: true, pages: ["one", "two", "three"] });
  await expect(page.getByText("of 3")).toBeVisible();
  const heightBefore = await page.evaluate(
    () =>
      window.__editTest!.viewportRoot.firstElementChild!.getBoundingClientRect()
        .height,
  );
  await page.evaluate(async () => {
    const session = await window.__editTest!.viewer.edit();
    await session.apply([{ op: "insertPage", index: 1, text: "inserted" }]);
  });
  await expect(page.getByText("of 4")).toBeVisible();
  await expect
    .poll(() =>
      page.evaluate(
        () =>
          window.__editTest!.viewportRoot.firstElementChild!.getBoundingClientRect()
            .height,
      ),
    )
    .toBeGreaterThan(heightBefore);
  expect(
    await page.evaluate(
      () => window.__editTest!.viewer.getDocumentInfo().pageCount,
    ),
  ).toBe(4);
});

test("view-geometry helpers agree with the painted canvas across zoom and scroll", async ({
  page,
}) => {
  await setup(page, {
    ui: false,
    pages: ["one", "two", "three", "four", "five"],
  });
  const rect = { x: 30, y: 40, width: 50, height: 60 };
  for (const zoom of [0.5, 1, 2]) {
    for (const scroll of [0, 120]) {
      const result = await page.evaluate(
        async ({ zoom, scroll, rect }) => {
          const t = window.__editTest!;
          t.viewer.setZoom(zoom);
          await t.settle();
          t.viewportRoot.scrollTop = scroll;
          await t.settle();
          const canvas = t.canvasRect(0)!;
          const client = t.viewer.pageToClient(0, rect)!;
          const back = t.viewer.clientToPage(client.left + 1, client.top + 1)!;
          const gap = t.viewer.clientToPage(
            canvas.left + 10,
            canvas.bottom + 5,
          );
          return {
            expectedLeft: canvas.left + rect.x * zoom,
            expectedTop: canvas.top + rect.y * zoom,
            expectedWidth: rect.width * zoom,
            client,
            back,
            gap,
            unmounted: t.viewer.pageToClient(4, rect),
          };
        },
        { zoom, scroll, rect },
      );
      expect(
        Math.abs(result.client.left - result.expectedLeft),
      ).toBeLessThanOrEqual(1);
      expect(
        Math.abs(result.client.top - result.expectedTop),
      ).toBeLessThanOrEqual(1);
      expect(
        Math.abs(result.client.width - result.expectedWidth),
      ).toBeLessThanOrEqual(1);
      expect(result.back.pageIndex).toBe(0);
      expect(
        Math.abs(result.back.point.x - rect.x - 1 / zoom),
      ).toBeLessThanOrEqual(1 / zoom);
      expect(
        Math.abs(result.back.point.y - rect.y - 1 / zoom),
      ).toBeLessThanOrEqual(1 / zoom);
      expect(result.gap).toBeUndefined();
      expect(result.unmounted).toBeUndefined();
    }
  }
});

test.describe("edit paint continuity", () => {
  test("keeps the painted page visible until an edited frame is ready", async ({
    page,
  }) => {
    await setup(page, { ui: false, pages: ["alpha", "beta"] });
    const before = await page.evaluate(() => {
      const t = window.__editTest;
      if (!t) throw new Error("Edit fixture is unavailable");
      t.holdPaint("alpha edited");
      return t.pixel(0);
    });
    await page.evaluate(async () => {
      const t = window.__editTest;
      if (!t) throw new Error("Edit fixture is unavailable");
      const session = await t.viewer.edit();
      await session.apply([
        { op: "setText", pageIndex: 0, text: "alpha edited" },
      ]);
    });
    await page.waitForFunction(() =>
      window.__editTest?.paintStarted("alpha edited"),
    );
    const textLayer = page.locator(
      '[data-page-index="0"] [data-zrimo-layer="text"]',
    );
    const whilePreparing = {
      pixel: await page.evaluate(() => window.__editTest?.pixel(0)),
      text: await textLayer.textContent(),
    };
    await page.evaluate(() => window.__editTest?.releasePaint("alpha edited"));
    await expect
      .poll(() => page.evaluate(() => window.__editTest?.pixel(0)))
      .toBe(
        await page.evaluate(() =>
          window.__editTest?.expectedPixel("alpha edited"),
        ),
      );
    await expect(textLayer).toHaveText("alpha edited");
    expect(
      whilePreparing,
      "The last raster and matching text must stay visible until the new frame is ready",
    ).toEqual({ pixel: before, text: "alpha" });
  });

  test("a late older paint cannot overwrite the newest edited frame", async ({
    page,
  }) => {
    await setup(page, { ui: false, pages: ["alpha"] });
    await page.evaluate(async () => {
      const t = window.__editTest;
      if (!t) throw new Error("Edit fixture is unavailable");
      t.holdPaint("older edit");
      t.holdPaint("newest edit");
      const session = await t.viewer.edit();
      await session.apply([
        { op: "setText", pageIndex: 0, text: "older edit" },
      ]);
    });
    await page.waitForFunction(() =>
      window.__editTest?.paintStarted("older edit"),
    );
    await page.evaluate(async () => {
      const t = window.__editTest;
      if (!t) throw new Error("Edit fixture is unavailable");
      const session = await t.viewer.edit();
      await session.apply([
        { op: "setText", pageIndex: 0, text: "newest edit" },
      ]);
    });
    await page.waitForFunction(() =>
      window.__editTest?.paintStarted("newest edit"),
    );
    await page.evaluate(() => window.__editTest?.releasePaint("newest edit"));
    const expected = await page.evaluate(() =>
      window.__editTest?.expectedPixel("newest edit"),
    );
    await expect
      .poll(() => page.evaluate(() => window.__editTest?.pixel(0)))
      .toBe(expected);
    // A backend may complete work that was already dispatched before abort.
    await page.evaluate(() => window.__editTest?.releasePaint("older edit"));
    await page.waitForFunction(() =>
      window.__editTest?.paintFinished("older edit"),
    );
    await page.evaluate(() => window.__editTest?.settle());
    expect(await page.evaluate(() => window.__editTest?.pixel(0))).toBe(
      expected,
    );
  });

  test("rejects an obsolete completed frame before the next viewport frame starts", async ({
    page,
  }) => {
    await setup(page, { ui: false, pages: ["alpha"] });
    const original = await page.evaluate(() => window.__editTest?.pixel(0));
    await page.evaluate(async () => {
      const t = window.__editTest;
      if (!t) throw new Error("Edit fixture is unavailable");
      t.holdPaint("obsolete edit");
      t.holdPaint("current edit");
      const session = await t.viewer.edit();
      await session.apply([
        { op: "setText", pageIndex: 0, text: "obsolete edit" },
      ]);
    });
    await page.waitForFunction(() =>
      window.__editTest?.paintStarted("obsolete edit"),
    );
    await page.evaluate(async () => {
      const t = window.__editTest;
      if (!t) throw new Error("Edit fixture is unavailable");
      t.holdFrames();
      const session = await t.viewer.edit();
      await session.apply([
        { op: "setText", pageIndex: 0, text: "current edit" },
      ]);
      t.releasePaint("obsolete edit");
    });
    // A separate browser task observes completion after its microtasks drain,
    // while the next viewport frame is explicitly held rather than timed.
    const beforeFrame = await page.evaluate(() => {
      const t = window.__editTest;
      if (!t) throw new Error("Edit fixture is unavailable");
      const snapshot = {
        pixel: t.pixel(0),
        obsoleteFinished: t.paintFinished("obsolete edit"),
        currentStarted: t.paintStarted("current edit"),
      };
      t.releaseFrames();
      return snapshot;
    });
    await page.waitForFunction(() =>
      window.__editTest?.paintStarted("current edit"),
    );
    await page.evaluate(() => window.__editTest?.releasePaint("current edit"));
    await expect
      .poll(() => page.evaluate(() => window.__editTest?.pixel(0)))
      .toBe(
        await page.evaluate(() =>
          window.__editTest?.expectedPixel("current edit"),
        ),
      );
    expect(beforeFrame).toEqual({
      pixel: original,
      obsoleteFinished: true,
      currentStarted: false,
    });
  });

  test("an unchanged page can finish painting after another page is edited", async ({
    page,
  }) => {
    await setup(page, { ui: false, pages: ["alpha", "beta"] });
    await page.evaluate(async () => {
      const t = window.__editTest;
      if (!t) throw new Error("Edit fixture is unavailable");
      t.holdPaint("beta edited");
      const session = await t.viewer.edit();
      await session.apply([
        { op: "setText", pageIndex: 1, text: "beta edited" },
      ]);
    });
    await page.waitForFunction(() =>
      window.__editTest?.paintStarted("beta edited"),
    );
    await page.evaluate(async () => {
      const t = window.__editTest;
      if (!t) throw new Error("Edit fixture is unavailable");
      const session = await t.viewer.edit();
      await session.apply([
        { op: "setText", pageIndex: 0, text: "alpha edited later" },
      ]);
    });
    const expectedFirst = await page.evaluate(() =>
      window.__editTest?.expectedPixel("alpha edited later"),
    );
    await expect
      .poll(() => page.evaluate(() => window.__editTest?.pixel(0)))
      .toBe(expectedFirst);
    await page.evaluate(() => window.__editTest?.releasePaint("beta edited"));
    await expect
      .poll(() => page.evaluate(() => window.__editTest?.pixel(1)))
      .toBe(
        await page.evaluate(() =>
          window.__editTest?.expectedPixel("beta edited"),
        ),
      );
    expect(await page.evaluate(() => window.__editTest?.pixel(0))).toBe(
      expectedFirst,
    );
  });

  test("a failed edit repaint leaves the last painted page visible", async ({
    page,
  }) => {
    await setup(page, { ui: false, pages: ["alpha"] });
    const before = await page.evaluate(() => window.__editTest?.pixel(0));
    await page.evaluate(async () => {
      const t = window.__editTest;
      if (!t) throw new Error("Edit fixture is unavailable");
      t.holdPaint("failed edit");
      const session = await t.viewer.edit();
      await session.apply([
        { op: "setText", pageIndex: 0, text: "failed edit" },
      ]);
    });
    await page.waitForFunction(() =>
      window.__editTest?.paintStarted("failed edit"),
    );
    await page.evaluate(() => window.__editTest?.failPaint("failed edit"));
    await page.waitForFunction(() =>
      window.__editTest?.paintFinished("failed edit"),
    );
    await page.evaluate(() => window.__editTest?.settle());
    expect(await page.evaluate(() => window.__editTest?.pixel(0))).toBe(before);
  });

  test("a different document clears old pixels and rejects a late edit paint", async ({
    page,
  }) => {
    await setup(page, { ui: false, pages: ["alpha"] });
    const original = await page.evaluate(() => window.__editTest?.pixel(0));
    await page.evaluate(async () => {
      const t = window.__editTest;
      if (!t) throw new Error("Edit fixture is unavailable");
      t.holdPaint("old document edit");
      t.holdPaint("replacement document");
      const session = await t.viewer.edit();
      await session.apply([
        { op: "setText", pageIndex: 0, text: "old document edit" },
      ]);
    });
    await page.waitForFunction(() =>
      window.__editTest?.paintStarted("old document edit"),
    );
    await page.evaluate(() =>
      window.__editTest?.loadDocument(["replacement document"]),
    );
    await page.waitForFunction(() =>
      window.__editTest?.paintStarted("replacement document"),
    );
    expect(await page.evaluate(() => window.__editTest?.pixel(0))).not.toBe(
      original,
    );
    expect(
      await page.evaluate(() => window.__editTest?.viewer.getEditSession()),
    ).toBeUndefined();
    await page.evaluate(() =>
      window.__editTest?.releasePaint("replacement document"),
    );
    const expected = await page.evaluate(() =>
      window.__editTest?.expectedPixel("replacement document"),
    );
    await expect
      .poll(() => page.evaluate(() => window.__editTest?.pixel(0)))
      .toBe(expected);
    await page.evaluate(() =>
      window.__editTest?.releasePaint("old document edit"),
    );
    await page.waitForFunction(() =>
      window.__editTest?.paintFinished("old document edit"),
    );
    await page.evaluate(() => window.__editTest?.settle());
    expect(await page.evaluate(() => window.__editTest?.pixel(0))).toBe(
      expected,
    );
  });
});
