import { readFile } from "node:fs/promises";

import { expect, test, type Page } from "@playwright/test";

import { syntheticDeck } from "../../packages/viewer/test/fixtures/pptx-builder.js";

/*
 * PPTX editing through the public API against the real adapter: the
 * @silurus/ooxml renderer paints, the OOXML edit worker edits. The renderer
 * is also loaded on its own (vendored by the example build) as the oracle
 * for element geometry.
 */

const CORPUS = new URL("../../.cache/corpus/", import.meta.url);
const EDIT_ASSETS = ["/workers/ooxml-edit-worker.js"];
const RENDERER = "/vendor/ooxml-pptx/pptx.mjs";

async function loadDeck(
  page: Page,
  bytes: Uint8Array,
  fileName = "fixture.pptx",
): Promise<void> {
  await page.goto("/");
  await page.evaluate(
    async ({ data, fileName }) => {
      const { ViewerClient } = (await import("/main.js")) as {
        ViewerClient: { create(config: unknown): { createViewer(): unknown } };
      };
      const client = ViewerClient.create({
        assetBaseUrl: new URL("/", location.href),
      });
      const viewer = client.createViewer() as {
        load(bytes: Uint8Array, options: unknown): Promise<void>;
      };
      await viewer.load(new Uint8Array(data), { fileName });
      (window as unknown as { __viewer: unknown }).__viewer = viewer;
    },
    { data: Array.from(bytes), fileName },
  );
}

test("starts the OOXML worker only on edit() and lists every shape with the renderer's geometry", async ({
  page,
}) => {
  const original = new Uint8Array(
    await readFile(new URL("sample.pptx", CORPUS)),
  );
  const requests: string[] = [];
  page.on("request", (request) =>
    requests.push(new URL(request.url()).pathname),
  );
  await loadDeck(page, original, "sample.pptx");

  const info = await page.evaluate(() => {
    const viewer = (window as unknown as { __viewer: any }).__viewer;
    return {
      editing: viewer.getDocumentInfo().capabilities.editing,
      pageCount: viewer.state.pageCount,
    };
  });
  expect(info).toEqual({ editing: true, pageCount: 2 });
  expect(requests.filter((path) => EDIT_ASSETS.includes(path))).toEqual([]);

  const result = await page.evaluate(
    async ({ data, renderer }) => {
      const viewer = (window as unknown as { __viewer: any }).__viewer;
      const session = await viewer.edit();
      const elements = (await session.getElements()).items as {
        id: string;
        pageIndex: number;
        kind: string;
        name: string;
        text?: string;
        bounds: { x: number; y: number; width: number; height: number };
        frame?: {
          x: number;
          y: number;
          width: number;
          height: number;
          rotation: number;
          flipH: boolean;
          flipV: boolean;
        };
      }[];
      const slides = (await session.getSlides()).items;
      const saved = await session.save();
      const identical =
        saved.bytes.length === data.length &&
        saved.bytes.every(
          (byte: number, index: number) => byte === data[index],
        );

      const { PptxPresentation } = (await import(renderer)) as any;
      const presentation = await PptxPresentation.load(
        new Uint8Array(data).buffer,
        { useGoogleFonts: false, mode: "main" },
      );
      const EMU_PER_PX = 9525;
      const oracle: Record<string, unknown> = {};
      for (let slide = 0; slide < presentation.slideCount; slide += 1) {
        const ids = elements
          .filter((element) => element.pageIndex === slide)
          .map((element) => element.id.split(":")[1]!.split("#")[0]!);
        const bounds = await presentation.getElementBoundsByIds(slide, ids);
        for (const item of bounds)
          oracle[`sld${slide + 1}:${item.elementId}`] = {
            origin: item.origin,
            x: item.bounds.x / EMU_PER_PX,
            y: item.bounds.y / EMU_PER_PX,
            width: item.bounds.width / EMU_PER_PX,
            height: item.bounds.height / EMU_PER_PX,
            rotation: item.bounds.rotation,
          };
      }
      presentation.destroy();
      return {
        elements,
        slides,
        identical,
        oracle,
        sessionFormat: session.format,
      };
    },
    { data: Array.from(original), renderer: RENDERER },
  );
  expect(requests.filter((path) => EDIT_ASSETS.includes(path))).toEqual(
    EDIT_ASSETS,
  );
  expect(result.sessionFormat).toBe("pptx");
  expect(result.identical).toBe(true);
  expect(result.elements.map((element) => element.id)).toEqual([
    "sld1:2",
    "sld1:3",
    "sld2:2",
    "sld2:3",
  ]);
  expect(result.slides.map((slide: { key: string }) => slide.key)).toEqual([
    "sld1",
    "sld2",
  ]);
  for (const element of result.elements) {
    const expected = result.oracle[element.id] as {
      origin: string;
      x: number;
      y: number;
      width: number;
      height: number;
      rotation: number;
    };
    expect(expected, `${element.id} known to the renderer`).toBeDefined();
    expect(expected.origin).toBe("slide");
    expect(Math.abs(element.frame!.x - expected.x)).toBeLessThan(1);
    expect(Math.abs(element.frame!.y - expected.y)).toBeLessThan(1);
    expect(Math.abs(element.frame!.width - expected.width)).toBeLessThan(1);
    expect(Math.abs(element.frame!.height - expected.height)).toBeLessThan(1);
    expect(element.frame!.rotation).toBe(expected.rotation);
  }
});

test("reads the elements of a built deck like the renderer: rotation, groups and placeholders", async ({
  page,
}) => {
  const deck = syntheticDeck(3);
  await loadDeck(page, deck);
  const result = await page.evaluate(
    async ({ data, renderer }) => {
      const viewer = (window as unknown as { __viewer: any }).__viewer;
      const session = await viewer.edit();
      const elements = (await session.getElements({ pageIndex: 1 })).items as {
        id: string;
        frame: {
          x: number;
          y: number;
          width: number;
          height: number;
          rotation: number;
        };
        text?: string;
      }[];
      const { PptxPresentation } = (await import(renderer)) as any;
      const presentation = await PptxPresentation.load(
        new Uint8Array(data).buffer,
        { useGoogleFonts: false, mode: "main" },
      );
      const bounds = await presentation.getElementBoundsByIds(
        1,
        elements.map((element) => element.id.split(":")[1]!),
      );
      const text = await viewer.getPageText(1);
      presentation.destroy();
      return {
        elements,
        oracle: bounds.map((item: any) => ({
          id: item.elementId,
          origin: item.origin,
          x: item.bounds.x / 9525,
          y: item.bounds.y / 9525,
          width: item.bounds.width / 9525,
          height: item.bounds.height / 9525,
        })),
        text,
      };
    },
    { data: Array.from(deck), renderer: RENDERER },
  );
  expect(result.elements.map((element) => element.id)).toEqual([
    "sld2:2",
    "sld2:3",
  ]);
  expect(result.text).toContain("Slide 2");
  expect(result.text).toContain("Body text of slide 2");
  for (const element of result.elements) {
    const expected = result.oracle.find(
      (item: { id: string }) => item.id === element.id.split(":")[1],
    );
    expect(expected).toBeDefined();
    expect(expected.origin).toBe("slide");
    expect(Math.abs(element.frame.x - expected.x)).toBeLessThan(1);
    expect(Math.abs(element.frame.y - expected.y)).toBeLessThan(1);
    expect(Math.abs(element.frame.width - expected.width)).toBeLessThan(1);
    expect(Math.abs(element.frame.height - expected.height)).toBeLessThan(1);
  }
});

/** Dark pixels inside a slide-space rectangle of the first page's canvas. */
async function darkPixelsIn(
  page: Page,
  pageIndex: number,
  rect: { x: number; y: number; width: number; height: number },
): Promise<number> {
  return page.evaluate(
    async ({ pageIndex, rect }) => {
      const viewer = (window as unknown as { __viewer: any }).__viewer;
      const canvas = document.createElement("canvas");
      await viewer.renderPage(pageIndex, canvas, {
        zoom: 1,
        devicePixelRatio: 1,
      });
      const context = canvas.getContext("2d")!;
      const x = Math.max(0, Math.floor(rect.x));
      const y = Math.max(0, Math.floor(rect.y));
      const width = Math.min(canvas.width - x, Math.ceil(rect.width));
      const height = Math.min(canvas.height - y, Math.ceil(rect.height));
      const { data } = context.getImageData(x, y, width, height);
      let dark = 0;
      for (let offset = 0; offset < data.length; offset += 4)
        if (data[offset]! + data[offset + 1]! + data[offset + 2]! < 384)
          dark += 1;
      return dark;
    },
    { pageIndex, rect },
  );
}

test("replaces and restyles text so the renderer shows it, and the edit survives save and reload", async ({
  page,
}) => {
  const original = new Uint8Array(
    await readFile(new URL("sample.pptx", CORPUS)),
  );
  await loadDeck(page, original, "sample.pptx");
  const before = await page.evaluate(async () => {
    const viewer = (window as unknown as { __viewer: any }).__viewer;
    const session = await viewer.edit();
    const title = (await session.getElement("sld1:2")).item;
    return { bounds: title.bounds, text: await viewer.getPageText(0) };
  });
  expect(before.text).toContain("Title of the first slide");
  const darkBefore = await darkPixelsIn(page, 0, before.bounds);

  const after = await page.evaluate(async () => {
    const viewer = (window as unknown as { __viewer: any }).__viewer;
    const session = viewer.getEditSession();
    const receipt = await session.replaceText({
      target: "sld1:2",
      text: "Patched by web-doc",
    });
    const styled = await session.setTextStyle({
      target: "sld1:2",
      style: { bold: true, color: "#FF0000", fontSize: 60 },
    });
    const title = (await session.getElement("sld1:2")).item;
    const saved = await session.save();
    return {
      revision: styled.revision,
      changedPages: receipt.changedPages,
      text: await viewer.getPageText(0),
      title,
      saved: Array.from(saved.bytes as Uint8Array),
      dirty: session.state.dirty,
    };
  });
  expect(after.revision).toBe(2);
  expect(after.changedPages).toEqual([0]);
  expect(after.dirty).toBe(true);
  expect(after.text).toContain("Patched by web-doc");
  expect(after.text).not.toContain("Title of the first slide");
  expect(after.title.text).toBe("Patched by web-doc");
  expect(after.title.textStyle).toMatchObject({
    bold: true,
    color: "#FF0000",
    fontSize: 60,
  });
  const darkAfter = await darkPixelsIn(page, 0, before.bounds);
  expect(darkAfter).toBeGreaterThan(50);
  expect(darkAfter).not.toBe(darkBefore);

  // The saved bytes reopen with the edit; nothing else of the deck changed.
  const reloaded = await page.evaluate(async (data) => {
    const viewer = (window as unknown as { __viewer: any }).__viewer;
    await viewer.load(new Uint8Array(data), { fileName: "edited.pptx" });
    const session = await viewer.edit();
    const title = (await session.getElement("sld1:2")).item;
    return {
      text: await viewer.getPageText(0),
      pageCount: viewer.state.pageCount,
      bold: title.textStyle.bold,
      second: await viewer.getPageText(1),
    };
  }, after.saved);
  expect(reloaded.pageCount).toBe(2);
  expect(reloaded.text).toContain("Patched by web-doc");
  expect(reloaded.bold).toBe(true);
  expect(reloaded.second).toContain("This is the second slide");
});

test("spike: renderer load time for 10, 100 and 500 slides, with and without progressive layout", async ({
  page,
}) => {
  test.setTimeout(180_000);
  const decks = [10, 100, 500].map(
    (count) => [count, syntheticDeck(count)] as const,
  );
  await page.goto("/");
  const timings = await page.evaluate(
    async ({ decks, renderer }) => {
      const { PptxPresentation } = (await import(renderer)) as any;
      const out: Record<string, number> = {};
      for (const [count, data] of decks)
        for (const progressive of [false, true]) {
          const buffer = new Uint8Array(data).buffer;
          const started = performance.now();
          const presentation = await PptxPresentation.load(buffer, {
            useGoogleFonts: false,
            mode: "main",
            progressiveLayout: progressive,
          });
          const loaded = performance.now() - started;
          const canvas = document.createElement("canvas");
          await presentation.renderSlide(canvas, 0, { width: 960, dpr: 1 });
          const painted = performance.now() - started;
          await presentation.waitUntilLayoutComplete?.();
          const complete = performance.now() - started;
          presentation.destroy();
          out[`${count}/${progressive ? "progressive" : "full"}/load`] = loaded;
          out[`${count}/${progressive ? "progressive" : "full"}/firstPaint`] =
            painted;
          out[`${count}/${progressive ? "progressive" : "full"}/complete`] =
            complete;
        }
      return out;
    },
    {
      decks: decks.map(([count, bytes]) => [count, Array.from(bytes)] as const),
      renderer: RENDERER,
    },
  );
  for (const [key, value] of Object.entries(timings))
    console.log(`pptx renderer ${key}: ${value.toFixed(1)} ms`);
  expect(timings["500/full/load"]).toBeLessThan(10_000);
});
