import { readFile } from "node:fs/promises";

import { expect, test, type Page } from "@playwright/test";

import {
  localRecordOf,
  parseZip,
} from "../../packages/viewer/src/edit/ooxml/zip.js";
import { defaultResourceLimits } from "../../packages/viewer/src/limits.js";
import {
  buildDeck,
  group,
  syntheticDeck,
  textShape,
} from "../../packages/viewer/test/fixtures/pptx-builder.js";
import type { Viewer } from "../../packages/viewer/src/viewer.js";

/*
 * PPTX editing through the public API against the real adapter: the
 * @silurus/ooxml renderer paints, the OOXML edit worker edits. The renderer
 * is also loaded on its own (vendored by the example build) as the oracle
 * for element geometry.
 */

const CORPUS = new URL("../../.cache/corpus/", import.meta.url);

/** Names of the ZIP entries whose local records differ between two packages. */
function changedEntries(a: Uint8Array, b: Uint8Array): string[] {
  const left = parseZip(a, defaultResourceLimits);
  const right = parseZip(b, defaultResourceLimits);
  const changed: string[] = [];
  for (const entry of right.entries) {
    const before = left.entries.find(
      (candidate) => candidate.name === entry.name,
    );
    if (!before) {
      changed.push(entry.name);
      continue;
    }
    const x = localRecordOf(left, before);
    const y = localRecordOf(right, entry);
    const bytesBefore = a.subarray(x.headerOffset, x.recordEnd);
    const bytesAfter = b.subarray(y.headerOffset, y.recordEnd);
    if (
      bytesBefore.length !== bytesAfter.length ||
      bytesBefore.some((byte, index) => byte !== bytesAfter[index])
    )
      changed.push(entry.name);
  }
  for (const entry of left.entries)
    if (!right.entries.some((candidate) => candidate.name === entry.name))
      changed.push(`-${entry.name}`);
  return changed.sort();
}
const EDIT_ASSETS = ["/workers/ooxml-edit-worker.js"];
const RENDERER = "/vendor/ooxml/pptx.mjs";

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

test("keeps visible paint order for renderer-produced overflow hits", async ({
  page,
}) => {
  const unit = 9525;
  const original = buildDeck({
    slides: [
      {
        shapes: [
          textShape({
            id: 2,
            x: 40 * unit,
            y: 100 * unit,
            cx: 220 * unit,
            cy: 24 * unit,
            paragraphs: [
              [
                {
                  text: "Overflow text continues beneath the foreground shape on several lines",
                  rPr: 'sz="2400"',
                },
              ],
            ],
          }),
          textShape({
            id: 3,
            x: 20 * unit,
            y: 130 * unit,
            cx: 500 * unit,
            cy: 300 * unit,
            paragraphs: [[]],
            fill: '<a:solidFill><a:srgbClr val="FFFFFF"/></a:solidFill>',
          }),
          textShape({
            id: 4,
            x: 40 * unit,
            y: 100 * unit,
            cx: 100 * unit,
            cy: 20 * unit,
            paragraphs: [[{ text: "Top", rPr: 'sz="1200"' }]],
          }),
        ],
      },
    ],
  });
  await loadDeck(page, original);
  const result = await page.evaluate(async () => {
    const viewer = (window as typeof window & { __viewer: Viewer }).__viewer;
    const session = await viewer.edit();
    if (session.format !== "pptx") throw new Error("Expected a PPTX session");
    const selection = await viewer.selectText({
      startPageIndex: 0,
      endPageIndex: 0,
      startOffset: 0,
      endOffset: 1000,
    });
    const back = (await session.getElement("sld1:2")).item;
    if (!back) throw new Error("Expected the background text shape");
    const overflow = selection.runs.find(
      (run) =>
        run.shapeOrigin?.x === 40 &&
        run.shapeOrigin.y === 100 &&
        run.y > back.bounds.y + back.bounds.height,
    );
    if (!overflow)
      throw new Error(
        "The real renderer must produce a run below the back text frame",
      );
    const point = {
      x: overflow.x + overflow.width / 2,
      y: overflow.y + overflow.height / 2,
    };
    const covered = (await session.elementsAt(0, point)).items.map(
      (element) => element.id,
    );
    // The foreground covers the real run. Moving it away must reveal ink,
    // not merely change API metadata for an empty region.
    const canvas = document.createElement("canvas");
    await viewer.renderPage(0, canvas, { zoom: 1, devicePixelRatio: 1 });
    const context = canvas.getContext("2d");
    if (!context) throw new Error("Expected the rendered slide canvas");
    const ink = () => {
      const pixels = context.getImageData(
        Math.ceil(overflow.x),
        Math.ceil(overflow.y),
        Math.floor(overflow.width),
        Math.floor(overflow.height),
      ).data;
      let count = 0;
      for (let index = 0; index < pixels.length; index += 4) {
        if (pixels[index]! + pixels[index + 1]! + pixels[index + 2]! < 600)
          count += 1;
      }
      return count;
    };
    const coveredInk = ink();
    await session.moveElement({ target: "sld1:3", by: { dx: 500, dy: 0 } });
    const uncovered = (await session.elementsAt(0, point)).items.map(
      (element) => element.id,
    );
    await viewer.renderPage(0, canvas, { zoom: 1, devicePixelRatio: 1 });
    const uncoveredInk = ink();
    await session.replaceText({
      target: "sld1:2",
      text: "Edited visible text",
    });
    const saved = await session.save();
    await viewer.load(saved.bytes, { fileName: "edited.pptx" });
    const reopened = await viewer.edit();
    return {
      covered,
      uncovered,
      coveredInk,
      uncoveredInk,
      text: (await reopened.getElement("sld1:2")).item?.text,
      foregroundX: (await reopened.getElement("sld1:3")).item?.bounds.x,
      untouched: (await reopened.getElement("sld1:4")).item?.text,
    };
  });
  expect(result.covered).toEqual(["sld1:3", "sld1:2"]);
  expect(result.uncovered).toEqual(["sld1:2"]);
  expect(result.coveredInk).toBe(0);
  expect(result.uncoveredInk).toBeGreaterThan(10);
  expect(result.text).toBe("Edited visible text");
  expect(result.foregroundX).toBe(520);
  expect(result.untouched).toBe("Top");
});

for (const childSpace of ["explicit", "missing"] as const) {
  test(`resizes rotated groups with ${childSpace} child space through the browser worker without shearing child paint or history`, async ({
    page,
  }) => {
    await verifyGroupResizePaint(page, childSpace);
  });
}

async function verifyGroupResizePaint(
  page: Page,
  childSpace: "explicit" | "missing",
) {
  const unit = 9525;
  const groupXml = group({
    id: 2,
    x: 96 * unit,
    y: 96 * unit,
    cx: 96 * unit,
    cy: 96 * unit,
    rotation: 45,
    flipH: true,
    child: { x: 0, y: 0, cx: 96 * unit, cy: 96 * unit },
    children: [
      textShape({
        id: 3,
        x: 8 * unit,
        y: 8 * unit,
        cx: 24 * unit,
        cy: 16 * unit,
        paragraphs: [[]],
        fill: '<a:solidFill><a:srgbClr val="FF0000"/></a:solidFill>',
        line: "<a:ln><a:noFill/></a:ln>",
      }),
      textShape({
        id: 4,
        x: 60 * unit,
        y: 60 * unit,
        cx: 20 * unit,
        cy: 28 * unit,
        paragraphs: [[]],
        fill: '<a:solidFill><a:srgbClr val="0000FF"/></a:solidFill>',
        line: "<a:ln><a:noFill/></a:ln>",
      }),
    ],
  });
  const shape =
    childSpace === "missing"
      ? groupXml.replace(
          '<a:chOff x="0" y="0"/><a:chExt cx="914400" cy="914400"/>',
          "",
        )
      : groupXml;
  const original = buildDeck({ slides: [{ shapes: [shape] }] });
  const workers: string[] = [];
  page.on("worker", (worker) => workers.push(new URL(worker.url()).pathname));
  await loadDeck(page, original);
  const result = await page.evaluate(async () => {
    const viewer = (window as typeof window & { __viewer: Viewer }).__viewer;
    const session = await viewer.edit();
    if (session.format !== "pptx") throw new Error("Expected PPTX session");
    const canvas = document.createElement("canvas");
    const paint = async () => {
      await viewer.renderPage(0, canvas, { zoom: 1, devicePixelRatio: 1 });
      const context = canvas.getContext("2d");
      if (!context) throw new Error("Expected slide canvas context");
      const pixels = context.getImageData(
        0,
        0,
        canvas.width,
        canvas.height,
      ).data;
      const colorBounds = (color: "red" | "blue") => {
        let left = canvas.width,
          top = canvas.height,
          right = -1,
          bottom = -1,
          count = 0;
        for (let y = 0; y < canvas.height; y += 1) {
          for (let x = 0; x < canvas.width; x += 1) {
            const offset = (y * canvas.width + x) * 4;
            const red = pixels[offset],
              green = pixels[offset + 1],
              blue = pixels[offset + 2];
            if (red === undefined || green === undefined || blue === undefined)
              throw new Error("Incomplete raster pixel");
            if (
              green > 20 ||
              (color === "red"
                ? red < 240 || blue > 20
                : blue < 240 || red > 20)
            )
              continue;
            left = Math.min(left, x);
            right = Math.max(right, x);
            top = Math.min(top, y);
            bottom = Math.max(bottom, y);
            count += 1;
          }
        }
        if (count === 0) throw new Error(`Missing ${color} child paint`);
        return {
          x: left,
          y: top,
          width: right - left + 1,
          height: bottom - top + 1,
          count,
        };
      };
      return { pixels, red: colorBounds("red"), blue: colorBounds("blue") };
    };
    const equalBytes = (
      a: Uint8Array | Uint8ClampedArray,
      b: Uint8Array | Uint8ClampedArray,
    ) => a.length === b.length && a.every((byte, index) => byte === b[index]);
    const baseline = await paint();
    const originalBytes = (await session.save()).bytes;
    const initialState = session.state;
    const before = (await session.getElement("sld1:2")).item;
    if (!before) throw new Error("Expected group");
    const rejected: boolean[] = [];
    for (const dryRun of [false, true]) {
      try {
        await session.resizeElement(
          {
            target: before.id,
            rect: { ...before.bounds, width: before.bounds.width * 2 },
          },
          { dryRun },
        );
        rejected.push(false);
      } catch (error) {
        rejected.push(
          typeof error === "object" &&
            error !== null &&
            "code" in error &&
            error.code === "invalid-operation",
        );
      }
      if (
        !equalBytes((await session.save()).bytes, originalBytes) ||
        JSON.stringify(session.state) !== JSON.stringify(initialState) ||
        !equalBytes((await paint()).pixels, baseline.pixels)
      )
        throw new Error("Rejected resize changed bytes, state or paint");
    }
    const rect = {
      x: 70,
      y: 70,
      width: before.bounds.width * 2,
      height: before.bounds.height * 2,
    };
    await session.resizeElement({ target: before.id, rect });
    const resized = await paint();
    const childBounds = async () => {
      const red = (await session.getElement("sld1:3")).item;
      const blue = (await session.getElement("sld1:4")).item;
      if (!red || !blue) throw new Error("Expected both child elements");
      return { red: red.bounds, blue: blue.bounds };
    };
    const editedChildren = await childBounds();
    const editedBytes = (await session.save()).bytes;
    await session.undo();
    const undone =
      equalBytes((await session.save()).bytes, originalBytes) &&
      !session.state.canUndo &&
      equalBytes((await paint()).pixels, baseline.pixels);
    await viewer.load(editedBytes, { fileName: "resized-group.pptx" });
    const reopened = await viewer.edit();
    const reopenRed = (await reopened.getElement("sld1:3")).item;
    const reopenBlue = (await reopened.getElement("sld1:4")).item;
    const reopenedPaint = await paint();
    return {
      rejected,
      undone,
      initialGroup: before.bounds,
      beforePaint: { red: baseline.red, blue: baseline.blue },
      resizedPaint: { red: resized.red, blue: resized.blue },
      editedChildren,
      reopenedChildren: { red: reopenRed?.bounds, blue: reopenBlue?.bounds },
      paintSurvivedReload: equalBytes(resized.pixels, reopenedPaint.pixels),
    };
  });
  expect(workers).toContain("/workers/ooxml-edit-worker.js");
  expect(result.rejected).toEqual([true, true]);
  expect(result.undone).toBe(true);
  expect(result.paintSurvivedReload).toBe(true);
  expect(result.reopenedChildren).toEqual(result.editedChildren);
  // A fully opaque pixel threshold erodes antialiased polygon tips. Doubling
  // an already-eroded raster box is not a scale oracle. Derive the child
  // polygons from the fixture's known 45° rotation and horizontal reflection.
  const origin = 144 - 96 / Math.sqrt(2);
  expect(result.initialGroup.x).toBeCloseTo(origin, 5);
  expect(result.initialGroup.y).toBeCloseTo(origin, 5);
  const children = {
    red: { x: 8, y: 8, width: 24, height: 16 },
    blue: { x: 60, y: 60, width: 20, height: 28 },
  };
  for (const color of ["red", "blue"] as const) {
    const child = children[color];
    const corners = [
      { x: child.x, y: child.y },
      { x: child.x + child.width, y: child.y },
      { x: child.x, y: child.y + child.height },
      { x: child.x + child.width, y: child.y + child.height },
    ].map((point) => ({
      x: 144 + (96 - point.x - point.y) / Math.sqrt(2),
      y: 144 + (point.y - point.x) / Math.sqrt(2),
    }));
    const left = Math.min(...corners.map((point) => point.x));
    const right = Math.max(...corners.map((point) => point.x));
    const top = Math.min(...corners.map((point) => point.y));
    const bottom = Math.max(...corners.map((point) => point.y));
    for (const scale of [1, 2]) {
      const raster =
        scale === 1 ? result.beforePaint[color] : result.resizedPaint[color];
      const offset = scale === 1 ? origin : 70;
      const expected = {
        x: offset + (left - origin) * scale,
        y: offset + (top - origin) * scale,
        width: (right - left) * scale,
        height: (bottom - top) * scale,
      };
      expect(Math.abs(raster.x - expected.x)).toBeLessThanOrEqual(2);
      expect(Math.abs(raster.y - expected.y)).toBeLessThanOrEqual(2);
      expect(
        Math.abs(raster.x + raster.width - expected.x - expected.width),
      ).toBeLessThanOrEqual(2);
      expect(
        Math.abs(raster.y + raster.height - expected.y - expected.height),
      ).toBeLessThanOrEqual(2);
      // Fully covered pixels must include the polygon interior beyond a
      // sqrt(2)-pixel boundary strip (pixel-cell diagonal + grid sampling).
      const area = child.width * child.height * scale * scale;
      const perimeter = 2 * (child.width + child.height) * scale;
      expect(Math.abs(raster.count - area)).toBeLessThanOrEqual(
        perimeter * Math.SQRT2 + 4,
      );
      if (scale === 2) {
        const native = result.editedChildren[color];
        expect(native.x).toBeCloseTo(expected.x, 3);
        expect(native.y).toBeCloseTo(expected.y, 3);
        expect(native.width).toBeCloseTo(expected.width, 3);
        expect(native.height).toBeCloseTo(expected.height, 3);
      }
    }
  }
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

  // Only the slide's entry changed; every other entry is the original record.
  expect(changedEntries(original, new Uint8Array(after.saved))).toEqual([
    "ppt/slides/slide1.xml",
  ]);

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

test("recolours, moves, resizes, deletes and inserts shapes that the renderer draws where the engine says", async ({
  page,
}) => {
  const original = new Uint8Array(
    await readFile(new URL("sample.pptx", CORPUS)),
  );
  await loadDeck(page, original, "sample.pptx");
  const result = await page.evaluate(
    async ({ renderer }) => {
      const viewer = (window as unknown as { __viewer: any }).__viewer;
      const session = await viewer.edit();
      const title = (await session.getElement("sld1:2")).item;
      const receipt = await session.apply([
        {
          op: "setShapeStyle",
          target: "sld1:2",
          fill: "#FF0000",
          line: "none",
        },
        { op: "moveElement", target: "sld1:3", by: { dx: 40, dy: 60 } },
        {
          op: "resizeElement",
          target: "sld1:3",
          rect: { x: 200, y: 400, width: 300, height: 80 },
        },
        {
          op: "insertTextBox",
          pageIndex: 1,
          rect: { x: 100, y: 500, width: 400, height: 60 },
          text: "Inserted by web-doc",
          style: { fontSize: 32, bold: true },
        },
        { op: "deleteElement", target: "sld2:3" },
      ]);
      const elements = (await session.getElements()).items as {
        id: string;
        pageIndex: number;
        frame: {
          x: number;
          y: number;
          width: number;
          height: number;
          rotation: number;
        };
      }[];
      const saved = await session.save();
      const { PptxPresentation } = (await import(renderer)) as any;
      const presentation = await PptxPresentation.load(
        (saved.bytes as Uint8Array).slice().buffer,
        { useGoogleFonts: false, mode: "main" },
      );
      const oracle: Record<
        string,
        { x: number; y: number; width: number; height: number }
      > = {};
      for (let slide = 0; slide < presentation.slideCount; slide += 1) {
        const ids = elements
          .filter((element) => element.pageIndex === slide)
          .map((element) => element.id.split(":")[1]!);
        for (const item of await presentation.getElementBoundsByIds(slide, ids))
          oracle[`sld${slide + 1}:${item.elementId}`] = {
            x: item.bounds.x / 9525,
            y: item.bounds.y / 9525,
            width: item.bounds.width / 9525,
            height: item.bounds.height / 9525,
          };
      }
      presentation.destroy();
      return {
        titleBounds: title.bounds,
        createdIds: receipt.createdIds,
        removedIds: receipt.removedIds,
        changedPages: receipt.changedPages,
        elements,
        oracle,
        firstText: await viewer.getPageText(0),
        secondText: await viewer.getPageText(1),
      };
    },
    { renderer: RENDERER },
  );
  expect(result.createdIds).toEqual(["sld2:4"]);
  expect(result.removedIds).toEqual(["sld2:3"]);
  expect(result.changedPages).toEqual([0, 1]);
  expect(result.elements.map((element) => element.id)).toEqual([
    "sld1:2",
    "sld1:3",
    "sld2:2",
    "sld2:4",
  ]);
  expect(result.secondText).toContain("Inserted by web-doc");
  expect(result.secondText).not.toContain("bullet points");
  expect(result.firstText).toContain("Subtitle of the first slide");
  for (const element of result.elements) {
    const expected = result.oracle[element.id]!;
    expect(
      expected,
      `${element.id} known to the renderer after the edit`,
    ).toBeDefined();
    expect(Math.abs(element.frame.x - expected.x)).toBeLessThan(1);
    expect(Math.abs(element.frame.y - expected.y)).toBeLessThan(1);
    expect(Math.abs(element.frame.width - expected.width)).toBeLessThan(1);
    expect(Math.abs(element.frame.height - expected.height)).toBeLessThan(1);
  }
  const moved = result.elements.find((element) => element.id === "sld1:3")!;
  expect(moved.frame).toMatchObject({ x: 200, y: 400, width: 300, height: 80 });

  // The title's box is now red where it used to be white.
  const red = await page.evaluate(async (bounds) => {
    const viewer = (window as unknown as { __viewer: any }).__viewer;
    const canvas = document.createElement("canvas");
    await viewer.renderPage(0, canvas, { zoom: 1, devicePixelRatio: 1 });
    const { data } = canvas
      .getContext("2d")!
      .getImageData(Math.floor(bounds.x) + 2, Math.floor(bounds.y) + 2, 20, 20);
    let count = 0;
    for (let offset = 0; offset < data.length; offset += 4)
      if (
        data[offset]! > 200 &&
        data[offset + 1]! < 60 &&
        data[offset + 2]! < 60
      )
        count += 1;
    return count;
  }, result.titleBounds);
  expect(red).toBe(400);
});

test("inserts a picture and a table that the renderer draws, and edits a cell", async ({
  page,
}) => {
  const original = new Uint8Array(
    await readFile(new URL("sample.pptx", CORPUS)),
  );
  const png = Array.from(
    Buffer.from(
      "iVBORw0KGgoAAAANSUhEUgAAAAIAAAACCAIAAAD91JpzAAAAFElEQVR4nGP4z8DwHwyBFJAAgAAA//8R7AP8Ky7YKAAAAABJRU5ErkJggg==",
      "base64",
    ),
  );
  await loadDeck(page, original, "sample.pptx");
  const result = await page.evaluate(
    async ({ png, renderer }) => {
      const viewer = (window as unknown as { __viewer: any }).__viewer;
      const session = await viewer.edit();
      const receipt = await session.apply([
        {
          op: "insertImage",
          pageIndex: 1,
          rect: { x: 600, y: 80, width: 120, height: 120 },
          data: new Uint8Array(png),
          mimeType: "image/png",
        },
        {
          op: "insertTable",
          pageIndex: 1,
          rect: { x: 60, y: 420, width: 500, height: 100 },
          rows: [
            ["Metric", "Value"],
            ["Edits", "two"],
          ],
        },
        { op: "setTableCell", target: "$1", row: 1, column: 1, text: "three" },
      ]);
      const elements = (await session.getElements({ pageIndex: 1 })).items as {
        id: string;
        kind: string;
        frame: { x: number; y: number; width: number; height: number };
        table?: { rows: string[][] };
      }[];
      const saved = await session.save();
      const { PptxPresentation } = (await import(renderer)) as any;
      const presentation = await PptxPresentation.load(
        (saved.bytes as Uint8Array).slice().buffer,
        { useGoogleFonts: false, mode: "main" },
      );
      const bounds = await presentation.getElementBoundsByIds(
        1,
        elements.map((element) => element.id.split(":")[1]!),
      );
      presentation.destroy();
      const canvas = document.createElement("canvas");
      await viewer.renderPage(1, canvas, { zoom: 1, devicePixelRatio: 1 });
      const { data } = canvas.getContext("2d")!.getImageData(610, 90, 100, 100);
      let painted = 0;
      for (let offset = 0; offset < data.length; offset += 4)
        if (data[offset]! + data[offset + 1]! + data[offset + 2]! < 720)
          painted += 1;
      return {
        createdIds: receipt.createdIds,
        elements,
        oracle: bounds.map((item: any) => ({
          id: item.elementId,
          type: item.elementType,
          x: item.bounds.x / 9525,
          y: item.bounds.y / 9525,
          width: item.bounds.width / 9525,
          height: item.bounds.height / 9525,
        })),
        painted,
        text: await viewer.getPageText(1),
      };
    },
    { png, renderer: RENDERER },
  );
  expect(result.createdIds).toEqual(["sld2:4", "sld2:5"]);
  expect(result.elements.map((element) => element.kind)).toEqual([
    "shape",
    "shape",
    "image",
    "table",
  ]);
  expect(result.elements[3]!.table).toEqual({
    rows: [
      ["Metric", "Value"],
      ["Edits", "three"],
    ],
  });
  expect(result.text).toContain("Metric");
  expect(result.text).toContain("three");
  expect(result.painted).toBeGreaterThan(300);
  for (const element of result.elements) {
    const expected = result.oracle.find(
      (item: { id: string }) => item.id === element.id.split(":")[1],
    );
    expect(expected, `${element.id} drawn by the renderer`).toBeDefined();
    expect(Math.abs(element.frame.x - expected.x)).toBeLessThan(1);
    expect(Math.abs(element.frame.y - expected.y)).toBeLessThan(1);
    expect(Math.abs(element.frame.width - expected.width)).toBeLessThan(1);
    expect(Math.abs(element.frame.height - expected.height)).toBeLessThan(1);
  }
  expect(
    result.oracle.map((item: { type: string }) => item.type).sort(),
  ).toEqual(["picture", "shape", "shape", "table"].sort());
});

test("inserts, duplicates, moves and deletes slides that the renderer paints in the new order", async ({
  page,
}) => {
  const original = new Uint8Array(
    await readFile(new URL("sample.pptx", CORPUS)),
  );
  await loadDeck(page, original, "sample.pptx");
  const result = await page.evaluate(
    async ({ renderer }) => {
      const viewer = (window as unknown as { __viewer: any }).__viewer;
      const session = await viewer.edit();
      const layouts = (await session.getLayouts()).items as {
        id: string;
        name: string;
      }[];
      const titleOnly = layouts.find((layout) => layout.name === "Title Only")!;
      const inserted = await session.apply([
        { op: "insertSlide", index: 1, layout: titleOnly.id },
        { op: "replaceText", target: "$0", text: "Inserted slide title" },
      ]);
      const afterInsert = viewer.state.pageCount;
      const duplicated = await session.duplicateSlide({ pageIndex: 0 });
      const moved = await session.moveSlide({ from: 3, to: 0 });
      const removed = await session.deleteSlide({ pageIndex: 1 });
      const slides = (await session.getSlides()).items as {
        key: string;
        layout: string;
      }[];
      const texts: string[] = [];
      for (let index = 0; index < viewer.state.pageCount; index += 1)
        texts.push(await viewer.getPageText(index));
      const canvas = document.createElement("canvas");
      await viewer.renderPage(1, canvas, { zoom: 1, devicePixelRatio: 1 });
      const { data } = canvas
        .getContext("2d")!
        .getImageData(0, 0, canvas.width, canvas.height);
      let dark = 0;
      for (let offset = 0; offset < data.length; offset += 4)
        if (data[offset]! + data[offset + 1]! + data[offset + 2]! < 384)
          dark += 1;
      const saved = await session.save();
      const { PptxPresentation } = (await import(renderer)) as any;
      const presentation = await PptxPresentation.load(
        (saved.bytes as Uint8Array).slice().buffer,
        { useGoogleFonts: false, mode: "main" },
      );
      const slideCount = presentation.slideCount;
      presentation.destroy();
      return {
        titleOnly: titleOnly.id,
        inserted: {
          createdIds: inserted.createdIds,
          changedPages: inserted.changedPages,
        },
        afterInsert,
        duplicated: duplicated.createdIds,
        moved: moved.changedPages,
        removed: removed.removedIds,
        pageCount: viewer.state.pageCount,
        slides,
        texts,
        dark,
        slideCount,
        dirty: session.state.dirty,
      };
    },
    { renderer: RENDERER },
  );
  expect(result.inserted.createdIds).toEqual(["sld3:2"]);
  expect(result.inserted.changedPages).toEqual([1, 2]);
  expect(result.afterInsert).toBe(3);
  expect(result.duplicated).toEqual(["sld4:2", "sld4:3"]);
  expect(result.moved).toEqual([0, 1, 2, 3]);
  expect(result.removed).toEqual(["sld1:2", "sld1:3"]);
  expect(result.pageCount).toBe(3);
  expect(result.slideCount).toBe(3);
  // [sld1, sld2] → insert → [sld1, sld3, sld2] → duplicate → [sld1, sld4, sld3, sld2]
  // → move 3 to 0 → [sld2, sld1, sld4, sld3] → delete 1 → [sld2, sld4, sld3].
  expect(result.slides.map((slide) => slide.key)).toEqual([
    "sld2",
    "sld4",
    "sld3",
  ]);
  expect(result.slides[2]!.layout).toBe(result.titleOnly);
  expect(result.texts[0]).toContain("This is the second slide");
  expect(result.texts[1]).toContain("Title of the first slide");
  expect(result.texts[2]).toContain("Inserted slide title");
  expect(result.dark).toBeGreaterThan(50);
  expect(result.dirty).toBe(true);
});

test(
  "applies an edit within the budget on 10-, 100- and 500-slide decks",
  { tag: "@performance" },
  async ({ page }) => {
    test.setTimeout(180_000);
    const timings: Record<string, number[]> = {};
    for (const count of [10, 100, 500]) {
      await loadDeck(page, syntheticDeck(count), `deck-${count}.pptx`);
      timings[count] = await page.evaluate(async () => {
        const viewer = (window as unknown as { __viewer: any }).__viewer;
        const session = await viewer.edit();
        const out: number[] = [];
        for (const text of ["First edit", "Second edit"]) {
          const started = performance.now();
          await session.replaceText({ target: "sld1:3", text });
          out.push(performance.now() - started);
        }
        const started = performance.now();
        await session.insertTextBox({
          pageIndex: viewer.state.pageCount - 1,
          rect: { x: 50, y: 50, width: 300, height: 40 },
          text: "Last slide",
        });
        out.push(performance.now() - started);
        return out;
      });
      console.log(
        `pptx apply ${count} slides: replaceText ${timings[count]![0]!.toFixed(0)} ms then ${timings[count]![1]!.toFixed(0)} ms, insertTextBox on the last slide ${timings[count]![2]!.toFixed(0)} ms`,
      );
      for (const value of timings[count]!) expect(value).toBeLessThan(3000);
    }
  },
);

test(
  "spike: renderer load time for 10, 100 and 500 slides, with and without progressive layout",
  { tag: "@performance" },
  async ({ page }) => {
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
            out[`${count}/${progressive ? "progressive" : "full"}/load`] =
              loaded;
            out[`${count}/${progressive ? "progressive" : "full"}/firstPaint`] =
              painted;
            out[`${count}/${progressive ? "progressive" : "full"}/complete`] =
              complete;
          }
        return out;
      },
      {
        decks: decks.map(
          ([count, bytes]) => [count, Array.from(bytes)] as const,
        ),
        renderer: RENDERER,
      },
    );
    for (const [key, value] of Object.entries(timings))
      console.log(`pptx renderer ${key}: ${value.toFixed(1)} ms`);
    expect(timings["500/full/load"]).toBeLessThan(10_000);
  },
);
