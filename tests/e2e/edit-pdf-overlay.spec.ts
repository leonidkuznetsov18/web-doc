import { expect, test, type Page } from "@playwright/test";
import { readFileSync } from "node:fs";

import { buildPdf } from "../../packages/viewer/test/fixtures/pdf-builder.js";
import {
  CFF_TEXT_WIDTHS,
  cffTextPdf,
  strippedTrueType,
  trueTypeTextPdf,
} from "../../packages/viewer/test/fixtures/pdf-fonts.js";

/*
 * The PDF overlay primitives (ACTION-825) against the real viewer: PDF.js
 * draws the page and its text layer, the PDFium worker answers the layout,
 * selection and render reads, and a host-style flow selects text with the
 * mouse, resolves it to a range, replaces it and restores the selection.
 */

type Rect = { left: number; top: number; width: number; height: number };

async function mountPdf(page: Page, bytes: Uint8Array): Promise<void> {
  await page.goto("/");
  await page.evaluate(async (data) => {
    const { ViewerClient } = (await import("/main.js")) as any;
    const host = document.createElement("div");
    host.dataset.testid = "overlay-host";
    Object.assign(host.style, { width: "800px", height: "900px" });
    document.body.replaceChildren(host);
    const client = ViewerClient.create({
      assetBaseUrl: new URL("/", location.href),
    });
    const viewer = client.createViewer({ container: host, initialZoom: 1 });
    await viewer.load(new Uint8Array(data), { fileName: "fixture.pdf" });
    (window as unknown as { __overlayViewer: unknown }).__overlayViewer =
      viewer;
  }, Array.from(bytes));
  await page
    .locator('[data-zrimo-layer="text"] [data-start]')
    .first()
    .waitFor();
}

const variants = [
  { name: "upright", page: { rotation: 0 as const } },
  { name: "rotated 90°", page: { rotation: 1 as const } },
  { name: "rotated 180°", page: { rotation: 2 as const } },
  { name: "rotated 270°", page: { rotation: 3 as const } },
  { name: "cropped", page: { cropBox: [50, 50, 400, 500] as const } },
];

for (const variant of variants)
  test(`glyph layout agrees with the PDF.js text layer within 1 CSS px on a ${variant.name} page`, async ({
    page,
  }) => {
    const y = "cropBox" in variant.page ? 120 : 700;
    await mountPdf(
      page,
      await buildPdf([
        {
          texts: [{ text: "Hello world", x: 72, y, fontSize: 24 }],
          ...variant.page,
        },
      ]),
    );
    const result = await page.evaluate(async () => {
      const viewer = (window as any).__overlayViewer;
      const session = await viewer.edit();
      const { item: layout } = await session.getTextLayout("p0:o0");
      const line = layout.lines[0];
      const advance = line.glyphs.reduce(
        (sum: number, glyph: { advance: number }) => sum + glyph.advance,
        0,
      );
      const { item: element } = await session.getElement("p0:o0");
      // The page's rotation turns the reading direction.
      const rotation = (() => {
        const first = line.glyphs[0].box;
        const last = line.glyphs[line.glyphs.length - 1].box;
        const dx = last.x - first.x;
        const dy = last.y - first.y;
        if (Math.abs(dx) >= Math.abs(dy)) return dx >= 0 ? 0 : 2;
        return dy >= 0 ? 1 : 3;
      })();
      const start = line.baseline;
      const end = [
        { x: start.x + advance, y: start.y },
        { x: start.x, y: start.y + advance },
        { x: start.x - advance, y: start.y },
        { x: start.x, y: start.y - advance },
      ][rotation];
      const point = (p: { x: number; y: number }) =>
        viewer.pageToClient(0, { ...p, width: 0, height: 0 });
      const span = [
        ...document.querySelectorAll<HTMLElement>(
          '[data-zrimo-layer="text"] [data-start]',
        ),
      ].find((candidate) => candidate.textContent?.includes("Hello"))!;
      const rect = span.getBoundingClientRect();
      return {
        rotation,
        start: point(start),
        end: point(end),
        bounds: viewer.pageToClient(0, line.bounds),
        elementBounds: viewer.pageToClient(0, element.bounds),
        span: {
          left: rect.left,
          top: rect.top,
          width: rect.width,
          height: rect.height,
        } as Rect,
        text: span.textContent,
      };
    });
    expect(result.text).toBe("Hello world");
    const { span } = result;
    const along = result.rotation % 2 === 0 ? "x" : "y";
    const spanStart = along === "x" ? span.left : span.top;
    const spanEnd =
      along === "x" ? span.left + span.width : span.top + span.height;
    const starts = [
      result.start[along === "x" ? "left" : "top"],
      result.end[along === "x" ? "left" : "top"],
    ];
    // Where the text starts and ends along the reading direction agrees with
    // the text layer's span within one CSS pixel.
    expect(Math.abs(Math.min(...starts) - spanStart)).toBeLessThanOrEqual(1);
    expect(Math.abs(Math.max(...starts) - spanEnd)).toBeLessThanOrEqual(1);
    // Across it the ink lies inside the span's box.
    const crossStart = along === "x" ? span.top : span.left;
    const crossEnd =
      along === "x" ? span.top + span.height : span.left + span.width;
    const inkStart = along === "x" ? result.bounds.top : result.bounds.left;
    const inkEnd =
      along === "x"
        ? result.bounds.top + result.bounds.height
        : result.bounds.left + result.bounds.width;
    expect(inkStart).toBeGreaterThanOrEqual(crossStart - 1);
    expect(inkEnd).toBeLessThanOrEqual(crossEnd + 1);
  });

test("selects text in the viewer, resolves it to a range, replaces it and restores the selection", async ({
  page,
}) => {
  await mountPdf(
    page,
    await buildPdf([
      { texts: [{ text: "Hello brave world", x: 72, y: 700, fontSize: 24 }] },
    ]),
  );
  const handles = await page.evaluate(async () => {
    const viewer = (window as any).__overlayViewer;
    const session = await viewer.edit();
    const { item: layout } = await session.getTextLayout("p0:o0");
    const glyph = (offset: number) =>
      viewer.pageToClient(0, layout.lines[0].glyphs[offset].box);
    const from = glyph(6);
    const to = glyph(10);
    return {
      from: { x: from.left + 1, y: from.top + from.height / 2 },
      to: { x: to.left + to.width - 1, y: to.top + to.height / 2 },
    };
  });
  await page.mouse.move(handles.from.x, handles.from.y);
  await page.mouse.down();
  await page.mouse.move(handles.to.x, handles.to.y, { steps: 12 });
  await page.mouse.up();
  await expect
    .poll(() =>
      page.evaluate(
        () => (window as any).__overlayViewer.getSelection()?.text ?? "",
      ),
    )
    .toBe("brave");
  const result = await page.evaluate(async () => {
    const viewer = (window as any).__overlayViewer;
    const session = viewer.getEditSession();
    const selection = viewer.getSelection();
    const { items: ranges } = await session.elementsForSelection(selection);
    const range = ranges[0];
    const from = session.state.revision;
    const receipt = await session.replaceText({
      target: range.start.elementId,
      text: "bold",
      range,
    });
    const { item: mapped } = await session.mapRange(range, from);
    // Restore the selection on the new text: the page's logical offsets are
    // the element's here, one text object on the page.
    await viewer.selectText({
      startPageIndex: 0,
      startOffset: mapped.start.offset,
      endPageIndex: 0,
      endOffset: mapped.end.offset,
    });
    const after = (await session.getTextLayout("p0:o0")).item;
    return {
      ranges,
      createdIds: receipt.createdIds,
      mapped,
      selected: viewer.getSelection()?.text,
      pageText: await viewer.getPageText(0),
      lineText: after.lines[0].text,
    };
  });
  expect(result.ranges).toEqual([
    {
      start: { elementId: "p0:o0", offset: 6 },
      end: { elementId: "p0:o0", offset: 11 },
    },
  ]);
  expect(result.createdIds).toEqual([]);
  expect(result.mapped).toEqual({
    start: { elementId: "p0:o0", offset: 6 },
    end: { elementId: "p0:o0", offset: 10 },
  });
  expect(result.selected).toBe("bold");
  expect(result.pageText).toContain("Hello bold world");
  expect(result.lineText).toBe("Hello bold world");
  if (process.env.OVERLAY_PROOF)
    await page.screenshot({
      path: `${process.env.OVERLAY_PROOF}/t37-selection.png`,
      clip: { x: 0, y: 0, width: 800, height: 420 },
    });
});

test("a suppressed render stands in for an element and the next normal render shows it again", async ({
  page,
}) => {
  await mountPdf(
    page,
    await buildPdf([
      {
        texts: [{ text: "Kept", x: 300, y: 400, fontSize: 36 }],
        rect: { x: 100, y: 600, width: 100, height: 100, fill: [220, 20, 20] },
      },
    ]),
  );
  const result = await page.evaluate(async () => {
    const viewer = (window as any).__overlayViewer;
    const session = await viewer.edit();
    const { items } = await session.getElements({ pageIndex: 0 });
    const square = items.find(
      (element: { kind: string }) => element.kind === "shape",
    );
    const centre = {
      x: square.bounds.x + square.bounds.width / 2,
      y: square.bounds.y + square.bounds.height / 2,
    };
    const pixel = (canvas: HTMLCanvasElement, scale: number) => {
      const data = canvas
        .getContext("2d")!
        .getImageData(
          Math.floor(centre.x * scale),
          Math.floor(centre.y * scale),
          1,
          1,
        ).data;
      return [data[0], data[1], data[2]];
    };
    const { item: bitmap } = await session.renderPageWithout(0, [square.id], {
      scale: 1,
    });
    const overlay = document.createElement("canvas");
    overlay.width = bitmap.width;
    overlay.height = bitmap.height;
    overlay
      .getContext("2d")!
      .putImageData(
        new ImageData(
          new Uint8ClampedArray(bitmap.data),
          bitmap.width,
          bitmap.height,
        ),
        0,
        0,
      );
    // The host lays the stand-in over the page while its input is open.
    const place = viewer.pageToClient(0, {
      x: 0,
      y: 0,
      width: bitmap.width / bitmap.scale,
      height: bitmap.height / bitmap.scale,
    });
    Object.assign(overlay.style, {
      position: "fixed",
      left: `${place.left}px`,
      top: `${place.top}px`,
      width: `${place.width}px`,
      height: `${place.height}px`,
      pointerEvents: "none",
    });
    overlay.dataset.testid = "stand-in";
    document.body.append(overlay);
    const normal = document.createElement("canvas");
    await viewer.renderPage(0, normal, { zoom: 1, devicePixelRatio: 1 });
    const saved = await session.save();
    return {
      size: [bitmap.width, bitmap.height],
      suppressed: pixel(overlay, 1),
      normal: pixel(normal, 1),
      revision: session.state.revision,
      dirty: session.state.dirty,
      bytes: saved.bytes.length,
    };
  });
  expect(result.size).toEqual([612, 792]);
  expect(result.suppressed).toEqual([255, 255, 255]);
  expect(result.normal[0]).toBeGreaterThan(180);
  expect(result.normal[1]).toBeLessThan(80);
  expect(result.revision).toBe(0);
  expect(result.dirty).toBe(false);
  if (process.env.OVERLAY_PROOF)
    await page.screenshot({
      path: `${process.env.OVERLAY_PROOF}/t37-stand-in.png`,
      clip: { x: 0, y: 0, width: 800, height: 420 },
    });
});

test("a host loads an element's embedded CFF font and types with the file's advances", async ({
  page,
}) => {
  await mountPdf(page, cffTextPdf());
  const result = await page.evaluate(async (characters) => {
    const viewer = (window as any).__overlayViewer;
    const session = await viewer.edit();
    const { item: font } = await session.getTextFont("p0:o0");
    const face = new FontFace(`face-${font.key}`, font.face.data);
    await face.load();
    document.fonts.add(face);
    const context = document.createElement("canvas").getContext("2d")!;
    context.font = `1000px "face-${font.key}"`;
    return {
      format: font.face.format,
      status: face.status,
      widths: Object.fromEntries(
        characters.map((character: string) => [
          character,
          context.measureText(character).width,
        ]),
      ),
    };
  }, Object.keys(CFF_TEXT_WIDTHS));
  expect(result.format).toBe("opentype");
  expect(result.status).toBe("loaded");
  // At 1000 px a font unit is a pixel: the advances are the PDF's widths.
  for (const [character, width] of Object.entries(CFF_TEXT_WIDTHS))
    expect(result.widths[character], character).toBeCloseTo(width, 0);
});

test("a host loads an embedded TrueType subset that lacks the tables browsers require", async ({
  page,
}) => {
  const noto = new Uint8Array(
    readFileSync(
      new URL(
        "../../packages/viewer/fonts/noto-sans-latin-cyrillic.ttf",
        import.meta.url,
      ),
    ),
  );
  await mountPdf(page, await trueTypeTextPdf(strippedTrueType(noto), "Привіт"));
  const result = await page.evaluate(async () => {
    const viewer = (window as any).__overlayViewer;
    const session = await viewer.edit();
    const { item: font } = await session.getTextFont("p0:o0");
    const face = new FontFace(`face-${font.key}`, font.face.data);
    await face.load();
    return { format: font.face.format, status: face.status };
  });
  expect(result).toEqual({ format: "truetype", status: "loaded" });
});
