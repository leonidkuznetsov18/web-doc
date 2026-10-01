import { mkdirSync } from "node:fs";
import { test, type Page } from "@playwright/test";

import { buildPdf } from "../../packages/viewer/test/fixtures/pdf-builder.js";

/*
 * A scripted walk through the PDF editing API in the mounted viewer. It is
 * not a test of behaviour (edit-core.spec.ts and edit-pdf.spec.ts are): it
 * captures one screenshot per step as proof for the ticket, and the frames
 * become the demo GIF. Run it on its own:
 *
 *   node node_modules/@playwright/test/cli.js test --config=playwright.demo.config.ts
 *
 * Frames land in $EDIT_DEMO_DIR (default test-results/edit-demo).
 */

const OUT = process.env.EDIT_DEMO_DIR ?? "test-results/edit-demo";

declare global {
  interface Window {
    __demo: {
      viewer: any;
      caption(text: string): void;
      settle(): Promise<void>;
      pngBytes(): Promise<Uint8Array>;
    };
  }
}

async function mount(page: Page, bytes: Uint8Array): Promise<void> {
  await page.goto("/");
  await page.setViewportSize({ width: 960, height: 760 });
  await page.evaluate(async (data) => {
    const { ViewerClient } = (await import("/main.js")) as {
      ViewerClient: {
        create(config: unknown): { createViewer(config: unknown): any };
      };
    };
    const wrapper = document.createElement("div");
    wrapper.id = "demo";
    Object.assign(wrapper.style, {
      position: "fixed",
      left: "0",
      top: "0",
      width: "960px",
      height: "760px",
      background: "#f4f4f5",
      fontFamily: "system-ui, sans-serif",
    });
    const caption = document.createElement("div");
    Object.assign(caption.style, {
      height: "40px",
      lineHeight: "40px",
      padding: "0 16px",
      fontSize: "16px",
      fontWeight: "600",
      color: "#18181b",
      background: "#e4e4e7",
    });
    const container = document.createElement("div");
    Object.assign(container.style, {
      position: "absolute",
      left: "0",
      top: "40px",
      width: "960px",
      height: "720px",
      overflow: "hidden",
    });
    wrapper.append(caption, container);
    document.body.append(wrapper);
    const client = ViewerClient.create({
      assetBaseUrl: new URL("/", location.href),
    });
    const viewer = client.createViewer({ container, ui: true, locale: "en" });
    await viewer.load(new Uint8Array(data), { fileName: "quarterly.pdf" });
    window.__demo = {
      viewer,
      caption(text) {
        caption.textContent = text;
      },
      settle: () =>
        new Promise((resolve) => setTimeout(resolve, 700)).then(
          () =>
            new Promise((resolve) =>
              requestAnimationFrame(() => requestAnimationFrame(resolve)),
            ),
        ),
      async pngBytes() {
        const canvas = document.createElement("canvas");
        canvas.width = 160;
        canvas.height = 100;
        const context = canvas.getContext("2d")!;
        const gradient = context.createLinearGradient(0, 0, 160, 100);
        gradient.addColorStop(0, "#2563eb");
        gradient.addColorStop(1, "#f97316");
        context.fillStyle = gradient;
        context.fillRect(0, 0, 160, 100);
        context.fillStyle = "rgba(255,255,255,0.9)";
        context.font = "bold 22px system-ui";
        context.fillText("PNG", 52, 60);
        const blob: Blob = await new Promise((resolve) =>
          canvas.toBlob((value) => resolve(value!), "image/png"),
        );
        return new Uint8Array(await blob.arrayBuffer());
      },
    };
  }, Array.from(bytes));
}

let frame = 0;

async function snap(page: Page, caption: string): Promise<void> {
  frame += 1;
  await page.evaluate((text) => window.__demo.caption(text), caption);
  await page.evaluate(() => window.__demo.settle());
  await page
    .locator("#demo")
    .screenshot({ path: `${OUT}/${String(frame).padStart(2, "0")}.png` });
}

test("pdf editing walk-through", async ({ page }) => {
  test.setTimeout(120_000);
  mkdirSync(OUT, { recursive: true });
  const original = await buildPdf([
    {
      texts: [
        { text: "Quarterly report", x: 72, y: 700, fontSize: 26 },
        { text: "Revenue grew 12% year over year.", x: 72, y: 660 },
        { text: "Costs stayed flat.", x: 72, y: 640 },
      ],
      rect: { x: 400, y: 620, width: 140, height: 90, fill: [226, 232, 240] },
    },
    { texts: [{ text: "Page two: regional breakdown", x: 72, y: 700 }] },
    { texts: [{ text: "Page three: outlook", x: 72, y: 700 }] },
  ]);
  await mount(page, original);
  await snap(
    page,
    "1 · Original PDF loaded; nothing editing-related fetched yet",
  );

  await page.evaluate(async () => {
    const { viewer } = window.__demo;
    viewer.setZoom(1.25);
    await viewer.search("Revenue");
  });
  await snap(page, "2 · Zoom 125 % and a search highlight before any edit");

  await page.evaluate(async () => {
    const { viewer } = window.__demo;
    const session = await viewer.edit();
    await session.insertTextBox({
      pageIndex: 0,
      rect: { x: 72, y: 180, width: 300, height: 60 },
      text: "Додано через insertTextBox: Cyrillic uses the bundled font",
      style: { fontSize: 14, bold: true, color: "#b91c1c" },
    });
  });
  await snap(
    page,
    "3 · insertTextBox applied: zoom kept, highlight cleared, page re-rendered",
  );

  await page.evaluate(async () => {
    const session = await window.__demo.viewer.edit();
    const found = await session.findText("Costs stayed flat.");
    await session.replaceText({
      target: found[0].elementIds[0],
      text: "Costs fell 3% on lower cloud spend.",
    });
    await session.setTextStyle({
      target: found[0].elementIds[0],
      style: { color: "#1d4ed8", fontSize: 14 },
    });
  });
  await snap(
    page,
    "4 · replaceText + setTextStyle on text that already existed in the file",
  );

  await page.evaluate(async () => {
    const session = await window.__demo.viewer.edit();
    await session.apply([
      {
        op: "insertShape",
        pageIndex: 0,
        shape: "rectangle",
        rect: { x: 400, y: 180, width: 140, height: 60 },
        fill: { color: "#fde68a" },
        stroke: { color: "#92400e", width: 1.5 },
      },
      {
        op: "insertShape",
        pageIndex: 0,
        shape: "ellipse",
        rect: { x: 400, y: 260, width: 140, height: 60 },
        fill: { color: "#bbf7d0" },
      },
      {
        op: "insertShape",
        pageIndex: 0,
        shape: "line",
        from: { x: 72, y: 260 },
        to: { x: 372, y: 260 },
        stroke: { color: "#000000", width: 2 },
      },
    ]);
  });
  await snap(
    page,
    "5 · One apply() with three insertShape operations = one undo step",
  );

  await page.evaluate(async () => {
    const session = await window.__demo.viewer.edit();
    await session.insertImage({
      pageIndex: 0,
      rect: { x: 72, y: 290, width: 160, height: 100 },
      data: await window.__demo.pngBytes(),
      mimeType: "image/png",
    });
  });
  await snap(page, "6 · insertImage: a PNG decoded in the worker, alpha kept");

  await page.evaluate(async () => {
    const session = await window.__demo.viewer.edit();
    const receipt = await session.insertTable({
      pageIndex: 0,
      at: { x: 72, y: 420 },
      width: 468,
      rows: [
        ["Region", "Q1", "Q2", "Change"],
        ["EMEA", "1.2M", "1.4M", "+17%"],
        ["Americas", "2.0M", "2.1M", "+5%"],
      ],
      columnWidths: [2, 1, 1, 1],
      style: { headerFill: "#e0e7ff", fontSize: 11 },
    });
    await session.setTableCell({
      target: receipt.createdIds[0],
      row: 2,
      column: 3,
      text: "+5% (revised)",
    });
  });
  await snap(
    page,
    "7 · insertTable with a header fill, then setTableCell relaid the table",
  );

  await page.evaluate(async () => {
    const session = await window.__demo.viewer.edit();
    const shapes = await session.getElements({
      pageIndex: 0,
      kinds: ["shape"],
    });
    const boxes = await session.getElements({
      pageIndex: 0,
      kinds: ["textBox"],
    });
    await session.apply([
      { op: "moveElement", target: boxes[0].id, by: { dx: 0, dy: 40 } },
      {
        op: "resizeElement",
        target: shapes.at(-3).id,
        rect: { x: 400, y: 180, width: 100, height: 100 },
      },
      { op: "deleteElement", target: shapes.at(-2).id },
    ]);
  });
  await snap(
    page,
    "8 · moveElement (text box down), resizeElement (rectangle), deleteElement (ellipse)",
  );

  await page.evaluate(async () => {
    const session = await window.__demo.viewer.edit();
    await session.undo();
  });
  await snap(
    page,
    "9 · undo(): the ellipse is back, the box and rectangle return",
  );

  await page.evaluate(async () => {
    const session = await window.__demo.viewer.edit();
    await session.redo();
  });
  await snap(
    page,
    "10 · redo(): the move, resize and delete are applied again",
  );

  await page.evaluate(async () => {
    const session = await window.__demo.viewer.edit();
    await session.apply([
      { op: "insertPage", index: 1 },
      { op: "rotatePage", pageIndex: 2, rotation: 90 },
    ]);
  });
  await snap(
    page,
    "11 · insertPage + rotatePage: the page counter and layout follow",
  );

  const saved = await page.evaluate(async () => {
    const session = await window.__demo.viewer.edit();
    const { bytes }: { bytes: Uint8Array } = await session.save();
    return { bytes: Array.from(bytes), revision: session.state.revision };
  });
  await mount(page, Uint8Array.from(saved.bytes));
  await snap(
    page,
    `12 · save() bytes (revision ${saved.revision}) loaded in a fresh viewer: everything survived`,
  );
});
