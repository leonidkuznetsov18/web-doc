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
        await session.end();
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
