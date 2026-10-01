import { expect, test, type Page } from "@playwright/test";

import { buildPdf } from "../../packages/viewer/test/fixtures/pdf-builder.js";

/*
 * PDF editing through the public API against the real adapter: PDF.js renders,
 * the PDFium worker edits. Fixtures are built with PDFium in Node and handed
 * to the page as plain arrays.
 */

const EDIT_ASSETS = [
  "/workers/pdf-edit-worker.js",
  "/assets/pdfium/pdfium.wasm",
];

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
    const saved: Uint8Array = await session.save();
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
