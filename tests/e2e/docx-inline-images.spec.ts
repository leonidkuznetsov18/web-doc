import { readFile } from "node:fs/promises";
import { expect, test } from "@playwright/test";

const fixture = new URL(
  "../fixtures/docx/oversized-inline-image.docx",
  import.meta.url,
);

test("DOCX fits oversized inline images within the page content", async ({
  page,
}, testInfo) => {
  const bytes = [...(await readFile(fixture))];
  await page.goto("/");
  const result = await page.evaluate(async (bytes) => {
    const moduleUrl = "/main.js";
    const { ViewerClient }: typeof import("web-doc") = await import(moduleUrl);
    const client = ViewerClient.create({
      assetBaseUrl: new URL("/", location.href),
    });
    const viewer = client.createViewer();
    const canvas = document.createElement("canvas");
    document.body.append(canvas);
    const context = canvas.getContext("2d");
    if (!context) throw new Error("Canvas 2D is unavailable");
    try {
      await viewer.load(new Uint8Array(bytes), {
        fileName: "oversized-inline-image.docx",
      });
      await viewer.renderPage(0, canvas);
      const { data, width, height } = context.getImageData(
        0,
        0,
        canvas.width,
        canvas.height,
      );
      let left = width;
      let top = height;
      let right = -1;
      let bottom = -1;
      for (let y = 0; y < height; y++) {
        for (let x = 0; x < width; x++) {
          const offset = (y * width + x) * 4;
          if (
            data[offset] > 240 &&
            data[offset + 1] < 10 &&
            data[offset + 2] < 10
          ) {
            left = Math.min(left, x);
            top = Math.min(top, y);
            right = Math.max(right, x);
            bottom = Math.max(bottom, y);
          }
        }
      }
      return { pageWidth: width, left, top, right, bottom };
    } finally {
      await viewer.destroy();
      await client.destroy();
    }
  }, bytes);

  // The fixture has a letter-size page, one-inch margins, and a 4:1 image.
  // Check the actual raster so a stale or detached parsed model cannot pass.
  const inch = result.pageWidth / 8.5;
  const imageWidth = result.right - result.left + 1;
  const imageHeight = result.bottom - result.top + 1;
  expect(imageWidth).toBeGreaterThan(0);
  expect.soft(Math.abs(result.left - inch)).toBeLessThanOrEqual(2);
  expect.soft(Math.abs(imageWidth - 6.5 * inch)).toBeLessThanOrEqual(2);
  expect(Math.abs(imageHeight - imageWidth / 4)).toBeLessThanOrEqual(2);
  await page.screenshot({ path: testInfo.outputPath("inline-image.png") });
});
