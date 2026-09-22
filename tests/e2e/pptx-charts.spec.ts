import { readFile } from "node:fs/promises";
import { expect, test } from "@playwright/test";

const fixture = new URL(
  "../fixtures/pptx/chart-point-colors.pptx",
  import.meta.url,
);

for (const scenario of [
  {
    name: "horizontal bars with an automatic axis",
    pageIndex: 0,
    horizontal: true,
    ticks: ["0", "0.5", "1", "1.5", "2", "2.5"],
  },
  {
    name: "columns with an authored axis interval",
    pageIndex: 1,
    horizontal: false,
    ticks: ["0", "0.75", "1.5", "2.25", "3"],
  },
]) {
  test(`PPTX preserves point colors and ticks for ${scenario.name}`, async ({
    page,
  }, testInfo) => {
    const bytes = [...(await readFile(fixture))];
    await page.goto("/");
    const result = await page.evaluate(
      async ({ bytes, pageIndex, horizontal }) => {
        const moduleUrl = "/main.js";
        const { ViewerClient }: typeof import("web-doc") = await import(
          moduleUrl
        );
        const client = ViewerClient.create({
          assetBaseUrl: new URL("/", location.href),
        });
        const viewer = client.createViewer();
        const canvas = document.createElement("canvas");
        document.body.append(canvas);
        const context = canvas.getContext("2d");
        if (!context) throw new Error("Canvas 2D is unavailable");
        const labels: string[] = [];
        const fillText = context.fillText.bind(context);
        context.fillText = (text, x, y, maxWidth) => {
          labels.push(text);
          fillText(text, x, y, maxWidth);
        };
        try {
          await viewer.load(new Uint8Array(bytes), {
            fileName: "chart-point-colors.pptx",
          });
          await viewer.renderPage(pageIndex, canvas);
          const { data, width, height } = context.getImageData(
            0,
            0,
            canvas.width,
            canvas.height,
          );
          // Find the colored bars in the actual raster, independent of plot margins.
          // At least ten solid pixels excludes anti-aliased text and axis strokes.
          const bands: string[] = [];
          let previous = "";
          for (let band = 0; band < (horizontal ? height : width); band++) {
            let gray = 0;
            let coral = 0;
            for (
              let offset = 0;
              offset < (horizontal ? width : height);
              offset++
            ) {
              const x = horizontal ? offset : band;
              const y = horizontal ? band : offset;
              const i = (y * width + x) * 4;
              if (data[i] === 107 && data[i + 1] === 107 && data[i + 2] === 123)
                gray++;
              if (data[i] === 234 && data[i + 1] === 91 && data[i + 2] === 79)
                coral++;
            }
            const color = gray >= 10 ? "gray" : coral >= 10 ? "coral" : "";
            if (color && color !== previous) bands.push(color);
            previous = color;
          }
          return {
            bands,
            ticks: labels.filter((label) => /^\d+(\.\d+)?$/.test(label)),
          };
        } finally {
          await viewer.destroy();
          await client.destroy();
        }
      },
      { bytes, pageIndex: scenario.pageIndex, horizontal: scenario.horizontal },
    );
    // Horizontal category axes place the first category at the bottom by default.
    expect
      .soft(result.bands)
      .toEqual(
        scenario.horizontal
          ? ["coral", "gray", "gray", "gray"]
          : ["gray", "gray", "gray", "coral"],
      );
    expect(result.ticks).toEqual(scenario.ticks);
    await page.screenshot({ path: testInfo.outputPath("chart.png") });
  });
}
