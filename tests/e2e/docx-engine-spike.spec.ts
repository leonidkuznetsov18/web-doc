import { readFile } from "node:fs/promises";

import { expect, test, type Page } from "@playwright/test";

import { installDeterministicOfficeFonts } from "./deterministic-fonts.js";

/*
 * Spike for the DOCX engine upgrade (module 05): the 0.88 line of
 * @silurus/ooxml, vendored by the example build as the PPTX engine, renders
 * the DOCX fixtures next to the 0.72.2 engine the viewer ships. The spec
 * reports the structural similarity of the two renderings, what the 0.88
 * text runs carry (w14:paraId, story source), the load and layout cost,
 * and whether the inline-image fitting still has a model to adjust. It
 * records numbers; it fails only when 0.88 cannot open a fixture at all.
 */

const CORPUS = new URL("../../.cache/corpus/", import.meta.url);
const FIXTURES = new URL("../fixtures/docx/", import.meta.url);
const NEW_ENGINE = "/vendor/ooxml-pptx/docx.mjs";
const WIDTH = 816;

interface Rendering {
  readonly pageCount: number;
  readonly loadMs: number;
  readonly renderMs: number;
  readonly pixels: string;
  readonly width: number;
  readonly height: number;
}

async function renderWithViewer(
  page: Page,
  bytes: Uint8Array,
  fileName: string,
): Promise<Rendering> {
  return page.evaluate(
    async ({ data, fileName, width }) => {
      const { ViewerClient } = (await import("/main.js")) as any;
      const client = ViewerClient.create({
        assetBaseUrl: new URL("/", location.href),
        fontPolicy: { mode: "offline" },
      });
      const viewer = client.createViewer();
      const started = performance.now();
      await viewer.load(new Uint8Array(data), { fileName });
      const loadMs = performance.now() - started;
      const canvas = document.createElement("canvas");
      const renderStarted = performance.now();
      await viewer.renderPage(0, canvas, {
        zoom: 1,
        devicePixelRatio: 1,
        width,
      });
      const renderMs = performance.now() - renderStarted;
      const pixels = canvas.toDataURL("image/png");
      const pageCount = viewer.state.pageCount;
      await viewer.destroy();
      await client.destroy();
      return {
        pageCount,
        loadMs,
        renderMs,
        pixels,
        width: canvas.width,
        height: canvas.height,
      };
    },
    { data: Array.from(bytes), fileName, width: WIDTH },
  );
}

async function renderWithNewEngine(
  page: Page,
  bytes: Uint8Array,
): Promise<
  Rendering & {
    readonly runs: number;
    readonly withParagraphId: number;
    readonly withSource: number;
    readonly firstRun: unknown;
    readonly hasDocumentModel: boolean;
    readonly layoutComplete: boolean;
  }
> {
  return page.evaluate(
    async ({ data, engine, width }) => {
      const { DocxDocument } = (await import(engine)) as any;
      const started = performance.now();
      const document = await DocxDocument.load(new Uint8Array(data).buffer, {
        useGoogleFonts: false,
        maxZipEntryBytes: 64 * 1024 * 1024,
        mode: "main",
      });
      const loadMs = performance.now() - started;
      const canvas = window.document.createElement("canvas");
      const renderStarted = performance.now();
      await document.renderPage(canvas, 0, { width, dpr: 1 });
      const renderMs = performance.now() - renderStarted;
      const runs = (await document.collectPageRuns(0, {
        width,
        dpr: 1,
      })) as Record<string, unknown>[];
      let hasDocumentModel = false;
      try {
        hasDocumentModel =
          typeof document.document === "object" && document.document !== null;
      } catch {
        hasDocumentModel = false;
      }
      const result = {
        pageCount: document.pageCount as number,
        loadMs,
        renderMs,
        pixels: canvas.toDataURL("image/png"),
        width: canvas.width,
        height: canvas.height,
        runs: runs.length,
        withParagraphId: runs.filter(
          (run) => typeof run.paragraphId === "string",
        ).length,
        withSource: runs.filter((run) => run.source !== undefined).length,
        firstRun: runs[0]
          ? Object.fromEntries(
              Object.entries(runs[0]).filter(([key]) =>
                [
                  "text",
                  "paragraphId",
                  "sourceRunIndex",
                  "source",
                  "font",
                  "fontSize",
                  "direction",
                ].includes(key),
              ),
            )
          : undefined,
        hasDocumentModel,
        layoutComplete: document.layoutComplete !== false,
      };
      document.destroy();
      return result;
    },
    { data: Array.from(bytes), engine: NEW_ENGINE, width: WIDTH },
  );
}

/** Structural similarity of two PNG renderings of the same size, as the fidelity gate computes it. */
async function ssim(page: Page, a: string, b: string): Promise<number> {
  return page.evaluate(
    async ({ a, b }) => {
      const decode = async (url: string) => {
        const image = new Image();
        image.src = url;
        await image.decode();
        const canvas = document.createElement("canvas");
        canvas.width = image.width;
        canvas.height = image.height;
        const context = canvas.getContext("2d")!;
        context.drawImage(image, 0, 0);
        return context.getImageData(0, 0, canvas.width, canvas.height);
      };
      const [left, right] = await Promise.all([decode(a), decode(b)]);
      if (left.width !== right.width || left.height !== right.height) return 0;
      const count = left.width * left.height;
      const luminance = (data: Uint8ClampedArray, offset: number) =>
        0.2126 * data[offset]! +
        0.7152 * data[offset + 1]! +
        0.0722 * data[offset + 2]!;
      let leftMean = 0;
      let rightMean = 0;
      for (let pixel = 0; pixel < count; pixel += 1) {
        leftMean += luminance(left.data, pixel * 4);
        rightMean += luminance(right.data, pixel * 4);
      }
      leftMean /= count;
      rightMean /= count;
      let leftVariance = 0;
      let rightVariance = 0;
      let covariance = 0;
      for (let pixel = 0; pixel < count; pixel += 1) {
        const dl = luminance(left.data, pixel * 4) - leftMean;
        const dr = luminance(right.data, pixel * 4) - rightMean;
        leftVariance += dl * dl;
        rightVariance += dr * dr;
        covariance += dl * dr;
      }
      const denominator = Math.max(1, count - 1);
      leftVariance /= denominator;
      rightVariance /= denominator;
      covariance /= denominator;
      const c1 = (0.01 * 255) ** 2;
      const c2 = (0.03 * 255) ** 2;
      return (
        ((2 * leftMean * rightMean + c1) * (2 * covariance + c2)) /
        ((leftMean ** 2 + rightMean ** 2 + c1) *
          (leftVariance + rightVariance + c2))
      );
    },
    { a, b },
  );
}

for (const [name, url] of [
  ["sample.docx", new URL("sample.docx", CORPUS)],
  [
    "oversized-inline-image.docx",
    new URL("oversized-inline-image.docx", FIXTURES),
  ],
] as const)
  test(`spike: ${name} through the 0.88 DOCX engine next to the shipped 0.72.2`, async ({
    page,
  }, testInfo) => {
    test.setTimeout(120_000);
    const bytes = new Uint8Array(await readFile(url));
    await page.goto("/");
    await installDeterministicOfficeFonts(page);
    const current = await renderWithViewer(page, bytes, name);
    const next = await renderWithNewEngine(page, bytes);
    const similarity =
      current.width === next.width && current.height === next.height
        ? await ssim(page, current.pixels, next.pixels)
        : 0;
    const report = {
      name,
      pages: { current: current.pageCount, next: next.pageCount },
      size: {
        current: [current.width, current.height],
        next: [next.width, next.height],
      },
      loadMs: {
        current: Math.round(current.loadMs),
        next: Math.round(next.loadMs),
      },
      renderMs: {
        current: Math.round(current.renderMs),
        next: Math.round(next.renderMs),
      },
      ssim: Number(similarity.toFixed(4)),
      runs: next.runs,
      withParagraphId: next.withParagraphId,
      withSource: next.withSource,
      firstRun: next.firstRun,
      hasDocumentModel: next.hasDocumentModel,
      layoutComplete: next.layoutComplete,
    };
    console.log(`docx engine spike ${JSON.stringify(report)}`);
    await testInfo.attach(`${name}-0.72.png`, {
      body: Buffer.from(current.pixels.split(",")[1]!, "base64"),
      contentType: "image/png",
    });
    await testInfo.attach(`${name}-0.88.png`, {
      body: Buffer.from(next.pixels.split(",")[1]!, "base64"),
      contentType: "image/png",
    });
    expect(next.pageCount).toBeGreaterThan(0);
    expect(next.runs).toBeGreaterThanOrEqual(0);
  });
