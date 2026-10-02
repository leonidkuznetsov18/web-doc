import { readFile } from "node:fs/promises";

import { expect, test, type Page } from "@playwright/test";

import { prepareDocxForDisplay } from "../../packages/viewer/src/adapters/docx-prepass.js";
import { OoxmlPackage } from "../../packages/viewer/src/edit/ooxml/package.js";
import { defaultResourceLimits } from "../../packages/viewer/src/limits.js";
import {
  buildDocx,
  paragraph,
  sectPr,
} from "../../packages/viewer/test/fixtures/docx-builder.js";
import { installDeterministicOfficeFonts } from "./deterministic-fonts.js";

/*
 * Regression for the DOCX engine upgrade (module 05). The viewer renders
 * DOCX through the 0.88 line of @silurus/ooxml after an XML pre-pass on the
 * bytes. For each fixture the spec paints the page three ways: through the
 * viewer, through the bare engine on the pre-passed bytes and through the
 * bare engine on the original bytes. The viewer and the engine must agree on
 * the pre-passed bytes to the fidelity gate's SSIM threshold; the pre-pass
 * must change what the engine draws for an oversized inline picture and
 * nothing for a corpus document; and every text run the engine reports must
 * resolve to the `_wd<id>` bookmark of its paragraph, the bridge the adapter
 * uses for paragraph ids, which the viewer's own runs then carry as
 * `paragraphId`. Load and render times are logged for the record. This spec
 * began as the upgrade's spike (T50) and kept its measurements.
 */

const CORPUS = new URL("../../.cache/corpus/", import.meta.url);
const FIXTURES = new URL("../fixtures/docx/", import.meta.url);
const ENGINE = "/vendor/ooxml/docx.mjs";
const WIDTH = 816;
/** The fidelity gate's threshold for the modern Office family. */
const SSIM_THRESHOLD = 0.94;

interface Rendering {
  readonly pageCount: number;
  readonly loadMs: number;
  readonly renderMs: number;
  readonly pixels: string;
  readonly width: number;
  readonly height: number;
}

/** A page's runs as the viewer reports them: text and paragraph id. */
interface ViewerRun {
  readonly text: string;
  readonly paragraphId: string | undefined;
}

async function renderWithViewer(
  page: Page,
  bytes: Uint8Array,
  fileName: string,
): Promise<Rendering & { readonly pages: readonly (readonly ViewerRun[])[] }> {
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
      const pageCount = viewer.state.pageCount as number;
      // The public path to the adapter's runs: a whole-page selection.
      const pages: ViewerRun[][] = [];
      for (let pageIndex = 0; pageIndex < pageCount; pageIndex += 1) {
        const selection = await viewer.selectText({
          startPageIndex: pageIndex,
          startOffset: 0,
          endPageIndex: pageIndex,
          endOffset: Number.MAX_SAFE_INTEGER,
        });
        pages.push(
          (selection.runs as { text: string; paragraphId?: string }[]).map(
            (run) => ({ text: run.text, paragraphId: run.paragraphId }),
          ),
        );
      }
      await viewer.destroy();
      await client.destroy();
      return {
        pageCount,
        loadMs,
        renderMs,
        pixels,
        width: canvas.width,
        height: canvas.height,
        pages,
      };
    },
    { data: Array.from(bytes), fileName, width: WIDTH },
  );
}

/** The ids the pre-pass bookmarks carry, from the display copy's main part. */
async function markedIds(display: Uint8Array): Promise<Set<string>> {
  const pkg = await OoxmlPackage.open(display, {
    limits: defaultResourceLimits,
  });
  const xml = new TextDecoder().decode(await pkg.part("/word/document.xml"));
  return new Set(
    [...xml.matchAll(/w:name="_wd([0-9A-Fa-f]{8})"/g)].map(
      (match) => match[1]!,
    ),
  );
}

async function renderWithEngine(
  page: Page,
  bytes: Uint8Array,
): Promise<
  Rendering & {
    readonly runs: number;
    readonly withParagraphId: number;
    readonly resolvedIds: number;
    readonly withSource: number;
    readonly firstRun: unknown;
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
      // The bridge the adapter uses: a run's story path names a model
      // paragraph; the pre-pass bookmark on it, or on the nearest earlier
      // paragraph of the same container (a page break splits a paragraph),
      // carries the id.
      const resolveId = (run: Record<string, unknown>): string | undefined => {
        const source = run.source as
          { story: string; path: number[] } | undefined;
        if (!source || source.story !== "body") return undefined;
        const model = document.document as { body: unknown[] };
        let container: unknown[] = model.body;
        const path = [...source.path];
        const last = path.pop()!;
        for (let depth = 0; depth < path.length; depth += 1) {
          const node = container[path[depth]!] as Record<string, unknown>;
          if (!node) return undefined;
          if (node.type === "table") {
            const row = (node.rows as { cells: { content: unknown[] }[] }[])[
              path[depth + 1]!
            ];
            const cell = row?.cells[path[depth + 2]!];
            container = cell?.content ?? [];
            depth += 2;
          } else if (Array.isArray(node.content))
            container = node.content as unknown[];
        }
        for (let index = last; index >= 0; index -= 1) {
          const item = container[index] as
            { type?: string; bookmarks?: string[] } | undefined;
          if (!item) return undefined;
          const name = item.bookmarks?.find((candidate) =>
            candidate.startsWith("_wd"),
          );
          if (name) return name.slice(3);
          if (item.type === "paragraph") return undefined;
        }
        return undefined;
      };
      const resolved = runs.filter(
        (run) => resolveId(run) !== undefined,
      ).length;
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
        resolvedIds: resolved,
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
        layoutComplete: document.layoutComplete !== false,
      };
      document.destroy();
      return result;
    },
    { data: Array.from(bytes), engine: ENGINE, width: WIDTH },
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

const sameSize = (a: Rendering, b: Rendering): boolean =>
  a.width === b.width && a.height === b.height;

for (const { name, url, scaledImages } of [
  { name: "sample.docx", url: new URL("sample.docx", CORPUS), scaledImages: 0 },
  {
    name: "oversized-inline-image.docx",
    url: new URL("oversized-inline-image.docx", FIXTURES),
    scaledImages: 1,
  },
] as const)
  test(`DOCX engine: ${name} through the viewer, the pre-pass and the bare engine`, async ({
    page,
  }, testInfo) => {
    test.setTimeout(120_000);
    const bytes = new Uint8Array(await readFile(url));
    await page.goto("/");
    await installDeterministicOfficeFonts(page);
    const viewer = await renderWithViewer(page, bytes, name);
    const display = await prepareDocxForDisplay(bytes, defaultResourceLimits);
    const prepassed = await renderWithEngine(page, display.bytes);
    const original = await renderWithEngine(page, bytes);
    const viewerToEngine = sameSize(viewer, prepassed)
      ? await ssim(page, viewer.pixels, prepassed.pixels)
      : 0;
    const prepassEffect = sameSize(original, prepassed)
      ? await ssim(page, original.pixels, prepassed.pixels)
      : 0;
    const report = {
      name,
      prepass: {
        scaledImages: display.scaledImages,
        markedParagraphs: display.markedParagraphs,
        generatedIds: display.generatedIds,
      },
      pages: { viewer: viewer.pageCount, engine: prepassed.pageCount },
      size: {
        viewer: [viewer.width, viewer.height],
        engine: [prepassed.width, prepassed.height],
      },
      loadMs: {
        viewer: Math.round(viewer.loadMs),
        engine: Math.round(prepassed.loadMs),
      },
      renderMs: {
        viewer: Math.round(viewer.renderMs),
        engine: Math.round(prepassed.renderMs),
      },
      ssim: {
        viewerToEngine: Number(viewerToEngine.toFixed(4)),
        originalToPrepassed: Number(prepassEffect.toFixed(4)),
      },
      runs: prepassed.runs,
      withParagraphId: prepassed.withParagraphId,
      resolvedIds: prepassed.resolvedIds,
      withSource: prepassed.withSource,
      firstRun: prepassed.firstRun,
      layoutComplete: prepassed.layoutComplete,
      viewerRuns: viewer.pages[0]?.length ?? 0,
      viewerParagraphs: new Set(
        viewer.pages[0]?.map((run) => run.paragraphId) ?? [],
      ).size,
    };
    console.log(`docx engine ${JSON.stringify(report)}`);
    await testInfo.attach(`${name}-viewer.png`, {
      body: Buffer.from(viewer.pixels.split(",")[1]!, "base64"),
      contentType: "image/png",
    });
    await testInfo.attach(`${name}-engine.png`, {
      body: Buffer.from(prepassed.pixels.split(",")[1]!, "base64"),
      contentType: "image/png",
    });
    expect(viewer.pageCount).toBe(prepassed.pageCount);
    expect(prepassed.pageCount).toBeGreaterThan(0);
    expect(display.scaledImages).toBe(scaledImages);
    expect(viewerToEngine).toBeGreaterThanOrEqual(SSIM_THRESHOLD);
    if (scaledImages > 0) expect(prepassEffect).toBeLessThan(SSIM_THRESHOLD);
    else expect(prepassEffect).toBeGreaterThanOrEqual(SSIM_THRESHOLD);
    expect(prepassed.resolvedIds).toBe(prepassed.runs);
    expect(prepassed.layoutComplete).toBe(true);
    // The adapter's runs carry the ids the pre-pass wrote, nothing else.
    const marked = await markedIds(display.bytes);
    const viewerRuns = viewer.pages[0] ?? [];
    expect(viewerRuns.length).toBe(prepassed.runs);
    for (const run of viewerRuns) {
      expect(run.paragraphId).toMatch(/^[0-9A-F]{8}$/);
      expect(marked.has(run.paragraphId!)).toBe(true);
    }
  });

test("DOCX engine: a paragraph that continues on the next page keeps one paragraph id", async ({
  page,
}) => {
  test.setTimeout(120_000);
  // A paragraph with a page break in the middle, a table and a trailing
  // paragraph; the engine splits the first around a hoisted page break.
  const bytes = buildDocx({
    body:
      `<w:p><w:r><w:t>Before the break</w:t></w:r><w:r><w:br w:type="page"/></w:r><w:r><w:t>After the break</w:t></w:r></w:p>` +
      `<w:tbl><w:tr><w:tc>${paragraph("In a cell")}</w:tc></w:tr></w:tbl>` +
      paragraph("Last paragraph") +
      sectPr(),
  });
  await page.goto("/");
  await installDeterministicOfficeFonts(page);
  const rendered = await renderWithViewer(page, bytes, "split.docx");
  const display = await prepareDocxForDisplay(bytes, defaultResourceLimits);
  const marked = await markedIds(display.bytes);
  const idsOf = (pageIndex: number, text: string): string[] =>
    (rendered.pages[pageIndex] ?? [])
      .filter((run) => run.text.includes(text))
      .map((run) => run.paragraphId ?? "");
  console.log(
    `docx engine split ${JSON.stringify({ pages: rendered.pages, marked: [...marked] })}`,
  );
  expect(rendered.pageCount).toBe(2);
  expect(marked.size).toBe(3);
  const before = idsOf(0, "Before");
  const after = idsOf(1, "After");
  const cell = idsOf(1, "cell");
  const last = idsOf(1, "Last");
  expect(before.length).toBeGreaterThan(0);
  expect(after.length).toBeGreaterThan(0);
  expect(new Set([...before, ...after]).size).toBe(1);
  expect(cell).not.toEqual([]);
  expect(last).not.toEqual([]);
  expect(new Set([before[0], cell[0], last[0]]).size).toBe(3);
  for (const run of rendered.pages.flat())
    expect(marked.has(run.paragraphId ?? "")).toBe(true);
});
