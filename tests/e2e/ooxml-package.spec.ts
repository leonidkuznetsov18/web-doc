import { readFile, writeFile } from "node:fs/promises";

import { expect, test } from "@playwright/test";

import { OoxmlPackage } from "../../packages/viewer/src/edit/ooxml/package.js";
import { patches } from "../../packages/viewer/src/edit/ooxml/patch.js";
import { defaultResourceLimits } from "../../packages/viewer/src/limits.js";

/*
 * The OOXML package layer against the real viewer: a package patched by the
 * layer is loaded through the normal adapter and renders, every changed part
 * is well-formed for the browser's own parser, and the text the viewer
 * extracts carries the edit.
 */

const CORPUS = new URL("../../.cache/corpus/", import.meta.url);

async function patched(
  file: URL,
  partPattern: RegExp,
  textTag: string,
  replacement: string,
): Promise<{ bytes: number[]; changed: { name: string; text: string }[] }> {
  const original = new Uint8Array(await readFile(file));
  const pkg = await OoxmlPackage.open(original, {
    limits: defaultResourceLimits,
  });
  const target = pkg.partNames.find((name) => partPattern.test(name))!;
  const part = await pkg.xml(target);
  const node = part
    .findAll(textTag)
    .find((candidate) => part.textOf(candidate).trim().length > 0)!;
  const transaction = pkg.transaction();
  transaction.patch(part, [
    patches.replaceContent(part, node, patches.text(replacement)),
  ]);
  const change = await transaction.commit();
  const bytes = await pkg.save();
  const changed = await Promise.all(
    change.changedParts.map(async (name) => ({
      name,
      text: new TextDecoder().decode(await pkg.part(name)),
    })),
  );
  return { bytes: [...bytes], changed };
}

for (const [fileName, pattern, tag, replacement] of [
  ["sample.pptx", /\/slides\/slide1\.xml$/, "a:t", "Patched by web-doc"],
  ["sample.docx", /\/document\.xml$/, "w:t", "Patched by web-doc"],
] as const)
  test(`a ${fileName} patched by the package layer reopens in the viewer with the edit`, async ({
    page,
  }, testInfo) => {
    try {
      await page.exposeFunction(
        "__action938AttachFirstRaster",
        async (phase: unknown, metadata: unknown, png: unknown) => {
          const MAX_METADATA_BYTES = 32 * 1024;
          const MAX_ENCODED_PNG_BYTES = 16 * 1024 * 1024;
          const prefix = "data:image/png;base64,";
          const phases = ["before-page-text", "raster-copy", "after-page-text"];
          const keys = new Set([
            "provenance",
            "phase",
            "source",
            "width",
            "height",
            "dark",
            "redHistogram",
            "opaque",
            "transparent",
            "partialAlpha",
            "nonWhiteRGB",
            "fontStatus",
            "textPresent",
            "parserErrorCounts",
            "captureErrors",
          ]);
          const stages = new Set([
            "attachment-binding-unavailable",
            "attachment-call-failed",
            "capture-budget-exceeded",
            "incomplete-rgba-buffer",
            "copy-context-unavailable",
            "first-buffer-png-copy-failed",
            "first-buffer-metadata-capture-failed",
          ]);
          const count = (value: unknown) =>
            typeof value === "number" &&
            Number.isSafeInteger(value) &&
            value >= 0;
          const record = (value: unknown): value is Record<string, unknown> =>
            typeof value === "object" &&
            value !== null &&
            !Array.isArray(value);
          try {
            // Reject untrusted or excessive bridge data before allocating buffers.
            if (
              typeof phase !== "string" ||
              !phases.includes(phase) ||
              typeof metadata !== "string" ||
              metadata.length > MAX_METADATA_BYTES ||
              Buffer.byteLength(metadata, "utf8") > MAX_METADATA_BYTES
            ) {
              process.stderr.write("ACTION938 diagnostic input rejected\n");
              return;
            }
            let data: unknown;
            try {
              data = JSON.parse(metadata);
            } catch {
              process.stderr.write("ACTION938 diagnostic input rejected\n");
              return;
            }
            if (
              !record(data) ||
              Object.keys(data).some((key) => !keys.has(key)) ||
              data.provenance !== "originalFirstBuffer" ||
              data.phase !== phase ||
              !count(data.dark) ||
              typeof data.dark !== "number" ||
              data.dark > 50 ||
              ("source" in data && data.source !== "first-getImageData") ||
              [
                "width",
                "height",
                "opaque",
                "transparent",
                "partialAlpha",
                "nonWhiteRGB",
              ].some((key) => key in data && !count(data[key])) ||
              ("fontStatus" in data &&
                data.fontStatus !== "loaded" &&
                data.fontStatus !== "loading") ||
              ("textPresent" in data &&
                data.textPresent !== null &&
                typeof data.textPresent !== "boolean") ||
              ("redHistogram" in data &&
                (!Array.isArray(data.redHistogram) ||
                  data.redHistogram.length !== 256 ||
                  !data.redHistogram.every(
                    (value: unknown) =>
                      count(value) &&
                      typeof value === "number" &&
                      value <= 4_194_304,
                  ) ||
                  data.redHistogram.reduce(
                    (sum: number, value: unknown) =>
                      sum + (typeof value === "number" ? value : Infinity),
                    0,
                  ) > 4_194_304)) ||
              ("parserErrorCounts" in data &&
                (!Array.isArray(data.parserErrorCounts) ||
                  data.parserErrorCounts.length > 16 ||
                  !data.parserErrorCounts.every(count))) ||
              !Array.isArray(data.captureErrors) ||
              data.captureErrors.length > 16 ||
              !data.captureErrors.every(
                (stage: unknown) =>
                  typeof stage === "string" && stages.has(stage),
              ) ||
              (png !== null &&
                (phase !== "raster-copy" ||
                  typeof png !== "string" ||
                  png.length > MAX_ENCODED_PNG_BYTES ||
                  !png.startsWith(`${prefix}iVBORw0KGgo`) ||
                  (png.length - prefix.length) % 4 !== 0 ||
                  !/^[A-Za-z0-9+/]*={0,2}$/.test(png.slice(prefix.length))))
            ) {
              process.stderr.write("ACTION938 diagnostic input rejected\n");
              return;
            }
            const name =
              phase === "before-page-text"
                ? "action938-before-page-text"
                : phase === "raster-copy"
                  ? "action938-raster-copy"
                  : "action938-after-page-text";
            const metadataPath = testInfo.outputPath(`${name}.json`);
            await writeFile(metadataPath, metadata, "utf8");
            try {
              await testInfo.attach(`${name}.json`, {
                path: metadataPath,
                contentType: "application/json",
              });
            } catch {
              process.stderr.write("ACTION938 metadata attachment failed\n");
            }
            if (typeof png === "string") {
              const pngPath = testInfo.outputPath(
                "action938-original-first-buffer.png",
              );
              await writeFile(
                pngPath,
                Buffer.from(png.slice(prefix.length), "base64"),
              );
              try {
                await testInfo.attach("action938-original-first-buffer.png", {
                  path: pngPath,
                  contentType: "image/png",
                });
              } catch {
                process.stderr.write("ACTION938 PNG attachment failed\n");
              }
            }
          } catch {
            try {
              const errorPath = testInfo.outputPath(
                "action938-capture-error.json",
              );
              await writeFile(
                errorPath,
                JSON.stringify({ stage: "diagnostic-persistence-failed" }),
                "utf8",
              );
              await testInfo.attach("action938-capture-error.json", {
                path: errorPath,
                contentType: "application/json",
              });
            } catch {
              process.stderr.write("ACTION938 diagnostic persistence failed\n");
            }
          }
        },
      );
    } catch {
      console.warn("ACTION938 diagnostic binding registration failed");
    }
    const { bytes, changed } = await patched(
      new URL(fileName, CORPUS),
      pattern,
      tag,
      replacement,
    );
    expect(changed).toHaveLength(1);
    await page.goto("/");
    const result = await page.evaluate(
      async ({ bytes, changed, fileName, replacement }) => {
        const { ViewerClient } = (await import("/main.js")) as any;
        const client = ViewerClient.create({
          assetBaseUrl: new URL("/", location.href),
        });
        const viewer = client.createViewer();
        await viewer.load(new Uint8Array(bytes), { fileName });
        const canvas = document.createElement("canvas");
        await viewer.renderPage(0, canvas, { zoom: 1, devicePixelRatio: 1 });
        const firstImageData = canvas
          .getContext("2d")!
          .getImageData(0, 0, canvas.width, canvas.height);
        const pixels = firstImageData.data;
        let dark = 0;
        for (let offset = 0; offset < pixels.length; offset += 4)
          if (pixels[offset]! < 128) dark += 1;
        const parserErrors = changed.map(({ name, text }) => {
          const document = new DOMParser().parseFromString(
            text,
            "application/xml",
          );
          return [
            name,
            document.getElementsByTagName("parsererror").length,
          ] as const;
        });
        // A success produces no histogram or PNG; no later native readback is used.
        const captureErrors: string[] = [];
        const attachFailure = async (
          phase: "before-page-text" | "raster-copy" | "after-page-text",
          metadata: string,
          png: string | null,
        ) => {
          try {
            const attach: unknown = Reflect.get(
              window,
              "__action938AttachFirstRaster",
            );
            if (typeof attach !== "function") {
              captureErrors.push("attachment-binding-unavailable");
              return;
            }
            await attach(phase, metadata, png);
          } catch {
            captureErrors.push("attachment-call-failed");
            console.warn("ACTION938 diagnostic attachment failed");
          }
        };
        if (dark <= 50) {
          const MAX_CAPTURE_PIXELS = 4_194_304;
          const MAX_ENCODED_PNG_BYTES = 16 * 1024 * 1024;
          const basic = {
            provenance: "originalFirstBuffer",
            source: "first-getImageData",
            width: firstImageData.width,
            height: firstImageData.height,
            dark,
            textPresent: null,
            parserErrorCounts: parserErrors.map(([, errors]) => errors),
          };
          try {
            if (
              firstImageData.width * firstImageData.height >
                MAX_CAPTURE_PIXELS ||
              pixels.length > MAX_CAPTURE_PIXELS * 4
            ) {
              captureErrors.push("capture-budget-exceeded");
              await attachFailure(
                "before-page-text",
                JSON.stringify({
                  ...basic,
                  phase: "before-page-text",
                  captureErrors,
                }),
                null,
              );
            } else {
              const redHistogram = Array.from({ length: 256 }, () => 0);
              let opaque = 0;
              let transparent = 0;
              let partialAlpha = 0;
              let nonWhiteRGB = 0;
              for (let offset = 0; offset < pixels.length; offset += 4) {
                const red = pixels[offset],
                  green = pixels[offset + 1],
                  blue = pixels[offset + 2],
                  alpha = pixels[offset + 3];
                if (
                  red === undefined ||
                  green === undefined ||
                  blue === undefined ||
                  alpha === undefined
                ) {
                  captureErrors.push("incomplete-rgba-buffer");
                  break;
                }
                redHistogram[red] = (redHistogram[red] ?? 0) + 1;
                if (alpha === 255) opaque += 1;
                else if (alpha === 0) transparent += 1;
                else partialAlpha += 1;
                if (red !== 255 || green !== 255 || blue !== 255)
                  nonWhiteRGB += 1;
              }
              // Snapshot only; global font status does not prove individual faces.
              const fontStatus = document.fonts.status;
              // Durable metadata precedes PNG allocation/encoding and text extraction.
              await attachFailure(
                "before-page-text",
                JSON.stringify({
                  ...basic,
                  phase: "before-page-text",
                  redHistogram,
                  opaque,
                  transparent,
                  partialAlpha,
                  nonWhiteRGB,
                  fontStatus,
                  captureErrors,
                }),
                null,
              );
              let png: string | null = null;
              try {
                const copyCanvas = document.createElement("canvas");
                copyCanvas.width = firstImageData.width;
                copyCanvas.height = firstImageData.height;
                const copyContext = copyCanvas.getContext("2d");
                if (!copyContext)
                  captureErrors.push("copy-context-unavailable");
                else {
                  copyContext.putImageData(
                    new ImageData(
                      new Uint8ClampedArray(pixels),
                      firstImageData.width,
                      firstImageData.height,
                    ),
                    0,
                    0,
                  );
                  png = copyCanvas.toDataURL("image/png");
                  if (png.length > MAX_ENCODED_PNG_BYTES) {
                    png = null;
                    captureErrors.push("capture-budget-exceeded");
                  }
                }
              } catch {
                captureErrors.push("first-buffer-png-copy-failed");
              }
              await attachFailure(
                "raster-copy",
                JSON.stringify({
                  ...basic,
                  phase: "raster-copy",
                  captureErrors,
                }),
                png,
              );
            }
          } catch {
            captureErrors.push("first-buffer-metadata-capture-failed");
            await attachFailure(
              "before-page-text",
              JSON.stringify({
                ...basic,
                phase: "before-page-text",
                captureErrors,
              }),
              null,
            );
          }
        }
        const text: string = await viewer.getPageText(0);
        if (dark <= 50) {
          await attachFailure(
            "after-page-text",
            JSON.stringify({
              provenance: "originalFirstBuffer",
              phase: "after-page-text",
              dark,
              textPresent:
                typeof text === "string" && text.includes(replacement),
              parserErrorCounts: parserErrors.map(([, errors]) => errors),
              captureErrors,
            }),
            null,
          );
        }
        return { text, dark, parserErrors };
      },
      { bytes, changed, fileName, replacement },
    );
    expect(result.dark).toBeGreaterThan(50);
    expect(result.text).toContain(replacement);
    for (const [, errors] of result.parserErrors) expect(errors).toBe(0);
  });
