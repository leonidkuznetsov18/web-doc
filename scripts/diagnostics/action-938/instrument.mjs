import { createHash } from "node:crypto";
import { readFile, writeFile } from "node:fs/promises";
import { resolve } from "node:path";

const root = resolve(process.argv[2] ?? ".");
const hashes = [];

async function instrument(path, replacements) {
  const file = resolve(root, path);
  const original = await readFile(file, "utf8");
  let updated = original;
  for (const [before, after] of replacements) {
    if (updated.split(before).length !== 2)
      throw new Error(`Expected exactly one diagnostic boundary in ${path}`);
    updated = updated.replace(before, after);
  }
  let reversed = updated;
  for (const [before, after] of [...replacements].reverse())
    reversed = reversed.replace(after, before);
  if (reversed !== original)
    throw new Error(`Diagnostics changed original source in ${path}`);
  hashes.push({
    path,
    originalSHA256: createHash("sha256").update(original).digest("hex"),
    instrumentedSHA256: createHash("sha256").update(updated).digest("hex"),
  });
  await writeFile(file, updated);
}

await instrument("tests/e2e/ooxml-package.spec.ts", [
  [
    '    await page.goto("/");',
    `    const rendererEvents: string[] = [];
    page.on("console", (message) => {
      if (message.text().startsWith("ACTION938 "))
        rendererEvents.push(message.text());
    });
    await page.goto("/");`,
  ],
  [
    "        const viewer = client.createViewer();",
    `        const viewer = client.createViewer();
        const warnings: unknown[] = [];
        viewer.on("warning", (warning: unknown) => warnings.push(warning));`,
  ],
  [
    '        const canvas = document.createElement("canvas");',
    `        const canvas = document.createElement("canvas");
        const fontsBefore = document.fonts.status;
        const renderStarted = performance.now();`,
  ],
  [
    "        const parserErrors = changed.map(({ name, text }) => {",
    `        const histogram = new Array<number>(256).fill(0);
        let opaque = 0;
        let nonWhite = 0;
        for (let offset = 0; offset < pixels.length; offset += 4) {
          const red = pixels[offset];
          if (red === undefined) throw new Error("Missing red pixel channel");
          histogram[red] = (histogram[red] ?? 0) + 1;
          if (pixels[offset + 3] === 255) opaque += 1;
          if (red < 255 || pixels[offset + 1] !== 255 || pixels[offset + 2] !== 255)
            nonWhite += 1;
        }
        const diagnostics = {
          width: canvas.width,
          height: canvas.height,
          connected: canvas.isConnected,
          dark,
          opaque,
          nonWhite,
          histogram,
          fontsBefore,
          fontsAfter: document.fonts.status,
          renderAndReadbackMs: performance.now() - renderStarted,
          warnings,
          png: canvas.toDataURL("image/png"),
        };
        const parserErrors = changed.map(({ name, text }) => {`,
  ],
  [
    "          parserErrors,",
    "          parserErrors,\n          diagnostics,",
  ],
  [
    "    expect(result.dark).toBeGreaterThan(50);",
    `    const { png, ...diagnostics } = result.diagnostics;
    await test.info().attach("direct-render-pixels", {
      body: Buffer.from(png.slice(png.indexOf(",") + 1), "base64"),
      contentType: "image/png",
    });
    await test.info().attach("direct-render-diagnostics", {
      body: JSON.stringify({ fileName, ...diagnostics, rendererEvents,
        text: result.text, parserErrors: result.parserErrors, changed }, null, 2),
      contentType: "application/json",
    });
    expect(result.dark).toBeGreaterThan(50);`,
  ],
]);

const boundary = `        await handle.backend.renderSlide(target, viewport.pageIndex, {
          width: (handle.backend.slideWidth / 9525) * viewport.zoom,
          dpr: viewport.devicePixelRatio,
        });`;
await instrument("packages/viewer/src/adapters/office.ts", [
  [
    boundary,
    `        console.debug("ACTION938 renderSlide-start", performance.now());
${boundary}
        console.debug("ACTION938 renderSlide-complete", performance.now());`,
  ],
]);
await writeFile(
  resolve(root, "action-938-source-hashes.json"),
  JSON.stringify(hashes, null, 2),
);
