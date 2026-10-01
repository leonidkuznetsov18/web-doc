import { readFile } from "node:fs/promises";

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
  }) => {
    const { bytes, changed } = await patched(
      new URL(fileName, CORPUS),
      pattern,
      tag,
      replacement,
    );
    expect(changed).toHaveLength(1);
    await page.goto("/");
    const result = await page.evaluate(
      async ({ bytes, changed, fileName }) => {
        const { ViewerClient } = (await import("/main.js")) as any;
        const client = ViewerClient.create({
          assetBaseUrl: new URL("/", location.href),
        });
        const viewer = client.createViewer();
        await viewer.load(new Uint8Array(bytes), { fileName });
        const canvas = document.createElement("canvas");
        await viewer.renderPage(0, canvas, { zoom: 1, devicePixelRatio: 1 });
        const pixels = canvas
          .getContext("2d")!
          .getImageData(0, 0, canvas.width, canvas.height).data;
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
        return {
          text: await viewer.getPageText(0),
          dark,
          parserErrors,
        };
      },
      { bytes, changed, fileName },
    );
    expect(result.dark).toBeGreaterThan(50);
    expect(result.text).toContain(replacement);
    for (const [, errors] of result.parserErrors) expect(errors).toBe(0);
  });
