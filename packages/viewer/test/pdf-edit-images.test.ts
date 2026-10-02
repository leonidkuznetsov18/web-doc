import assert from "node:assert/strict";
import { before, describe, it } from "node:test";

import { createPdfEditHandler } from "../src/edit/pdf/engine/handler.js";
import { jpegSize, pngSize } from "../src/edit/pdf/engine/images.js";
import {
  loadPdfEditEngine,
  type PdfEditEngineClient,
} from "../src/edit/pdf/provider.js";
import type { PageRect, PdfOperation } from "../src/index.js";
import { defaultResourceLimits, ViewerError } from "../src/index.js";
import { loopbackWorker } from "./fixtures/loopback-worker.js";
import { buildPdf, fixturePdfium } from "./fixtures/pdf-builder.js";
import { decodePng, samplePng } from "./fixtures/png.js";
import { tinyJpeg } from "./fixtures/tiny-jpeg.js";

const signal = new AbortController().signal;
const op = <T extends PdfOperation>(operation: T): T => operation;

async function engineFor(
  original: Uint8Array,
  limits = defaultResourceLimits,
): Promise<PdfEditEngineClient> {
  const pair = loopbackWorker(
    createPdfEditHandler({
      loadPdfium: () => fixturePdfium(),
      fetchBytes: async () => {
        throw new Error("no fonts");
      },
      decodeImage: async (bytes) => decodePng(bytes),
    }),
  );
  return loadPdfEditEngine(
    original,
    { format: "pdf", limits, signal },
    { createWorker: () => pair.worker },
  );
}

function near(actual: PageRect, expected: PageRect, slack: number): void {
  for (const key of ["x", "y", "width", "height"] as const)
    assert.ok(
      Math.abs(actual[key] - expected[key]) <= slack,
      `${key}: ${actual[key]} vs ${expected[key]} in ${JSON.stringify(actual)}`,
    );
}

/** Pixel size PDFium reports for the first image object on a page. */
async function imagePixelSize(bytes: Uint8Array, pageIndex: number) {
  const pdfium = await fixturePdfium();
  const { lib } = pdfium;
  const document = pdfium.openDocument(bytes);
  try {
    const page = lib.FPDF_LoadPage(document.handle, pageIndex);
    try {
      for (let index = 0; index < lib.FPDFPage_CountObjects(page); index += 1) {
        const object = lib.FPDFPage_GetObject(page, index);
        if (lib.FPDFPageObj_GetType(object) !== 3) continue;
        return pdfium.readNumbers(2, "i32", ([w, h]) =>
          lib.FPDFImageObj_GetImagePixelSize(object, w!, h!),
        );
      }
      return undefined;
    } finally {
      lib.FPDF_ClosePage(page);
    }
  } finally {
    document.close();
  }
}

describe("insertImage", () => {
  let original: Uint8Array;

  before(async () => {
    original = await buildPdf([
      "One",
      { width: 300, height: 400, rotation: 1 },
    ]);
  });

  it("reads image headers", () => {
    assert.deepEqual(jpegSize(tinyJpeg()), { width: 16, height: 8 });
    assert.deepEqual(pngSize(samplePng(24, 16)), { width: 24, height: 16 });
    assert.equal(jpegSize(samplePng()), undefined);
    assert.equal(pngSize(tinyJpeg()), undefined);
    assert.equal(jpegSize(new Uint8Array([0xff, 0xd8, 0xff])), undefined);
  });

  it("embeds JPEG data as it is and PNG pixels losslessly, upright on rotated pages", async () => {
    const engine = await engineFor(original);
    try {
      const jpeg = tinyJpeg();
      const png = samplePng();
      const change = await engine.apply(
        [
          op({
            op: "insertImage",
            pageIndex: 0,
            rect: { x: 100, y: 100, width: 160, height: 80 },
            data: jpeg,
            mimeType: "image/jpeg",
          }),
          op({
            op: "insertImage",
            pageIndex: 0,
            rect: { x: 300, y: 300, width: 48, height: 32 },
            data: png,
            mimeType: "image/png",
          }),
          op({
            op: "insertImage",
            pageIndex: 1,
            rect: { x: 20, y: 30, width: 100, height: 50 },
            data: jpeg,
            mimeType: "image/jpeg",
          }),
        ],
        signal,
      );
      assert.deepEqual(change.createdIds, [
        "p0:n1.0.0",
        "p0:n1.1.0",
        "p1:n1.2.0",
      ]);
      const elements = await engine.getElements({ kinds: ["image"] }, signal);
      near(
        elements[0]!.bounds,
        { x: 100, y: 100, width: 160, height: 80 },
        0.5,
      );
      near(elements[1]!.bounds, { x: 300, y: 300, width: 48, height: 32 }, 0.5);
      near(elements[2]!.bounds, { x: 20, y: 30, width: 100, height: 50 }, 0.5);

      const saved = await engine.materialize(signal);
      // The JPEG stream is stored verbatim: its entropy-coded tail is in the file.
      const tail = jpeg.subarray(jpeg.length - 64);
      assert.ok(indexOfBytes(saved, tail) >= 0, "JPEG bytes present verbatim");
      assert.deepEqual(await imagePixelSize(saved, 1), [16, 8]);
    } finally {
      await engine.dispose();
    }
  });

  it("stores PNG alpha and reports the decoded size", async () => {
    const engine = await engineFor(original);
    try {
      await engine.apply(
        [
          op({
            op: "insertImage",
            pageIndex: 0,
            rect: { x: 50, y: 50, width: 240, height: 160 },
            data: samplePng(24, 16),
            mimeType: "image/png",
          }),
        ],
        signal,
      );
      const saved = await engine.materialize(signal);
      assert.deepEqual(await imagePixelSize(saved, 0), [24, 16]);
      assert.ok(
        indexOfBytes(saved, new TextEncoder().encode("/SMask")) >= 0,
        "alpha kept as a soft mask",
      );
    } finally {
      await engine.dispose();
    }
  });

  it("rejects unreadable or oversized data", async () => {
    const engine = await engineFor(original, {
      ...defaultResourceLimits,
      maxDecodedPixels: 100,
    });
    try {
      const issues = await engine.validate(
        [
          op({
            op: "insertImage",
            pageIndex: 0,
            rect: { x: 0, y: 0, width: 10, height: 10 },
            data: new Uint8Array([1, 2, 3, 4, 5]),
            mimeType: "image/png",
          }),
          op({
            op: "insertImage",
            pageIndex: 0,
            rect: { x: 0, y: 0, width: 10, height: 10 },
            data: samplePng(4, 4),
            mimeType: "image/jpeg",
          }),
          op({
            op: "insertImage",
            pageIndex: 5,
            rect: { x: 0, y: 0, width: 10, height: 10 },
            data: tinyJpeg(),
            mimeType: "image/jpeg",
          }),
        ],
        signal,
      );
      assert.deepEqual(
        issues.map(
          (issue) => `${issue.operationIndex}${issue.path}:${issue.code}`,
        ),
        [
          "0/data:invalid-data",
          "1/data:invalid-data",
          "2/pageIndex:unknown-target",
        ],
      );
      await assert.rejects(
        engine.validate(
          [
            op({
              op: "insertImage",
              pageIndex: 0,
              rect: { x: 0, y: 0, width: 10, height: 10 },
              data: samplePng(24, 16),
              mimeType: "image/png",
            }),
          ],
          signal,
        ),
        (error: unknown) =>
          error instanceof ViewerError && error.code === "resource-limit",
      );
    } finally {
      await engine.dispose();
    }
  });

  it("replays image insertions deterministically", async () => {
    const batch = [
      op({
        op: "insertImage",
        pageIndex: 0,
        rect: { x: 10, y: 10, width: 48, height: 32 },
        data: samplePng(),
        mimeType: "image/png",
      }),
    ];
    const first = await engineFor(original);
    const second = await engineFor(original);
    try {
      await first.apply(batch, signal);
      await second.restore([batch], signal);
      assert.deepEqual(
        await second.materialize(signal),
        await first.materialize(signal),
      );
    } finally {
      await first.dispose();
      await second.dispose();
    }
  });
});

function indexOfBytes(haystack: Uint8Array, needle: Uint8Array): number {
  outer: for (let i = 0; i <= haystack.length - needle.length; i += 1) {
    for (let j = 0; j < needle.length; j += 1)
      if (haystack[i + j] !== needle[j]) continue outer;
    return i;
  }
  return -1;
}
