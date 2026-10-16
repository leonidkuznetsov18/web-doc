import assert from "node:assert/strict";
import { after, before, describe, it } from "node:test";
import type { TestContext } from "node:test";

import type {
  AdapterOpenContext,
  DocumentFormat,
  PdfBackend,
  ResourceLimits,
  TiffBackend,
} from "../src/index.js";
import {
  CsvDocumentAdapter,
  ImageDocumentAdapter,
  PdfDocumentAdapter,
  SvgDocumentAdapter,
  ViewerClient,
  ViewerError,
  defaultResourceLimits,
  parseDelimitedBytes,
  parseDelimitedText,
  sanitizeSvg,
} from "../src/index.js";
import { drawEncodedImage } from "../src/adapters/bitmap.js";

const originalCreateImageBitmap = globalThis.createImageBitmap;
let closedBitmaps = 0;

before(() => {
  Object.defineProperty(globalThis, "createImageBitmap", {
    configurable: true,
    value: async () => ({
      width: 2,
      height: 3,
      close: () => {
        closedBitmaps += 1;
      },
    }),
  });
});

after(() => {
  Object.defineProperty(globalThis, "createImageBitmap", {
    configurable: true,
    value: originalCreateImageBitmap,
  });
});

describe("CSV and TSV", () => {
  it("preserves quoted newlines, empty strings, and multilingual text", () => {
    const rows = parseDelimitedText(
      'name,note,empty\n"Иван","行1\n行2",\n"علي","नमस्ते",x',
      ",",
    );
    assert.deepEqual(rows, [
      ["name", "note", "empty"],
      ["Иван", "行1\n行2", ""],
      ["علي", "नमस्ते", "x"],
    ]);
  });

  it("detects delimiters and UTF-16 BOM without changing cell strings", () => {
    const text = "a;b\r\n1;2";
    const utf16 = new Uint8Array(2 + text.length * 2);
    utf16.set([0xff, 0xfe]);
    for (let index = 0; index < text.length; index += 1) {
      utf16[2 + index * 2] = text.charCodeAt(index);
      utf16[3 + index * 2] = text.charCodeAt(index) >> 8;
    }
    const parsed = parseDelimitedBytes(utf16, "csv");
    assert.equal(parsed.encoding, "utf-16le");
    assert.equal(parsed.delimiter, ";");
    assert.deepEqual(parsed.rows[1], ["1", "2"]);
  });

  it("enforces a bounded cell count", () => {
    assert.throws(
      () => parseDelimitedText("a,b\nc,d", ",", 3),
      isCode("resource-limit"),
    );
  });

  it("exposes one sheet and positioned cell text", async () => {
    const adapter = new CsvDocumentAdapter({
      parse: async (data, format) => parseDelimitedBytes(data, format),
    });
    const handle = await adapter.open(
      new TextEncoder().encode("ключ,значение\nمرحبا,नमस्ते"),
      context("csv", { fileName: "localization.csv" }),
    );
    const info = await adapter.getInfo(handle);
    assert.equal(info.unit, "sheet");
    assert.equal(info.sheetNames?.[0], "localization");
    assert.equal(info.sheets?.[0]?.maxRow, 2);
    const runs = await adapter.getTextMap(handle, 0);
    assert.equal(runs.length, 4);
    assert.equal(runs[2]?.direction, "rtl");
    assert.equal(runs[3]?.text, "नमस्ते");
  });
});

describe("SVG security", () => {
  it("renders sanitized SVG with the native decoder and an opaque source", async (t) => {
    let decodedSource = "";
    installNativeImageDecoder(t, async (source) => {
      decodedSource = source;
    });
    t.mock.method(globalThis, "createImageBitmap", async () => {
      throw new Error("SVG must not use the bitmap decoder");
    });
    t.mock.method(URL, "createObjectURL", () => {
      throw new Error("SVG must not create a same-origin object URL");
    });
    const adapter = new SvgDocumentAdapter();
    const handle = await adapter.open(
      new TextEncoder().encode(
        '<svg xmlns="http://www.w3.org/2000/svg" width="2" height="3" onload="steal()"><title>Логотип #1</title><rect fill="#00ff00" width="2" height="3"/></svg>',
      ),
      context("svg"),
    );
    const canvas = fakeCanvas();
    await adapter.render(handle, canvas, {
      pageIndex: 0,
      zoom: 2,
      devicePixelRatio: 1,
    });
    assert.ok(decodedSource.startsWith("data:image/svg+xml,"));
    const decodedSvg = decodeURIComponent(decodedSource.split(",")[1]!);
    assert.match(decodedSvg, /Логотип #1/);
    assert.match(decodedSvg, /fill="#00ff00"/);
    assert.doesNotMatch(decodedSvg, /onload|steal/);
    assert.equal(canvas.width, 4);
    assert.equal(canvas.height, 6);
  });

  it("normalizes native SVG decode failures", async (t) => {
    installNativeImageDecoder(t, async () => {
      throw new Error("Unsupported image");
    });
    const revoked = t.mock.method(URL, "revokeObjectURL", () => {});
    await assert.rejects(
      drawEncodedImage(
        new TextEncoder().encode("<svg/>"),
        "image/svg+xml",
        fakeCanvas(),
        { dpr: 1 },
      ),
      isCode("invalid-file"),
    );
    assert.equal(revoked.mock.callCount(), 0);
  });

  it("removes scripts, event handlers, foreign content, and external references", () => {
    const sanitized = sanitizeSvg(`
      <svg xmlns="http://www.w3.org/2000/svg" onload="steal()">
        <script>alert(1)</script>
        <foreignObject><iframe src="https://bad.invalid"/></foreignObject>
        <image href="data:image/png;base64,AAAA" />
        <a href="javascript:alert(1)"><rect width="10" height="10" /></a>
        <use href="#safe" />
      </svg>`);
    assert.doesNotMatch(
      sanitized,
      /script|foreignObject|iframe|onload|data:image|javascript:/i,
    );
    assert.match(sanitized, /href="#safe"/);
  });

  it("returns a warning when sanitization changes the source", async () => {
    const adapter = new SvgDocumentAdapter();
    const handle = await adapter.open(
      new TextEncoder().encode(
        '<svg xmlns="http://www.w3.org/2000/svg"><script/></svg>',
      ),
      context("svg"),
    );
    assert.equal(
      (await adapter.getInfo(handle)).warnings?.[0]?.code,
      "external-resource-blocked",
    );
    assert.deepEqual(await adapter.getTextMap(), []);
  });
});

describe("PDF adapter", () => {
  it("renders directly to canvas and exposes PDF.js text geometry", async () => {
    let renders = 0;
    let closes = 0;
    const backend: PdfBackend = {
      pageCount: 1,
      pageSize: async () => ({ width: 640, height: 900 }),
      renderPage: async (target) => {
        renders += 1;
        target.width = 2;
        target.height = 3;
      },
      pageText: async () => [
        { text: "A", x: 1, y: 2, width: 3, height: 4, direction: "ltr" },
        { text: "ش", x: 4, y: 2, width: 3, height: 4, direction: "rtl" },
      ],
      close: () => {
        closes += 1;
      },
    };
    const adapter = new PdfDocumentAdapter({ open: async () => backend });
    const handle = await adapter.open(
      new TextEncoder().encode("%PDF-1.7"),
      context("pdf"),
    );
    assert.deepEqual((await adapter.getInfo(handle)).pageSizes, [
      { width: 640, height: 900 },
    ]);
    const canvas = fakeCanvas();
    await adapter.render(handle, canvas, {
      pageIndex: 0,
      zoom: 1,
      devicePixelRatio: 1,
    });
    await adapter.render(handle, canvas, {
      pageIndex: 0,
      zoom: 1,
      devicePixelRatio: 1,
    });
    assert.equal(renders, 2);
    const runs = await adapter.getTextMap(handle, 0);
    assert.deepEqual(runs[0], {
      text: "A",
      x: 1,
      y: 2,
      width: 3,
      height: 4,
      direction: "ltr",
    });
    assert.equal(runs[1]?.direction, "rtl");
    await adapter.close(handle);
    assert.equal(closes, 1);
  });

  it("normalizes PDF.js password failures", async () => {
    let opened = false;
    const adapter = new PdfDocumentAdapter({
      open: async () => {
        opened = true;
        const error = new Error("No password given");
        error.name = "PasswordException";
        throw error;
      },
    });
    await assert.rejects(
      adapter.open(
        new TextEncoder().encode("%PDF-1.7\n/Encrypt 1 0 R"),
        context("pdf"),
      ),
      isCode("encrypted-document"),
    );
    assert.equal(opened, true);
  });
});

describe("image adapter", () => {
  it("uses native browser decoding for raster images", async () => {
    const beforeClosed = closedBitmaps;
    const adapter = new ImageDocumentAdapter();
    const handle = await adapter.open(png(2, 3), context("png"));
    const info = await adapter.getInfo(handle);
    assert.equal(info.pageCount, 1);
    // The decoded size feeds the viewport layout and fit modes; without it
    // fit-to-width would scale the bitmap against a letter-sized fallback.
    assert.deepEqual(info.pageSizes, [{ width: 2, height: 3 }]);
    await adapter.render(handle, fakeCanvas(), {
      pageIndex: 0,
      zoom: 2,
      devicePixelRatio: 1,
    });
    assert.deepEqual(await adapter.getTextMap(), []);
    assert.equal(closedBitmaps - beforeClosed, 2);
  });

  for (const fails of [false, true]) {
    it(`releases raster fallback object URLs after ${fails ? "failure" : "painting"}`, async (t) => {
      installNativeImageDecoder(t, async (source) => {
        assert.equal(source, "blob:raster-fallback");
        if (fails) throw new Error("Unsupported image");
      });
      const decoder = globalThis.createImageBitmap;
      Object.defineProperty(globalThis, "createImageBitmap", {
        configurable: true,
        value: undefined,
      });
      t.after(() => {
        Object.defineProperty(globalThis, "createImageBitmap", {
          configurable: true,
          value: decoder,
        });
      });
      t.mock.method(URL, "createObjectURL", () => "blob:raster-fallback");
      const revoked = t.mock.method(URL, "revokeObjectURL", () => {});
      const rendered = drawEncodedImage(png(2, 3), "image/png", fakeCanvas(), {
        dpr: 1,
      });
      if (fails) await assert.rejects(rendered, isCode("invalid-file"));
      else assert.deepEqual(await rendered, { width: 2, height: 3 });
      assert.equal(revoked.mock.callCount(), 1);
      assert.deepEqual(revoked.mock.calls[0]?.arguments, [
        "blob:raster-fallback",
      ]);
    });
  }

  it("exposes multi-page TIFF and releases its worker backend", async () => {
    let closed = 0;
    const backend: TiffBackend = {
      pages: [
        { width: 2, height: 3 },
        { width: 4, height: 5 },
      ],
      renderPagePng: async () => png(4, 5),
      close: () => {
        closed += 1;
      },
    };
    const adapter = new ImageDocumentAdapter({ openTiff: async () => backend });
    const handle = await adapter.open(
      Uint8Array.of(0x49, 0x49, 0x2a, 0),
      context("tiff"),
    );
    const info = await adapter.getInfo(handle);
    assert.equal(info.pageCount, 2);
    assert.deepEqual(info.pageSizes, [
      { width: 2, height: 3 },
      { width: 4, height: 5 },
    ]);
    await adapter.render(handle, fakeCanvas(), {
      pageIndex: 1,
      zoom: 1,
      devicePixelRatio: 1,
    });
    await adapter.close(handle);
    assert.equal(closed, 1);
  });
});

describe("image natural size", () => {
  it("opens without a page size when the browser cannot decode the image", async () => {
    const decoder = globalThis.createImageBitmap;
    Object.defineProperty(globalThis, "createImageBitmap", {
      configurable: true,
      value: async () => {
        throw new Error("unsupported");
      },
    });
    try {
      const adapter = new ImageDocumentAdapter();
      const handle = await adapter.open(png(2, 3), context("png"));
      assert.equal((await adapter.getInfo(handle)).pageSizes, undefined);
    } finally {
      Object.defineProperty(globalThis, "createImageBitmap", {
        configurable: true,
        value: decoder,
      });
    }
  });

  it("reads the SVG intrinsic size from width/height, units, or viewBox", async () => {
    const sizeOf = async (root: string) => {
      const adapter = new SvgDocumentAdapter();
      const handle = await adapter.open(
        new TextEncoder().encode(
          `<svg xmlns="http://www.w3.org/2000/svg" ${root}><rect/></svg>`,
        ),
        context("svg"),
      );
      return (await adapter.getInfo(handle)).pageSizes;
    };
    assert.deepEqual(await sizeOf('width="300" height="150px"'), [
      { width: 300, height: 150 },
    ]);
    assert.deepEqual(await sizeOf('width="72pt" height="1in"'), [
      { width: 96, height: 96 },
    ]);
    assert.deepEqual(await sizeOf('viewBox="0 0 800 600"'), [
      { width: 800, height: 600 },
    ]);
    assert.deepEqual(await sizeOf('width="400" viewBox="0,0,800,600"'), [
      { width: 400, height: 300 },
    ]);
    // Relative units cannot become CSS pixels; no size beats a wrong one.
    assert.equal(await sizeOf('width="100%" height="100%"'), undefined);
    assert.equal(await sizeOf(""), undefined);
  });
});

describe("default non-Office registration", () => {
  it("routes the complete agreed format matrix", () => {
    const registry = ViewerClient.create().registry;
    const expected: Readonly<Record<DocumentFormat, string>> = {
      docx: "office",
      docm: "office",
      xlsx: "office",
      xlsm: "office",
      pptx: "office",
      pptm: "office",
      ppsx: "office",
      doc: "office",
      xls: "office",
      ppt: "office",
      pdf: "pdf",
      csv: "delimited-data",
      tsv: "delimited-data",
      png: "image",
      jpeg: "image",
      gif: "image",
      webp: "image",
      svg: "svg",
      bmp: "image",
      tiff: "image",
    };
    for (const [format, adapterId] of Object.entries(expected))
      assert.equal(registry.resolve(format as DocumentFormat).id, adapterId);
  });
});

function context(
  format: DocumentFormat,
  extra: { fileName?: string; limits?: ResourceLimits } = {},
): AdapterOpenContext {
  return {
    format,
    signal: new AbortController().signal,
    limits: extra.limits ?? defaultResourceLimits,
    reportProgress: () => {},
    reportWarning: () => {},
    ...(extra.fileName ? { fileName: extra.fileName } : {}),
  };
}

function png(width: number, height: number): Uint8Array {
  const data = new Uint8Array(32);
  data.set([0x89, 0x50, 0x4e, 0x47, 0x0d, 0x0a, 0x1a, 0x0a]);
  const view = new DataView(data.buffer);
  view.setUint32(16, width);
  view.setUint32(20, height);
  return data;
}

function fakeCanvas(): HTMLCanvasElement | OffscreenCanvas {
  const context = {
    setTransform: () => {},
    clearRect: () => {},
    drawImage: () => {},
  };
  return {
    width: 1,
    height: 1,
    getContext: () => context,
  } as unknown as OffscreenCanvas;
}

function isCode(code: string): (error: unknown) => boolean {
  return (error) => error instanceof ViewerError && error.code === code;
}

function installNativeImageDecoder(
  t: TestContext,
  decode: (source: string) => Promise<void>,
): void {
  for (const [name, value] of Object.entries({
    document: {},
    Image: class {
      src = "";
      decoding = "auto";
      naturalWidth = 2;
      naturalHeight = 3;
      decode(): Promise<void> {
        return decode(this.src);
      }
    },
  })) {
    const previous = Object.getOwnPropertyDescriptor(globalThis, name);
    Object.defineProperty(globalThis, name, { configurable: true, value });
    t.after(() => {
      if (previous) Object.defineProperty(globalThis, name, previous);
      else Reflect.deleteProperty(globalThis, name);
    });
  }
}
