import assert from "node:assert/strict";
import { existsSync, readdirSync, readFileSync } from "node:fs";
import { describe, it } from "node:test";
import { pathToFileURL } from "node:url";

import {
  extensionOf,
  folderOf,
  partKey,
  partNameOf,
  relationshipsPartOf,
  resolveTarget,
} from "../src/edit/ooxml/names.js";
import { OoxmlPackage } from "../src/edit/ooxml/package.js";
import {
  crc32,
  inflateEntry,
  localRecordOf,
  parseZip,
} from "../src/edit/ooxml/zip.js";
import { defaultResourceLimits, ViewerError } from "../src/index.js";
import { buildZip, minimalPackage } from "./fixtures/zip-builder.js";

/*
 * Task 38 of the OOXML package layer: the ZIP reader, OPC names and the
 * package's reading half, on hand-built archives and on every corpus and
 * fixture package.
 */

const PACKAGE_DIR = pathToFileURL(`${process.cwd()}/`);
const CORPUS = new URL("../../.cache/corpus/", PACKAGE_DIR);
const FIXTURES = new URL("../../tests/fixtures/", PACKAGE_DIR);
const limits = defaultResourceLimits;
const decoder = new TextDecoder();

function code(error: unknown): string {
  return error instanceof ViewerError ? error.code : String(error);
}

async function rejectsWith(
  run: () => Promise<unknown> | unknown,
  expected: string,
  label: string,
): Promise<void> {
  try {
    await run();
  } catch (error) {
    assert.equal(code(error), expected, `${label}: ${String(error)}`);
    return;
  }
  assert.fail(`${label}: no error`);
}

describe("zip reader (ooxml package)", () => {
  it("reads stored and deflated entries, descriptors, extra fields and a comment", async () => {
    const extra = Uint8Array.of(0x55, 0x54, 0x05, 0x00, 1, 2, 3, 4, 5);
    const bytes = buildZip(
      [
        { name: "a.txt", data: "stored text", method: 0 },
        { name: "b.xml", data: "<b>deflated</b>".repeat(50), method: 8 },
        { name: "c.xml", data: "<c/>", method: 8, descriptor: true },
        { name: "d.xml", data: "<d/>", method: 8, descriptor: "unsigned" },
        { name: "e.bin", data: Uint8Array.of(1, 2, 3), extra, comment: "note" },
        { name: "dir/", data: new Uint8Array(0), method: 0 },
      ],
      { comment: "archive comment" },
    );
    const archive = parseZip(bytes, limits);
    assert.deepEqual(
      archive.entries.map((entry) => entry.name),
      ["a.txt", "b.xml", "c.xml", "d.xml", "e.bin", "dir/"],
    );
    assert.equal(decoder.decode(archive.comment), "archive comment");
    assert.deepEqual([...archive.entries[4]!.extra], [...extra]);
    assert.equal(decoder.decode(archive.entries[4]!.comment), "note");
    for (const [index, expected] of [
      "stored text",
      "<b>deflated</b>".repeat(50),
      "<c/>",
      "<d/>",
    ].entries()) {
      const entry = archive.entries[index]!;
      assert.equal(
        decoder.decode(await inflateEntry(archive, entry, limits)),
        expected,
      );
    }
    // Local records cover header, data and the descriptor when present.
    const [, deflated, signed, unsigned] = archive.entries;
    assert.equal(
      localRecordOf(archive, deflated!).recordEnd,
      localRecordOf(archive, deflated!).dataEnd,
    );
    assert.equal(
      localRecordOf(archive, signed!).recordEnd -
        localRecordOf(archive, signed!).dataEnd,
      16,
    );
    assert.equal(
      localRecordOf(archive, unsigned!).recordEnd -
        localRecordOf(archive, unsigned!).dataEnd,
      12,
    );
    // Local records tile the file up to the central directory.
    let cursor = 0;
    for (const entry of archive.entries) {
      const local = localRecordOf(archive, entry);
      assert.equal(local.headerOffset, cursor);
      cursor = local.recordEnd;
    }
    assert.equal(cursor, archive.centralDirectoryOffset);
  });

  it("refuses what it does not support and reports what is broken", async () => {
    const good = [{ name: "a.xml", data: "<a/>", method: 8 }];
    await rejectsWith(
      () => parseZip(buildZip(good, { zip64Locator: true }), limits),
      "unsupported-package",
      "zip64 locator",
    );
    await rejectsWith(
      () => parseZip(buildZip(good, { zip64Count: true }), limits),
      "unsupported-package",
      "zip64 count",
    );
    await rejectsWith(
      () =>
        parseZip(
          buildZip([
            { ...good[0]!, override: { centralUncompressedSize: 0xffffffff } },
          ]),
          limits,
        ),
      "unsupported-package",
      "zip64 size marker",
    );
    await rejectsWith(
      () =>
        parseZip(
          buildZip([{ ...good[0]!, override: { flags: 0x0001 } }]),
          limits,
        ),
      "unsupported-package",
      "encrypted",
    );
    await rejectsWith(
      () =>
        parseZip(
          buildZip([
            { ...good[0]!, override: { centralMethod: 12, localMethod: 12 } },
          ]),
          limits,
        ),
      "unsupported-package",
      "bzip2 method",
    );
    await rejectsWith(
      () => parseZip(new TextEncoder().encode("not a zip at all"), limits),
      "invalid-file",
      "no end record",
    );
    await rejectsWith(
      () => parseZip(buildZip(good).subarray(0, 40), limits),
      "invalid-file",
      "truncated",
    );
    const lyingSize = parseZip(
      buildZip([
        {
          name: "a.xml",
          data: "<a/>".repeat(100),
          method: 8,
          override: { declaredUncompressedSize: 4 },
        },
      ]),
      limits,
    );
    await rejectsWith(
      () => inflateEntry(lyingSize, lyingSize.entries[0]!, limits),
      "resource-limit",
      "lying uncompressed size",
    );
    const badCrc = parseZip(
      buildZip([{ ...good[0]!, override: { crc: 0x12345678 } }]),
      limits,
    );
    await rejectsWith(
      () => inflateEntry(badCrc, badCrc.entries[0]!, limits),
      "invalid-file",
      "bad crc",
    );
    const badName = parseZip(
      buildZip([{ ...good[0]!, override: { localName: "b.xml" } }]),
      limits,
    );
    await rejectsWith(
      () => localRecordOf(badName, badName.entries[0]!),
      "invalid-file",
      "local name differs",
    );
    const badOffset = parseZip(
      buildZip([
        good[0]!,
        { name: "b.xml", data: "<b/>", override: { localHeaderOffset: 3 } },
      ]),
      limits,
    );
    await rejectsWith(
      () => localRecordOf(badOffset, badOffset.entries[1]!),
      "invalid-file",
      "wrong local offset",
    );
    // Varied text so the deflated data is long enough to damage safely.
    const corrupt = buildZip([
      {
        name: "a.xml",
        data: Array.from({ length: 400 }, (_, i) => `<a${i * 7919}/>`).join(""),
        method: 8,
      },
    ]);
    const local = localRecordOf(
      parseZip(corrupt, limits),
      parseZip(corrupt, limits).entries[0]!,
    );
    assert.ok(local.dataEnd - local.dataStart > 40, "enough deflated bytes");
    corrupt.fill(0xff, local.dataStart + 8, local.dataStart + 24);
    const corruptArchive = parseZip(corrupt, limits);
    await rejectsWith(
      () => inflateEntry(corruptArchive, corruptArchive.entries[0]!, limits),
      "invalid-file",
      "corrupt deflate data",
    );
    await rejectsWith(
      () =>
        parseZip(
          buildZip([
            {
              name: "big.bin",
              data: new Uint8Array(10),
              override: {
                centralUncompressedSize: limits.maxZipEntryBytes + 1,
              },
            },
          ]),
          limits,
        ),
      "resource-limit",
      "entry over the limit",
    );
  });

  it("computes CRC-32 as ZIP does", () => {
    assert.equal(crc32(new TextEncoder().encode("123456789")), 0xcbf43926);
    assert.equal(crc32(new Uint8Array(0)), 0);
  });

  it("normalizes and resolves OPC part names", () => {
    assert.equal(partNameOf("ppt/slides/slide1.xml"), "/ppt/slides/slide1.xml");
    assert.equal(
      partNameOf("/ppt/slides/slide1.xml"),
      "/ppt/slides/slide1.xml",
    );
    assert.equal(
      partNameOf("ppt\\slides\\slide1.xml"),
      "/ppt/slides/slide1.xml",
    );
    assert.equal(partKey("/PPT/Slides/Slide1.XML"), "/ppt/slides/slide1.xml");
    assert.equal(folderOf("/ppt/slides/slide1.xml"), "/ppt/slides/");
    assert.equal(extensionOf("/ppt/media/image1.PNG"), "png");
    assert.equal(extensionOf("/ppt/media/noext"), "");
    assert.equal(
      relationshipsPartOf("/ppt/slides/slide1.xml"),
      "/ppt/slides/_rels/slide1.xml.rels",
    );
    assert.equal(relationshipsPartOf("/"), "/_rels/.rels");
    assert.equal(
      resolveTarget("/ppt/slides/slide1.xml", "../media/image1.png"),
      "/ppt/media/image1.png",
    );
    assert.equal(
      resolveTarget("/ppt/slides/slide1.xml", "slide2.xml"),
      "/ppt/slides/slide2.xml",
    );
    assert.equal(
      resolveTarget("/ppt/slides/slide1.xml", "/ppt/theme/theme1.xml"),
      "/ppt/theme/theme1.xml",
    );
    assert.equal(
      resolveTarget("/", "ppt/presentation.xml"),
      "/ppt/presentation.xml",
    );
    assert.equal(
      resolveTarget(
        "/ppt/slides/slide1.xml",
        "./../slideLayouts/./slideLayout1.xml",
      ),
      "/ppt/slideLayouts/slideLayout1.xml",
    );
  });

  it("opens a package, lists parts, looks them up case-insensitively and caches them", async () => {
    const bytes = minimalPackage([
      { name: "ppt/slides/slide1.xml", data: "<p:sld/>", method: 8 },
      { name: "ppt/media/", data: new Uint8Array(0), method: 0 },
      {
        name: "ppt/media/image1.png",
        data: Uint8Array.of(0x89, 0x50, 0x4e, 0x47),
      },
    ]);
    const pkg = await OoxmlPackage.open(bytes, { limits });
    assert.deepEqual(pkg.partNames, [
      "/[Content_Types].xml",
      "/_rels/.rels",
      "/ppt/slides/slide1.xml",
      "/ppt/media/image1.png",
    ]);
    assert.ok(pkg.has("/PPT/SLIDES/SLIDE1.XML"));
    assert.ok(pkg.has("ppt/slides/slide1.xml"));
    assert.equal(pkg.entry("/ppt/slides/slide1.xml")?.method, 8);
    assert.equal(
      decoder.decode(await pkg.part("/ppt/slides/slide1.xml")),
      "<p:sld/>",
    );
    const first = await pkg.part("/ppt/media/image1.png");
    const second = await pkg.part("/ppt/media/image1.png");
    assert.deepEqual([...first], [0x89, 0x50, 0x4e, 0x47]);
    assert.notEqual(first.buffer, second.buffer, "callers get their own copy");
    await rejectsWith(
      () => pkg.part("/missing.xml"),
      "invalid-file",
      "missing part",
    );
    assert.ok(pkg.original === bytes, "the original is held, not copied");
    await rejectsWith(
      () =>
        OoxmlPackage.open(buildZip([{ name: "a.xml", data: "<a/>" }]), {
          limits,
        }),
      "invalid-file",
      "no content types",
    );
    await rejectsWith(
      () => OoxmlPackage.open(bytes, { limits, signal: AbortSignal.abort() }),
      "aborted",
      "aborted open",
    );
  });

  it("opens every corpus and fixture package and inflates every entry", async (t) => {
    const files: URL[] = [];
    if (existsSync(CORPUS))
      for (const name of readdirSync(CORPUS))
        if (/\.(pptx|docx)$/i.test(name)) files.push(new URL(name, CORPUS));
    for (const folder of ["pptx/", "docx/"]) {
      const dir = new URL(folder, FIXTURES);
      if (existsSync(dir))
        for (const name of readdirSync(dir))
          if (/\.(pptx|docx|pptm|docm)$/i.test(name))
            files.push(new URL(name, dir));
    }
    if (files.length === 0) {
      t.skip("no packages; run npm run corpus:fetch");
      return;
    }
    for (const file of files) {
      const bytes = new Uint8Array(readFileSync(file));
      const started = performance.now();
      const pkg = await OoxmlPackage.open(bytes, { limits });
      const opened = performance.now() - started;
      assert.ok(
        pkg.has("/[Content_Types].xml"),
        `${file.pathname}: content types`,
      );
      let inflated = 0;
      const inflateStart = performance.now();
      for (const name of pkg.partNames)
        inflated += (await pkg.part(name)).byteLength;
      const inflateMs = performance.now() - inflateStart;
      console.log(
        `package ${file.pathname.split("/").pop()}: ${bytes.byteLength} B, ${pkg.partNames.length} parts, open ${opened.toFixed(2)} ms, inflate all ${inflateMs.toFixed(1)} ms → ${inflated} B`,
      );
      assert.ok(inflated > bytes.byteLength / 4, "parts inflate to something");
    }
  });
});
