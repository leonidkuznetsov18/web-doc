import assert from "node:assert/strict";
import { existsSync, readdirSync, readFileSync } from "node:fs";
import { describe, it } from "node:test";
import { pathToFileURL } from "node:url";

import { OoxmlPackage } from "../src/edit/ooxml/package.js";
import { writeZip } from "../src/edit/ooxml/writer.js";
import {
  inflateEntry,
  localRecordOf,
  parseZip,
  type ZipArchive,
} from "../src/edit/ooxml/zip.js";
import { defaultResourceLimits, ViewerError } from "../src/index.js";
import { buildZip, minimalPackage } from "./fixtures/zip-builder.js";

/*
 * Task 39 of the OOXML package layer: the writer copies untouched entries
 * byte for byte, writes changed and new entries with fresh headers, drops
 * removed ones, and a package without changes comes back identical.
 */

const PACKAGE_DIR = pathToFileURL(`${process.cwd()}/`);
const CORPUS = new URL("../../.cache/corpus/", PACKAGE_DIR);
const FIXTURES = new URL("../../tests/fixtures/", PACKAGE_DIR);
const limits = defaultResourceLimits;
const encoder = new TextEncoder();
const decoder = new TextDecoder();

function sameBytes(a: Uint8Array, b: Uint8Array): boolean {
  if (a.byteLength !== b.byteLength) return false;
  for (let index = 0; index < a.byteLength; index += 1)
    if (a[index] !== b[index]) return false;
  return true;
}

/** Names of the entries whose local record or central record (offset aside) differs. */
function differingEntries(
  before: ZipArchive,
  after: ZipArchive,
): { readonly local: string[]; readonly central: string[] } {
  const local: string[] = [];
  const central: string[] = [];
  const afterByName = new Map(
    after.entries.map((entry) => [entry.name, entry]),
  );
  for (const entry of before.entries) {
    const other = afterByName.get(entry.name);
    if (!other) {
      local.push(entry.name);
      central.push(entry.name);
      continue;
    }
    const a = localRecordOf(before, entry);
    const b = localRecordOf(after, other);
    if (
      !sameBytes(
        before.bytes.subarray(a.headerOffset, a.recordEnd),
        after.bytes.subarray(b.headerOffset, b.recordEnd),
      )
    )
      local.push(entry.name);
    const centralA = before.bytes.slice(
      entry.centralOffset,
      entry.centralOffset + entry.centralLength,
    );
    const centralB = after.bytes.slice(
      other.centralOffset,
      other.centralOffset + other.centralLength,
    );
    new DataView(centralA.buffer).setUint32(42, 0, true);
    new DataView(centralB.buffer).setUint32(42, 0, true);
    if (!sameBytes(centralA, centralB)) central.push(entry.name);
  }
  return { local, central };
}

function packageFiles(): URL[] {
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
  return files;
}

describe("zip writer (ooxml package)", () => {
  it("rebuilds an archive with no changes byte for byte", async () => {
    const bytes = buildZip(
      [
        { name: "a.txt", data: "stored", method: 0 },
        { name: "b.xml", data: "<b/>".repeat(40), method: 8 },
        { name: "c.xml", data: "<c/>", method: 8, descriptor: true },
        { name: "d.xml", data: "<d/>", method: 8, descriptor: "unsigned" },
        {
          name: "e.bin",
          data: Uint8Array.of(1, 2, 3),
          extra: Uint8Array.of(0x55, 0x54, 1, 0, 9),
          comment: "x",
        },
        { name: "dir/", data: new Uint8Array(0), method: 0 },
      ],
      { comment: "kept" },
    );
    const archive = parseZip(bytes, limits);
    const rebuilt = await writeZip(archive, {
      changed: new Map(),
      added: new Map(),
      removed: new Set(),
    });
    assert.ok(sameBytes(rebuilt, bytes), "identical");
    for (const file of packageFiles()) {
      const original = new Uint8Array(readFileSync(file));
      const same = await writeZip(parseZip(original, limits), {
        changed: new Map(),
        added: new Map(),
        removed: new Set(),
      });
      assert.ok(
        sameBytes(same, original),
        `${file.pathname} rebuilt identically`,
      );
    }
  });

  it("changes one entry and leaves every other entry's bytes alone", async () => {
    const original = minimalPackage([
      { name: "ppt/slides/slide1.xml", data: "<p:sld>one</p:sld>", method: 8 },
      {
        name: "ppt/slides/slide2.xml",
        data: "<p:sld>two</p:sld>",
        method: 8,
        descriptor: true,
      },
      {
        name: "ppt/media/image1.png",
        data: Uint8Array.of(1, 2, 3, 4),
        method: 0,
      },
    ]);
    const pkg = await OoxmlPackage.open(original, { limits });
    assert.ok(sameBytes(await pkg.save(), original), "no changes, same bytes");
    pkg.applyOverlay({
      set: new Map([
        ["/ppt/slides/slide1.xml", encoder.encode("<p:sld>ONE</p:sld>")],
      ]),
      remove: new Set(),
    });
    const saved = await pkg.save();
    const before = parseZip(original, limits);
    const after = parseZip(saved, limits);
    assert.deepEqual(
      after.entries.map((entry) => entry.name),
      before.entries.map((entry) => entry.name),
      "same entries in the same order",
    );
    assert.deepEqual(differingEntries(before, after), {
      local: ["ppt/slides/slide1.xml"],
      central: ["ppt/slides/slide1.xml"],
    });
    const changed = after.entries.find(
      (entry) => entry.name === "ppt/slides/slide1.xml",
    )!;
    assert.equal(changed.method, 0, "stored by default");
    assert.equal(
      decoder.decode(await inflateEntry(after, changed, limits)),
      "<p:sld>ONE</p:sld>",
    );
    assert.deepEqual(pkg.changedParts, ["/ppt/slides/slide1.xml"]);
    // Deflated on request: smaller, and it reads back the same.
    const deflated = parseZip(
      await pkg.save({ compression: "deflate" }),
      limits,
    );
    const deflatedEntry = deflated.entries.find(
      (entry) => entry.name === "ppt/slides/slide1.xml",
    )!;
    assert.equal(deflatedEntry.method, 8);
    assert.equal(
      decoder.decode(await inflateEntry(deflated, deflatedEntry, limits)),
      "<p:sld>ONE</p:sld>",
    );
    assert.deepEqual(differingEntries(before, deflated).local, [
      "ppt/slides/slide1.xml",
    ]);
  });

  it("adds parts at the end with a fixed time, removes parts, and restores snapshots", async () => {
    const original = minimalPackage([
      { name: "ppt/slides/slide1.xml", data: "<p:sld/>", method: 8 },
      { name: "ppt/slides/slide2.xml", data: "<p:sld/>", method: 8 },
    ]);
    const pkg = await OoxmlPackage.open(original, { limits });
    const untouched = pkg.snapshot();
    pkg.applyOverlay({
      set: new Map([
        ["/ppt/slides/slide3.xml", encoder.encode("<p:sld>new</p:sld>")],
        ["/ppt/media/image1.png", Uint8Array.of(9, 9)],
      ]),
      remove: new Set(["/ppt/slides/slide2.xml"]),
    });
    assert.ok(pkg.has("/ppt/slides/slide3.xml"));
    assert.equal(pkg.has("/ppt/slides/slide2.xml"), false);
    assert.deepEqual(pkg.currentPartNames, [
      "/[Content_Types].xml",
      "/_rels/.rels",
      "/ppt/slides/slide1.xml",
      "/ppt/slides/slide3.xml",
      "/ppt/media/image1.png",
    ]);
    assert.equal(
      decoder.decode(await pkg.part("/ppt/slides/slide3.xml")),
      "<p:sld>new</p:sld>",
    );
    await assert.rejects(
      pkg.part("/ppt/slides/slide2.xml"),
      (error: unknown) =>
        error instanceof ViewerError && error.code === "invalid-file",
    );
    const saved = parseZip(await pkg.save(), limits);
    assert.deepEqual(
      saved.entries.map((entry) => entry.name),
      [
        "[Content_Types].xml",
        "_rels/.rels",
        "ppt/slides/slide1.xml",
        "ppt/slides/slide3.xml",
        "ppt/media/image1.png",
      ],
    );
    const added = saved.entries[3]!;
    assert.equal(added.modTime, 0);
    assert.equal(added.modDate, 0x21, "1980-01-01");
    assert.equal(added.extra.byteLength, 0);
    assert.equal(
      decoder.decode(await inflateEntry(saved, added, limits)),
      "<p:sld>new</p:sld>",
    );
    const reopened = await OoxmlPackage.open(saved.bytes, { limits });
    assert.ok(
      sameBytes(await reopened.save(), saved.bytes),
      "the saved package round-trips",
    );
    // Back to the untouched state in O(1).
    const edited = pkg.snapshot();
    pkg.restore(untouched);
    assert.ok(pkg.has("/ppt/slides/slide2.xml"));
    assert.equal(pkg.has("/ppt/slides/slide3.xml"), false);
    assert.ok(sameBytes(await pkg.save(), original));
    pkg.restore(edited);
    assert.ok(
      sameBytes(await pkg.save(), saved.bytes),
      "the same overlay saves the same bytes",
    );
    assert.throws(
      () => pkg.restore({ revision: 99 }),
      (error: unknown) =>
        error instanceof ViewerError && error.code === "lifecycle-error",
    );
  });

  it("refuses output that would need ZIP64 and aborts on request", async () => {
    const many = buildZip(
      Array.from({ length: 65_534 }, (_, index) => ({
        name: `p${index}`,
        data: new Uint8Array(0),
        method: 0,
      })),
    );
    const archive = parseZip(many, limits);
    await assert.rejects(
      writeZip(archive, {
        changed: new Map(),
        added: new Map([
          ["/one-more", { name: "/one-more", bytes: new Uint8Array(1) }],
          ["/two-more", { name: "/two-more", bytes: new Uint8Array(1) }],
        ]),
        removed: new Set(),
      }),
      (error: unknown) =>
        error instanceof ViewerError && error.code === "unsupported-package",
    );
    const small = parseZip(minimalPackage(), limits);
    await assert.rejects(
      writeZip(
        small,
        { changed: new Map(), added: new Map(), removed: new Set() },
        { signal: AbortSignal.abort() },
      ),
      (error: unknown) =>
        error instanceof ViewerError && error.code === "aborted",
    );
  });

  it("changes one part of every corpus package and keeps the rest verbatim", async (t) => {
    const files = packageFiles();
    if (files.length === 0) {
      t.skip("no packages; run npm run corpus:fetch");
      return;
    }
    for (const file of files) {
      const original = new Uint8Array(readFileSync(file));
      const pkg = await OoxmlPackage.open(original, { limits });
      const target = pkg.partNames.find((name) =>
        /\/(slides\/slide1|document)\.xml$/.test(name),
      )!;
      const text = decoder.decode(await pkg.part(target));
      pkg.applyOverlay({
        set: new Map([
          [target, encoder.encode(text.replace("<", "<!-- web-doc -->\n<"))],
        ]),
        remove: new Set(),
      });
      const started = performance.now();
      const saved = await pkg.save();
      const ms = performance.now() - started;
      const before = parseZip(original, limits);
      const after = parseZip(saved, limits);
      const diff = differingEntries(before, after);
      assert.deepEqual(
        diff.local,
        [target.slice(1)],
        `${file.pathname}: only the target's local record`,
      );
      assert.deepEqual(
        diff.central,
        [target.slice(1)],
        `${file.pathname}: only the target's central record`,
      );
      const reopened = await OoxmlPackage.open(saved, { limits });
      for (const name of reopened.partNames) await reopened.part(name);
      console.log(
        `writer ${file.pathname.split("/").pop()}: one part changed, save ${ms.toFixed(2)} ms, ${original.byteLength} → ${saved.byteLength} B`,
      );
    }
  });
});
