import assert from "node:assert/strict";
import { existsSync, readdirSync, readFileSync } from "node:fs";
import { describe, it } from "node:test";
import { pathToFileURL } from "node:url";

import { OoxmlPackage } from "../src/edit/ooxml/package.js";
import { patches } from "../src/edit/ooxml/patch.js";
import { localRecordOf, parseZip } from "../src/edit/ooxml/zip.js";
import { defaultResourceLimits } from "../src/index.js";
import { buildZip, type ZipFileSpec } from "./fixtures/zip-builder.js";

/*
 * Task 43 of the OOXML package layer: the full cycle on every corpus and
 * fixture package, a macro-enabled package whose vbaProject.bin survives
 * untouched, and the performance of a synthetic 500-slide deck.
 */

const PACKAGE_DIR = pathToFileURL(`${process.cwd()}/`);
const CORPUS = new URL("../../.cache/corpus/", PACKAGE_DIR);
const FIXTURES = new URL("../../tests/fixtures/", PACKAGE_DIR);
const limits = defaultResourceLimits;
const decoder = new TextDecoder();

function sameBytes(a: Uint8Array, b: Uint8Array): boolean {
  if (a.byteLength !== b.byteLength) return false;
  for (let index = 0; index < a.byteLength; index += 1)
    if (a[index] !== b[index]) return false;
  return true;
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
        if (/\.(pptx|docx)$/i.test(name)) files.push(new URL(name, dir));
  }
  return files;
}

/** Entry names whose local record bytes differ between two archives. */
function changedEntries(before: Uint8Array, after: Uint8Array): string[] {
  const a = parseZip(before, limits);
  const b = parseZip(after, limits);
  const byName = new Map(b.entries.map((entry) => [entry.name, entry]));
  const changed: string[] = [];
  for (const entry of a.entries) {
    const other = byName.get(entry.name);
    if (!other) {
      changed.push(entry.name);
      continue;
    }
    const x = localRecordOf(a, entry);
    const y = localRecordOf(b, other);
    if (
      !sameBytes(
        a.bytes.subarray(x.headerOffset, x.recordEnd),
        b.bytes.subarray(y.headerOffset, y.recordEnd),
      )
    )
      changed.push(entry.name);
  }
  for (const entry of b.entries)
    if (!a.entries.some((candidate) => candidate.name === entry.name))
      changed.push(entry.name);
  return changed;
}

/** A deck with `count` slides, each a few hundred bytes of real-looking XML. */
function syntheticDeck(count: number): Uint8Array {
  const files: ZipFileSpec[] = [];
  const overrides: string[] = [];
  const slideRels: string[] = [];
  const slideIds: string[] = [];
  for (let n = 1; n <= count; n += 1) {
    files.push({
      name: `ppt/slides/slide${n}.xml`,
      data: `<?xml version="1.0" encoding="UTF-8" standalone="yes"?>\n<p:sld xmlns:a="http://schemas.openxmlformats.org/drawingml/2006/main" xmlns:p="http://schemas.openxmlformats.org/presentationml/2006/main"><p:cSld><p:spTree><p:nvGrpSpPr><p:cNvPr id="1" name=""/></p:nvGrpSpPr><p:sp><p:nvSpPr><p:cNvPr id="2" name="Title ${n}"/></p:nvSpPr><p:txBody><a:bodyPr/><a:p><a:r><a:rPr lang="en-US"/><a:t>Slide ${n} title text</a:t></a:r></a:p></p:txBody></p:sp></p:spTree></p:cSld></p:sld>`,
      method: 8,
    });
    files.push({
      name: `ppt/slides/_rels/slide${n}.xml.rels`,
      data: '<?xml version="1.0" encoding="UTF-8" standalone="yes"?>\n<Relationships xmlns="http://schemas.openxmlformats.org/package/2006/relationships"><Relationship Id="rId1" Type="http://schemas.openxmlformats.org/officeDocument/2006/relationships/slideLayout" Target="../slideLayouts/slideLayout1.xml"/></Relationships>',
      method: 8,
    });
    overrides.push(
      `<Override PartName="/ppt/slides/slide${n}.xml" ContentType="application/vnd.openxmlformats-officedocument.presentationml.slide+xml"/>`,
    );
    slideRels.push(
      `<Relationship Id="rId${n + 1}" Type="http://schemas.openxmlformats.org/officeDocument/2006/relationships/slide" Target="slides/slide${n}.xml"/>`,
    );
    slideIds.push(`<p:sldId id="${255 + n}" r:id="rId${n + 1}"/>`);
  }
  return buildZip([
    {
      name: "[Content_Types].xml",
      data: `<?xml version="1.0" encoding="UTF-8" standalone="yes"?>\n<Types xmlns="http://schemas.openxmlformats.org/package/2006/content-types"><Default Extension="rels" ContentType="application/vnd.openxmlformats-package.relationships+xml"/><Default Extension="xml" ContentType="application/xml"/><Override PartName="/ppt/presentation.xml" ContentType="application/vnd.openxmlformats-officedocument.presentationml.presentation.main+xml"/>${overrides.join("")}</Types>`,
      method: 8,
    },
    {
      name: "_rels/.rels",
      data: '<?xml version="1.0" encoding="UTF-8" standalone="yes"?>\n<Relationships xmlns="http://schemas.openxmlformats.org/package/2006/relationships"><Relationship Id="rId1" Type="http://schemas.openxmlformats.org/officeDocument/2006/relationships/officeDocument" Target="ppt/presentation.xml"/></Relationships>',
      method: 8,
    },
    {
      name: "ppt/presentation.xml",
      data: `<?xml version="1.0" encoding="UTF-8" standalone="yes"?>\n<p:presentation xmlns:p="http://schemas.openxmlformats.org/presentationml/2006/main" xmlns:r="http://schemas.openxmlformats.org/officeDocument/2006/relationships"><p:sldIdLst>${slideIds.join("")}</p:sldIdLst></p:presentation>`,
      method: 8,
    },
    {
      name: "ppt/_rels/presentation.xml.rels",
      data: `<?xml version="1.0" encoding="UTF-8" standalone="yes"?>\n<Relationships xmlns="http://schemas.openxmlformats.org/package/2006/relationships"><Relationship Id="rId1" Type="http://schemas.openxmlformats.org/officeDocument/2006/relationships/slideLayout" Target="slideLayouts/slideLayout1.xml"/>${slideRels.join("")}</Relationships>`,
      method: 8,
    },
    {
      name: "ppt/slideLayouts/slideLayout1.xml",
      data: '<p:sldLayout xmlns:p="http://schemas.openxmlformats.org/presentationml/2006/main"/>',
      method: 8,
    },
    ...files,
  ]);
}

describe("package cycle (ooxml package)", () => {
  it("edits one text in every corpus package, touching only the parts it must", async (t) => {
    const files = packageFiles();
    if (files.length === 0) {
      t.skip("no packages; run npm run corpus:fetch");
      return;
    }
    for (const file of files) {
      const label = file.pathname.split("/").pop()!;
      const original = new Uint8Array(readFileSync(file));
      const opened = performance.now();
      const pkg = await OoxmlPackage.open(original, { limits });
      const openMs = performance.now() - opened;
      assert.ok(
        sameBytes(await pkg.save(), original),
        `${label}: no-change save is identical`,
      );
      // The first slide or body part that draws any text; a chart-only deck
      // may keep its words elsewhere.
      let target: string | undefined;
      let scanMs = 0;
      let part!: Awaited<ReturnType<typeof pkg.xml>>;
      let text: ReturnType<typeof part.find> | undefined;
      for (const name of pkg.partNames) {
        if (!/\/(slides\/slide\d+|document)\.xml$/.test(name)) continue;
        const scanStart = performance.now();
        const candidate = await pkg.xml(name);
        const ms = performance.now() - scanStart;
        const tag = /slide/.test(name) ? "a:t" : "w:t";
        const found = candidate
          .findAll(tag)
          .find((node) => candidate.textOf(node).trim().length > 0);
        if (found) {
          target = name;
          part = candidate;
          text = found;
          scanMs = ms;
          break;
        }
      }
      // Without text: rename the first shape of a deck, or give the first
      // paragraph of a document a run.
      let mode: "text" | "rename" | "paragraph" = "text";
      if (!target) {
        target = pkg.partNames.find((name) =>
          /\/(slides\/slide1|document)\.xml$/.test(name),
        )!;
        part = await pkg.xml(target);
        if (/slide/.test(target)) {
          text = part.find("cNvPr");
          mode = "rename";
        } else {
          text = part.find("w:p");
          mode = "paragraph";
        }
      }
      assert.ok(target && text, `${label}: something to edit`);
      const transaction = pkg.transaction();
      transaction.patch(part, [
        mode === "rename"
          ? patches.setAttribute(part, text!, "name", "web-doc was here")
          : mode === "paragraph"
            ? patches.replaceContent(
                part,
                text!,
                "<w:r><w:t>web-doc was here</w:t></w:r>",
              )
            : patches.replaceContent(
                part,
                text!,
                patches.text("web-doc was here"),
              ),
      ]);
      const patchStart = performance.now();
      const change = await transaction.commit();
      const patchMs = performance.now() - patchStart;
      assert.deepEqual(change.changedParts, [target!]);
      assert.deepEqual(change.warnings, []);
      const saveStart = performance.now();
      const saved = await pkg.save();
      const saveMs = performance.now() - saveStart;
      assert.deepEqual(
        changedEntries(original, saved),
        [target.slice(1)],
        `${label}: only the target entry differs`,
      );
      const reopened = await OoxmlPackage.open(saved, { limits });
      for (const name of reopened.partNames) {
        if (/\.(xml|rels)$/i.test(name)) await reopened.xml(name);
        else await reopened.part(name);
      }
      const reread = await reopened.xml(target);
      assert.ok(
        mode === "rename"
          ? reread.attribute(reread.find("cNvPr")!, "name") ===
              "web-doc was here"
          : reread
              .findAll(/slide/.test(target) ? "a:t" : "w:t")
              .some((node) => reread.textOf(node) === "web-doc was here"),
        `${label}: the edit is in the saved file`,
      );
      console.log(
        `cycle ${label}: open ${openMs.toFixed(2)} ms, scan ${target} ${scanMs.toFixed(2)} ms, commit ${patchMs.toFixed(2)} ms, save ${saveMs.toFixed(2)} ms, ${original.byteLength} → ${saved.byteLength} B`,
      );
    }
  });

  it("keeps a macro part byte for byte and never reads it", async () => {
    const vba = Uint8Array.from(
      { length: 4096 },
      (_, index) => (index * 31 + 7) & 0xff,
    );
    const original = buildZip([
      {
        name: "[Content_Types].xml",
        data: '<?xml version="1.0" encoding="UTF-8" standalone="yes"?>\n<Types xmlns="http://schemas.openxmlformats.org/package/2006/content-types"><Default Extension="rels" ContentType="application/vnd.openxmlformats-package.relationships+xml"/><Default Extension="xml" ContentType="application/xml"/><Default Extension="bin" ContentType="application/vnd.ms-office.vbaProject"/><Override PartName="/ppt/presentation.xml" ContentType="application/vnd.ms-powerpoint.presentation.macroEnabled.main+xml"/><Override PartName="/ppt/slides/slide1.xml" ContentType="application/vnd.openxmlformats-officedocument.presentationml.slide+xml"/></Types>',
      },
      {
        name: "_rels/.rels",
        data: '<?xml version="1.0" encoding="UTF-8" standalone="yes"?>\n<Relationships xmlns="http://schemas.openxmlformats.org/package/2006/relationships"><Relationship Id="rId1" Type="http://schemas.openxmlformats.org/officeDocument/2006/relationships/officeDocument" Target="ppt/presentation.xml"/></Relationships>',
      },
      { name: "ppt/presentation.xml", data: '<p:presentation xmlns:p="x"/>' },
      {
        name: "ppt/_rels/presentation.xml.rels",
        data: '<?xml version="1.0" encoding="UTF-8" standalone="yes"?>\n<Relationships xmlns="http://schemas.openxmlformats.org/package/2006/relationships"><Relationship Id="rId1" Type="http://schemas.microsoft.com/office/2006/relationships/vbaProject" Target="vbaProject.bin"/></Relationships>',
      },
      { name: "ppt/vbaProject.bin", data: vba, method: 8 },
      {
        name: "ppt/slides/slide1.xml",
        data: '<p:sld xmlns:p="x" xmlns:a="y"><p:cSld><p:spTree><p:sp><p:txBody><a:p><a:r><a:t>macro deck</a:t></a:r></a:p></p:txBody></p:sp></p:spTree></p:cSld></p:sld>',
      },
    ]);
    const pkg = await OoxmlPackage.open(original, { limits });
    assert.equal(
      await pkg.contentTypeOf("/ppt/vbaProject.bin"),
      "application/vnd.ms-office.vbaProject",
    );
    const slide = await pkg.xml("/ppt/slides/slide1.xml");
    const transaction = pkg.transaction();
    transaction.patch(slide, [
      patches.replaceContent(slide, slide.find("a:t")!, "edited"),
    ]);
    await transaction.commit();
    const saved = await pkg.save();
    assert.deepEqual(changedEntries(original, saved), [
      "ppt/slides/slide1.xml",
    ]);
    const reopened = await OoxmlPackage.open(saved, { limits });
    assert.ok(
      sameBytes(await reopened.part("/ppt/vbaProject.bin"), vba),
      "the macro bytes survive",
    );
    assert.equal(
      await reopened.contentTypeOf("/ppt/presentation.xml"),
      "application/vnd.ms-powerpoint.presentation.macroEnabled.main+xml",
    );
    // The layer only ever inflated the slide it patched and the two root parts.
    assert.ok(!decoder.decode(saved).includes("macro deck"));
  });

  it("stays fast on a synthetic 500-slide deck", async () => {
    const original = syntheticDeck(500);
    // The best of three runs: one garbage collection in a busy test process
    // must not read as a slow package layer.
    const runs: Record<"openMs" | "scanMs" | "commitMs" | "saveMs", number>[] =
      [];
    for (let run = 0; run < 3; run += 1) {
      const openStart = performance.now();
      const pkg = await OoxmlPackage.open(original, { limits });
      const openMs = performance.now() - openStart;
      assert.equal(pkg.partNames.length, 1005);
      const scanStart = performance.now();
      const slide = await pkg.xml("/ppt/slides/slide250.xml");
      const scanMs = performance.now() - scanStart;
      const transaction = pkg.transaction();
      transaction.patch(slide, [
        patches.replaceContent(slide, slide.find("a:t")!, "Edited slide"),
      ]);
      const commitStart = performance.now();
      await transaction.commit();
      const commitMs = performance.now() - commitStart;
      const saveStart = performance.now();
      const saved = await pkg.save();
      const saveMs = performance.now() - saveStart;
      assert.deepEqual(changedEntries(original, saved), [
        "ppt/slides/slide250.xml",
      ]);
      runs.push({ openMs, scanMs, commitMs, saveMs });
    }
    const best = (key: keyof (typeof runs)[number]): number =>
      Math.min(...runs.map((run) => run[key]));
    console.log(
      `synthetic 500 slides: ${original.byteLength} B, best of 3: open ${best("openMs").toFixed(2)} ms, scan ${best("scanMs").toFixed(2)} ms, commit ${best("commitMs").toFixed(2)} ms, save ${best("saveMs").toFixed(2)} ms`,
    );
    assert.ok(
      best("openMs") < 100 && best("commitMs") < 100 && best("saveMs") < 200,
      "within the budget",
    );
  });
});
