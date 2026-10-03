# Third-party notices

The release artifact contains or depends on the following principal components. The generated SPDX SBOM in `artifacts/sbom.spdx.json` is the complete machine-readable inventory for the pinned lockfiles.

- [`@silurus/ooxml`](https://github.com/yukiyokotani/office-open-xml-viewer) — MIT; modern Office parsing/rendering (0.88.0 for DOCX, XLSX and PPTX). The package ships its own `THIRD_PARTY_NOTICES.md` for what the engine bundles; its optional region-map renderer carries a public-domain Natural Earth dataset and its optional math bundle MathJax, neither of which Zrimo imports.
- [`office_oxide`](https://github.com/yfedoseev/office_oxide) — MIT OR Apache-2.0; compound-file handling, Office IR/writer utilities and legacy XLS/PPT conversion.
- [`Fuse.js`](https://github.com/krisk/Fuse) — Apache-2.0; fuzzy matching behind the opt-in `search()` fallback.
- [`pdfjs-dist` / Mozilla PDF.js](https://github.com/mozilla/pdf.js) — Apache-2.0; browser PDF parsing,
  font/CMap handling, canvas rendering, and text extraction. The packaged
  standard-font, ICC, CMap, OpenJPEG, JBIG2, and QCMS assets retain the license
  files distributed with PDF.js under `dist/assets/pdfjs/`.
- [`@embedpdf/pdfium`](https://github.com/embedpdf/embed-pdf-viewer/tree/main/packages/pdfium) — MIT
  per the license file and `package.json` of the pinned 2.15.1 tarball (the
  upstream repository moved to Apache-2.0 on 2026-07-20); loaded only by the
  PDF edit worker. It bundles [PDFium](https://pdfium.googlesource.com/pdfium/)
  compiled to WebAssembly — BSD-3-Clause, see `LICENSE.pdfium` in that package —
  together with the third-party libraries PDFium builds in, such as FreeType
  (FreeType License), OpenJPEG (BSD-2-Clause), libpng (libpng License), zlib
  (Zlib License) and Anti-Grain Geometry 2.3, whose notices ship with the
  PDFium source tree.
- [`core-js`](https://github.com/zloirock/core-js) compatibility modules embedded in the PDF.js legacy browser build —
  MIT; polyfills required by the supported browser matrix, including the PDF
  worker realm.
- [`@napi-rs/canvas`](https://github.com/Brooooooklyn/canvas) — MIT; optional Node canvas backend pulled by PDF.js and excluded from Zrimo's browser bundle.
- [`image`](https://github.com/image-rs/image) and [`tiff`](https://github.com/image-rs/image-tiff) Rust crates — MIT OR Apache-2.0 / MIT; multi-page TIFF decoding and PNG output.
- [`wasm-bindgen`](https://github.com/wasm-bindgen/wasm-bindgen) — MIT OR Apache-2.0; browser bindings for project-owned Rust/WASM modules.
- [`zip`](https://github.com/zip-rs/zip2) — MIT; bounded in-memory ZIP/OOXML package manipulation.
- [Serde](https://github.com/serde-rs/serde), [`serde_json`](https://github.com/serde-rs/json) and [`thiserror`](https://github.com/dtolnay/thiserror) — MIT OR Apache-2.0; serialization and structured Rust errors.
- [GenOffice](https://github.com/genspark-ai/genoffice) — Apache-2.0; the
  selection-to-object matching ladder of the PDF overlay primitives (rectangle
  overlap, then a single containing object, then a text match after NFKC
  folding) follows the technique of its `apps/pdf/src/main/text-edit.ts`, and
  the PPTX text editing of `packages/viewer/src/edit/pptx/` follows its
  published ideas of tracing edited runs to their source so formatting
  survives, rebuilding only the paragraphs an edit touches, keeping
  `a:bodyPr` and `a:lstStyle`, and dropping a stale autofit scale. The DOCX
  editing of `packages/viewer/src/edit/docx/` follows its published ideas of
  a flat block index over the original bytes with the raw paragraph and run
  properties kept, property merges that keep the original bytes where a
  value does not change and place new children in schema order, and
  section properties that an edited paragraph never duplicates. The AI
  tooling of `packages/viewer/src/edit/ai/` follows its published ideas of
  a numbered document skeleton the model reads first with full content
  pulled on demand within a character budget, writes made only through a
  small validated tool set with a dry run, index addressing guarded by an
  optimistic "document seen" check, a per-turn snapshot to roll back, and
  Word tracked changes authored by the AI as the review channel. The
  implementations in `packages/viewer/src/edit/pdf/selection.ts`,
  `packages/viewer/src/edit/pptx/`, `packages/viewer/src/edit/docx/` and
  `packages/viewer/src/edit/ai/` are web-doc's own; no GenOffice code is
  included.
- [Adobe Glyph List and Adobe Glyph List For New Fonts](https://github.com/adobe-type-tools/agl-aglfn) — BSD-3-Clause; the glyph names and code points in `packages/viewer/src/edit/pdf/engine/glyph-names.ts` (AGLFN 1.7 and nine AGL 2.0 entries), which give an embedded CFF font a Unicode cmap. The ISOAdobe glyph names in `packages/viewer/src/edit/pdf/engine/cff.ts` are the standard strings of the CFF specification (Adobe Technical Note 5176). Their license:

  > Copyright 2002-2019 Adobe (http://www.adobe.com/).

  > Redistribution and use in source and binary forms, with or without modification, are permitted provided that the following conditions are met:

  > Redistributions of source code must retain the above copyright notice, this list of conditions and the following disclaimer.

  > Redistributions in binary form must reproduce the above copyright notice, this list of conditions and the following disclaimer in the documentation and/or other materials provided with the distribution.

  > Neither the name of Adobe nor the names of its contributors may be used to endorse or promote products derived from this software without specific prior written permission.

  > THIS SOFTWARE IS PROVIDED BY THE COPYRIGHT HOLDERS AND CONTRIBUTORS "AS IS" AND ANY EXPRESS OR IMPLIED WARRANTIES, INCLUDING, BUT NOT LIMITED TO, THE IMPLIED WARRANTIES OF MERCHANTABILITY AND FITNESS FOR A PARTICULAR PURPOSE ARE DISCLAIMED. IN NO EVENT SHALL THE COPYRIGHT HOLDER OR CONTRIBUTORS BE LIABLE FOR ANY DIRECT, INDIRECT, INCIDENTAL, SPECIAL, EXEMPLARY, OR CONSEQUENTIAL DAMAGES (INCLUDING, BUT NOT LIMITED TO, PROCUREMENT OF SUBSTITUTE GOODS OR SERVICES; LOSS OF USE, DATA, OR PROFITS; OR BUSINESS INTERRUPTION) HOWEVER CAUSED AND ON ANY THEORY OF LIABILITY, WHETHER IN CONTRACT, STRICT LIABILITY, OR TORT (INCLUDING NEGLIGENCE OR OTHERWISE) ARISING IN ANY WAY OUT OF THE USE OF THIS SOFTWARE, EVEN IF ADVISED OF THE POSSIBILITY OF SUCH DAMAGE.

- [Noto Sans](https://github.com/notofonts/noto-fonts) and [Noto Sans CJK](https://github.com/notofonts/noto-cjk) subset fonts — SIL Open Font License 1.1. The font manifest, complete OFL text, pinned source commits and SHA-256 hashes are included in `dist/fonts/`.
  No Microsoft proprietary font or copyleft runtime component is bundled. Transitive notices and license expressions are verified by `npm run licenses`.
