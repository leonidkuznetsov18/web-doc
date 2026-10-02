# Third-party notices

The release artifact contains or depends on the following principal components. The generated SPDX SBOM in `artifacts/sbom.spdx.json` is the complete machine-readable inventory for the pinned lockfiles.

- [`@silurus/ooxml`](https://github.com/yukiyokotani/office-open-xml-viewer) — MIT; modern Office parsing/rendering.
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
  `a:bodyPr` and `a:lstStyle`, and dropping a stale autofit scale. The
  implementations in `packages/viewer/src/edit/pdf/selection.ts` and
  `packages/viewer/src/edit/pptx/` are web-doc's own; no GenOffice code is
  included.
- [Noto Sans](https://github.com/notofonts/noto-fonts) and [Noto Sans CJK](https://github.com/notofonts/noto-cjk) subset fonts — SIL Open Font License 1.1. The font manifest, complete OFL text, pinned source commits and SHA-256 hashes are included in `dist/fonts/`.
  No Microsoft proprietary font or copyleft runtime component is bundled. Transitive notices and license expressions are verified by `npm run licenses`.
