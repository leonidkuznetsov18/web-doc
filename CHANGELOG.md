# Changelog

## [0.6.2](https://github.com/leonidkuznetsov18/web-doc/compare/v0.6.1...v0.6.2) (2026-09-22)

### Bug fixes

* **viewer:** preserve pptx chart colors and axis intervals ([d7f4ca2](https://github.com/leonidkuznetsov18/web-doc/commit/d7f4ca29289836be0fffb73bd90b8eaad2a91760))

## [0.6.1](https://github.com/leonidkuznetsov18/web-doc/compare/v0.6.0...v0.6.1) (2026-09-15)

### Bug fixes

* **viewer:** bound fuzzy citation highlights to one passage ([4e3e3d0](https://github.com/leonidkuznetsov18/web-doc/commit/4e3e3d0bd50ef7da1f630754133236f4f2c2d54b))

## [0.6.0](https://github.com/leonidkuznetsov18/web-doc/compare/v0.5.0...v0.6.0) (2026-09-12)

### Features

* **viewer:** match fuzzy searches in a worker over a document index ([2870dac](https://github.com/leonidkuznetsov18/web-doc/commit/2870dacf6f1c32d1d3fabe39f2ba64f06b1c4988))

## [0.5.0](https://github.com/leonidkuznetsov18/web-doc/compare/v0.4.0...v0.5.0) (2026-09-12)

### Features

* **viewer:** reveal the active match and bound the fuzzy scan ([230a3ed](https://github.com/leonidkuznetsov18/web-doc/commit/230a3eda3779957a4ce6275b5467e8a489ca4d09))

### Bug fixes

* **viewer:** fit oversized inline docx pictures to the page ([b0dcc87](https://github.com/leonidkuznetsov18/web-doc/commit/b0dcc8714c6d68ecade8f507413a32b0e2f7eb94))

## [0.4.0](https://github.com/leonidkuznetsov18/web-doc/compare/v0.3.0...v0.4.0) (2026-09-12)

### Features

* **viewer:** fuzzy search with an approximate page hint ([b4ebb38](https://github.com/leonidkuznetsov18/web-doc/commit/b4ebb38b573e4f6e7163422dcec01ea78ce0a891))

### Bug fixes

* **viewer:** expose natural page sizes for images and svg ([5713392](https://github.com/leonidkuznetsov18/web-doc/commit/5713392ae8e303d2ad4dfe6c28de26a8de9c2254))

## [0.3.0](https://github.com/leonidkuznetsov18/web-doc/compare/v0.2.0...v0.3.0) (2026-09-08)

### Features

* **release:** rename the package to web-doc ([7667d9e](https://github.com/leonidkuznetsov18/web-doc/commit/7667d9e44e292a64cfe3b1e8afbc6ad66e513acd))

## [0.2.0](https://github.com/leonidkuznetsov18/zrimo/compare/v0.1.2...v0.2.0) (2026-09-08)

### Features

* **viewer:** scope search to an explicit page range ([257dfb4](https://github.com/leonidkuznetsov18/zrimo/commit/257dfb40a6754f2d1d38393a8e4ae1bad76d340c))

## 0.1.2 — 2026-08-28

- Fixed OOXML subtype detection for media-heavy ZIP packages by reading
  central-directory entry names instead of scanning compressed part contents.
- Hardened ZIP end-of-central-directory parsing against signatures embedded in
  ZIP comments, malformed central-directory bounds and unsupported ZIP64
  markers; prefix sniffing remains available as a best-effort fallback.
- Updated PDF.js and build-time tooling dependencies to address current npm
  vulnerability advisories.

## 0.1.1 — 2026-07-18

- Added visible search-result highlighting and active-match navigation for
  spreadsheets.
- Replaced the generic legacy XLS projection with bounded BIFF8 conversion that
  preserves source cell styles, fonts, fills, borders, alignment, row/column
  geometry, merged ranges, frozen panes and safe hyperlinks.
- Fixed initial spreadsheet positioning and made programmatic zoom retain the
  viewport's top-left logical point; pointer zoom remains cursor-anchored.
- Made the built-in and custom React integrations start at the same explicit
  100% zoom instead of deriving different fit-width scales from their layouts.
- Reworked the npm README around product value, quick setup and integration
  choices, with direct website, documentation and live-demo links.

## 0.1.0 — 2026-07-17

- Renamed the product to Zrimo and moved the npm package to `@zrimo/viewer`.
- Renamed the optional UI integration surface to `.zrimo-ui`, `--zrimo-*`, and `data-zrimo-*`; the bundled fallback font family is now `Zrimo Noto`.
- Renamed the shared Rust crate to `zrimo-core` and the fuzz workspace package to `zrimo-fuzz`.
- Fixed DOCX selection geometry, PDF font/runtime compatibility, legacy DOC structured conversion, spreadsheet virtualization and multi-cell clipboard behavior.
- Added spreadsheet column resizing plus Shift/Ctrl/Cmd range selection.
- Added built-in loading indicators and complete Vanilla/React integration examples.
- Qualified structured Word 97–2003 DOC in Chromium, Firefox and WebKit and added it to the visual regression lane.
- Added the GitHub Pages landing/documentation/demo bundle and manual deployment workflow.
- Added reproducible repository, package-content, consumer, SBOM, license, vulnerability and size checks.

## 0.1.0-alpha.0 — 2026-07-16

- Added the complete v1 browser format pipeline for modern/legacy Office, PDF, raster/TIFF, SVG and CSV/TSV.
- Added headless and container APIs, virtualized viewport, navigation, pan/zoom/fit, Unicode search, text/cell selection and optional localized UI.
- Added lazy multilingual Noto font packs, worker cancellation/timeouts, bounded render scheduling and allocation limits.
- Added Chromium/Firefox/WebKit compatibility scenarios, browser fallbacks, SSIM goldens, fuzz targets, performance/size reports and license/vulnerability checks.
