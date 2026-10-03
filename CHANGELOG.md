# Changelog

## [0.11.0](https://github.com/leonidkuznetsov18/web-doc/compare/v0.10.0...v0.11.0) (2026-10-03)

### Features

* **viewer:** read the text style a range shows ([76ca144](https://github.com/leonidkuznetsov18/web-doc/commit/76ca14435cdd2b42b6dddbac8b7f94f76027e9a4))

### Bug fixes

* **viewer:** finish docx layout before replacing edited pages ([38d8cdb](https://github.com/leonidkuznetsov18/web-doc/commit/38d8cdb8ca9caf12191131dcee1c79e0f6a2bf28))
* **viewer:** require an explicit docx default style ([176ac3d](https://github.com/leonidkuznetsov18/web-doc/commit/176ac3d8e41a63b9c52264d40e6fd54429cb25e0))

## [0.10.0](https://github.com/leonidkuznetsov18/web-doc/compare/v0.9.0...v0.10.0) (2026-10-03)

### Features

* **viewer:** give hosts the embedded font of pdf text ([607e3b8](https://github.com/leonidkuznetsov18/web-doc/commit/607e3b8511a8d2fdae7af7a98e8046d3da4353ce))

### Bug fixes

* **viewer:** hit pptx text painted past its shape ([003cd3f](https://github.com/leonidkuznetsov18/web-doc/commit/003cd3ff27acb7363e79a4622968536785938458))
* **viewer:** keep painted pages visible during edits ([298df20](https://github.com/leonidkuznetsov18/web-doc/commit/298df20ccd7440cf928f290c401dec8530a3e1f2))
* **viewer:** preserve queued paints across page revisions ([10d2900](https://github.com/leonidkuznetsov18/web-doc/commit/10d290015722b5efab9bc50c4f04b3d6555f1237))
* **viewer:** publish rasters when text layers fail ([d96b288](https://github.com/leonidkuznetsov18/web-doc/commit/d96b288be3f9006473adfb57cfb4c07dbd97d5ec))
* **viewer:** size pdf text by the size it shows ([6ad5542](https://github.com/leonidkuznetsov18/web-doc/commit/6ad55423e6d6751eb8c830e5f9d907b07d4292bd))

## [0.9.0](https://github.com/leonidkuznetsov18/web-doc/compare/v0.8.0...v0.9.0) (2026-10-02)

### Features

* **viewer:** name the cell that holds a docx table paragraph ([898e620](https://github.com/leonidkuznetsov18/web-doc/commit/898e620385d82b884654cfa74cb742e12fe40d90))
* **viewer:** turn a pdf page from the angle it has ([1d01bba](https://github.com/leonidkuznetsov18/web-doc/commit/1d01bba8132320423de21525fdc5676d33b62a9e))

### Bug fixes

* **viewer:** embed cyrillic text typed with dashes and currency signs ([88636ef](https://github.com/leonidkuznetsov18/web-doc/commit/88636efca00f894be22e3940cb911cc6c9aef8c4))
* **viewer:** give objects made after a reopen an unused id ([644cd0d](https://github.com/leonidkuznetsov18/web-doc/commit/644cd0d411dc33a747426537b0abb323e54057f3))
* **viewer:** keep ids that stale marks name out of new ids ([ccde3b5](https://github.com/leonidkuznetsov18/web-doc/commit/ccde3b56c04b43a6b7e677af04db40f64ffb0152))
* **viewer:** read a rotate field given as undefined as absent ([ee7314a](https://github.com/leonidkuznetsov18/web-doc/commit/ee7314a80627d3ed617986ace7ad0d3128a1a68c))

## [0.8.0](https://github.com/leonidkuznetsov18/web-doc/compare/v0.7.0...v0.8.0) (2026-10-02)

### Features

* **viewer:** name, restore and drop edit checkpoints ([89f091d](https://github.com/leonidkuznetsov18/web-doc/commit/89f091df59193710d7600dacbfe79ac3e7dce590))
* **viewer:** outline and describe edit sessions for prompts ([5993843](https://github.com/leonidkuznetsov18/web-doc/commit/59938434a9899050c76259d91bd608890940202a))
* **viewer:** resolve edit targets from quoted text, citations and kinds ([3724be2](https://github.com/leonidkuznetsov18/web-doc/commit/3724be2f506362dc67de801dc75b9e7ee2ac8d29))
* **viewer:** tool definitions and a tool executor for models ([43d6745](https://github.com/leonidkuznetsov18/web-doc/commit/43d6745bc92adb5e8599bf1ecdc46b3b2f37e632))
* **viewer:** write docx batches as tracked changes and list revisions ([d3ad054](https://github.com/leonidkuznetsov18/web-doc/commit/d3ad0541419f0e6a37226172b5a4e49a5d0d5e73))

### Bug fixes

* **viewer:** harden the tracked writer and the restore bookkeeping ([039c578](https://github.com/leonidkuznetsov18/web-doc/commit/039c5786d280c65f9fa055022c6d3281dc375c3a))
* **viewer:** keep classes defined in consumer bundles without splitting ([0c0f3af](https://github.com/leonidkuznetsov18/web-doc/commit/0c0f3af49c7a87fd1d5c67d4957a2e7a7902a5e6))
* **viewer:** keep classes defined in consumer bundles without splitting ([5532ddc](https://github.com/leonidkuznetsov18/web-doc/commit/5532ddc2d040fc4795161ee8675af5724d2116a7))

## [0.7.0](https://github.com/leonidkuznetsov18/web-doc/compare/v0.6.2...v0.7.0) (2026-10-02)

### Features

* **viewer:** add document editing contracts ([74363da](https://github.com/leonidkuznetsov18/web-doc/commit/74363da4de11c2eafcf652aeb2d957422eba04eb))
* **viewer:** add pdf page operations ([c128af4](https://github.com/leonidkuznetsov18/web-doc/commit/c128af4c1706c380475c33d6ae0085c20cc8563b))
* **viewer:** add pdf text layout, caret positions and range rectangles ([cce9df6](https://github.com/leonidkuznetsov18/web-doc/commit/cce9df612fee520622c71979d341d057d68bf29a))
* **viewer:** add pdfium wasm bridge for pdf editing ([9451bd0](https://github.com/leonidkuznetsov18/web-doc/commit/9451bd04bb8685b6c9d746ef5b2da0403bdff4e8))
* **viewer:** add the edit session core ([cbaf514](https://github.com/leonidkuznetsov18/web-doc/commit/cbaf514c2ba87bc1ebad9e5abdd348c1978a9cdc))
* **viewer:** add the pdf insert-text-box operation ([f537c9d](https://github.com/leonidkuznetsov18/web-doc/commit/f537c9db134e0c4d8acce4e81973742e184e4731))
* **viewer:** add the revision-2 editing contract types ([a875573](https://github.com/leonidkuznetsov18/web-doc/commit/a8755734de64bda3ad869afcb552b594abb07807))
* **viewer:** add, duplicate, delete and reorder pptx slides ([00fae25](https://github.com/leonidkuznetsov18/web-doc/commit/00fae25fd7e2fa8a04222306cfdafcf3bb5d2ae6))
* **viewer:** answer pdf hover from a main-thread geometry cache ([8785365](https://github.com/leonidkuznetsov18/web-doc/commit/87853652d7ec3c9b9b55fd9825d668e7576ac8ca))
* **viewer:** carry paragraph ids on docx text runs ([ae2dda5](https://github.com/leonidkuznetsov18/web-doc/commit/ae2dda5836726dc50fcc85eaed28487b5eb218a4))
* **viewer:** derive edit ids from state ids and report removed ids ([7820cf5](https://github.com/leonidkuznetsov18/web-doc/commit/7820cf566dc8eff2db6330192dd5e64f7d7e9bc0))
* **viewer:** draw and restyle pdf shapes ([f5cf9e6](https://github.com/leonidkuznetsov18/web-doc/commit/f5cf9e62bb9d1436a8318cadd790870aab77e906))
* **viewer:** edit docx paragraph text, run and paragraph styles ([6712548](https://github.com/leonidkuznetsov18/web-doc/commit/6712548846f0ea29876464cd827e28cea6313a0a))
* **viewer:** edit pdf text boxes in place ([f3ba732](https://github.com/leonidkuznetsov18/web-doc/commit/f3ba732d52166a7d0be098cffc02bfb9c829cf6e))
* **viewer:** edit text that already exists in a pdf ([eb8366a](https://github.com/leonidkuznetsov18/web-doc/commit/eb8366acd098bb4b3f53f8693ff2b34a2555f6d2))
* **viewer:** embed fonts for pdf text the standard fonts cannot draw ([736c480](https://github.com/leonidkuznetsov18/web-doc/commit/736c4807d80030f5b0d52966f9494f1f080d91b5))
* **viewer:** finish pptx editing with latency, fixtures and docs ([968981e](https://github.com/leonidkuznetsov18/web-doc/commit/968981e9a900db925069c59d72cc8be36e90d087))
* **viewer:** finish the pdf revision-2 contract ([bcf8b23](https://github.com/leonidkuznetsov18/web-doc/commit/bcf8b2390df4adcc129f34c42f8261dde4befbf2))
* **viewer:** fit docx pictures and id paragraphs in a pre-pass ([e793b1d](https://github.com/leonidkuznetsov18/web-doc/commit/e793b1d056adbfeeb1fa185d9645ad5e35639fed))
* **viewer:** harden pdf editing for signed files and json clients ([b5fdcfa](https://github.com/leonidkuznetsov18/web-doc/commit/b5fdcfa1ae29a6137c01b7e0499ba9d743eca978))
* **viewer:** insert docx tables and edit their cells ([de8f988](https://github.com/leonidkuznetsov18/web-doc/commit/de8f98846a32c3e1199e8f733da425e3ed5b8d7d))
* **viewer:** insert jpeg and png images into pdf pages ([28a3e7e](https://github.com/leonidkuznetsov18/web-doc/commit/28a3e7e59b0dc1c7656f3c5362116a2061122338))
* **viewer:** insert pictures and tables into pptx slides ([2354aa8](https://github.com/leonidkuznetsov18/web-doc/commit/2354aa89a227b9a28be4cf8d77e41be83d3bdc21))
* **viewer:** insert tables into pdf pages and edit their cells ([c3a0d79](https://github.com/leonidkuznetsov18/web-doc/commit/c3a0d7929f06442486958756537167969b7fbb8d))
* **viewer:** insert, move and delete docx paragraphs and pictures ([b9e75e7](https://github.com/leonidkuznetsov18/web-doc/commit/b9e75e730b039c8f0bf2176d9dae9ba030b73848))
* **viewer:** inspect docx bodies through the edit session ([2b9c2ea](https://github.com/leonidkuznetsov18/web-doc/commit/2b9c2ead4e16354d0691763ab7b41babf7252e33))
* **viewer:** inspect pptx slides through the ooxml edit worker ([b25378e](https://github.com/leonidkuznetsov18/web-doc/commit/b25378e56d377e0e6cf20dcebd47c118f81ce181))
* **viewer:** keep edit checkpoints and an asset store ([f3d36ed](https://github.com/leonidkuznetsov18/web-doc/commit/f3d36edd81601d9885ec936d01e619fdeed102c7))
* **viewer:** keep the view across edits and map page geometry ([499d858](https://github.com/leonidkuznetsov18/web-doc/commit/499d858fb59a08e693b262dabb83f81bc4553f4c))
* **viewer:** map pdf selections and ranges across later edits ([6600a50](https://github.com/leonidkuznetsov18/web-doc/commit/6600a5004476ff0109624bdfa313577df2b3cedc))
* **viewer:** model pdf pages and elements for editing ([d73e5fc](https://github.com/leonidkuznetsov18/web-doc/commit/d73e5fc05a5e4e0d95648a7418ba707633b525fc))
* **viewer:** move docx and xlsx to the 0.88 ooxml engine ([5471c83](https://github.com/leonidkuznetsov18/web-doc/commit/5471c8308a58166a3895d2d457341fd60eb44321))
* **viewer:** move, resize and delete pdf elements ([bb028f3](https://github.com/leonidkuznetsov18/web-doc/commit/bb028f3b2063e505e58191107df97944f900652e))
* **viewer:** patch ooxml parts in verified, atomic transactions ([525fc3a](https://github.com/leonidkuznetsov18/web-doc/commit/525fc3a238d2b15c5ca15eaf70fe78ed8146a218))
* **viewer:** read ooxml packages with a verbatim-copy zip reader ([2533306](https://github.com/leonidkuznetsov18/web-doc/commit/2533306887a7c9dd3415c7fcf365f0942bb5f49a))
* **viewer:** render a pdf page with chosen elements left out ([fd3c479](https://github.com/leonidkuznetsov18/web-doc/commit/fd3c4794c05263d3b83595edaf6fabfcb3a4ee85))
* **viewer:** replace a range of a pdf text element ([a3ca9da](https://github.com/leonidkuznetsov18/web-doc/commit/a3ca9dada777fe855197d1bf98afe1ccfcbe5c3e))
* **viewer:** replace and restyle pptx text with minimal run patches ([6ab0da2](https://github.com/leonidkuznetsov18/web-doc/commit/6ab0da2dffb664c557a37ae2739a41c8d044af9a))
* **viewer:** return read envelopes from edit sessions ([21fba38](https://github.com/leonidkuznetsov18/web-doc/commit/21fba38f7747e5ca5c70dd320875ce13275db6b5))
* **viewer:** save pdfs in full by default with a compaction pass ([4520d1a](https://github.com/leonidkuznetsov18/web-doc/commit/4520d1a0ce421f96c240d87d5b95295365a897cb))
* **viewer:** scan ooxml parts with exact ranges ([f7141ab](https://github.com/leonidkuznetsov18/web-doc/commit/f7141ab7dd91e5175ba61e85a8fcecb03a7edbe2))
* **viewer:** show compacted pdf saves, fall back when compaction fails ([df81265](https://github.com/leonidkuznetsov18/web-doc/commit/df81265b81137e14647ff08e9f3c158562a745b5))
* **viewer:** start pdf edit sessions on a pdfium worker ([6fc4647](https://github.com/leonidkuznetsov18/web-doc/commit/6fc46476cdcb9c33dbe2726b83208c50dc2efd20))
* **viewer:** style, move, resize, delete and insert pptx shapes ([d232135](https://github.com/leonidkuznetsov18/web-doc/commit/d23213570d8223aef1be8a2ffd2387f7683721d0))
* **viewer:** swap edited documents in two phases and report layout ([c976040](https://github.com/leonidkuznetsov18/web-doc/commit/c976040919575de28e30a15866d5cef3fb635890))
* **viewer:** validate edit operations against json schemas ([bae3420](https://github.com/leonidkuznetsov18/web-doc/commit/bae34201a492fe885bc3adeb940f864799b82005))
* **viewer:** wire edit sessions into the document viewer ([6b92518](https://github.com/leonidkuznetsov18/web-doc/commit/6b925187bdaf4933c736dcbd0bc8ee56f4f28f94))
* **viewer:** write ooxml packages copying untouched entries verbatim ([2449565](https://github.com/leonidkuznetsov18/web-doc/commit/24495651adf9c857c2b26ba16e67c39e8721e7df))

### Bug fixes

* **viewer:** harden docx editing after review ([24f48e9](https://github.com/leonidkuznetsov18/web-doc/commit/24f48e959b891819ef846b530e64a383968d7396))
* **viewer:** harden pptx editing and the docx pre-pass after review ([7248c78](https://github.com/leonidkuznetsov18/web-doc/commit/7248c7830f42f710fd7c8c05eef2b2a3d747678a))
* **viewer:** harden the pdf full-save compaction pass ([7586f66](https://github.com/leonidkuznetsov18/web-doc/commit/7586f66e317de56e4ae8bdd83dabdd40eb127fcc))
* **viewer:** make edit sessions safe around the host ([149ae66](https://github.com/leonidkuznetsov18/web-doc/commit/149ae66e01e836f5d8313b4615f9a8c91ae88eda))

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
