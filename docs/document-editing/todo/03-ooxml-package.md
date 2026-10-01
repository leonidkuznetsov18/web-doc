# Module 03. `ooxml-package` — the shared OOXML package layer

**Status:** ✅ Done 2026-10-02 (T38–T43, Linear ACTION-810). Approved by
Leonid on 2026-10-02 with `store` as the default compression and `warn` for
dangling relationship targets.

## Goal

Give `pptx-edit` (module 04) and `docx-edit` (module 06) one way to change an
OOXML file losslessly: open the package once, read any part on demand, patch
exact ranges of an XML part, add media and parts, keep relationships and
content types consistent, and write a package in which every untouched entry
is the original entry byte for byte. Nothing in this module knows what a
slide or a paragraph is.

## Requirement sources

- [`../00-roadmap.md`](../00-roadmap.md): fixed decisions 3 (the original
  file is the source of truth), 6 (PPTX and DOCX are edited XML-first; untouched
  parts are copied as they are) and 7 (macro parts are copied byte for byte
  and never executed); the boundaries (no new runtime dependency without
  approval; preserve content the editor does not understand).
- [`./01-edit-core.md`](./01-edit-core.md): the engine interface the format
  modules implement on top of this layer — `restore({ base, batches })`,
  `materialize(purpose)`, deterministic output, "no changes returns the
  original bytes".
- Linear ACTION-810: unzip once; copy untouched ZIP entries as they are (same
  compressed bytes, order and attributes); an offset-preserving XML scanner so
  patches replace exact ranges instead of re-serializing; patches verify
  themselves by re-parsing and reading the target back; helpers for `_rels`
  and `[Content_Types].xml`, media and new parts; `vbaProject.bin` copied and
  never executed.
- ECMA-376 Part 2 / ISO/IEC 29500-2 (Open Packaging Conventions): part names
  are absolute, case-insensitive, `/`-separated; `[Content_Types].xml` maps
  extensions (`Default`) and part names (`Override`) to content types;
  relationships of `/a/b.xml` live in `/a/_rels/b.xml.rels`; the package's own
  relationships in `/_rels/.rels`; targets are relative to the source part
  unless `TargetMode="External"`.
- Research of 2026-10-01 (patterns only, no code):
  - GenOffice (Apache-2.0) proves the approach at scale: a byte-exact scanner
    that locates element ranges in the original XML so a save splices new
    fragments and copies untouched elements as raw substrings; a flat model
    that is "an index over the original bytes"; ZIP entries that are "never
    inflated or re-deflated" when untouched; a metered-inflation zip-bomb
    gate shared by its DOCX and PPTX engines. Its open issues name the two
    traps this spec designs around: `#1579` "patch splice shifts on
    multi-byte UTF-8 edits" (string offsets applied to bytes) and `#1687` "a
    rels part with tens of thousands of relationships makes the save
    quadratic".
  - SuperDoc (`superdoc/docx-editor`): its authors moved from a
    ProseMirror-authoritative model to "an OOXML-backed document model"
    because "a DOCX is a package of related XML parts, relationships, and
    assets rather than one editor tree" — the same conclusion this layer
    rests on. Nothing else is taken: the V1 pipeline re-serializes XML
    through a JSON tree (not byte-preserving) and is AGPL-3.0, the V2 engine
    is proprietary and its licence forbids analysis, and both are DOCX-only.
    See [decision 6](#decisions).
  - `@silurus/ooxml` 0.72.2 / 0.88.0 are read-only renderers with no package
    or serialization API; editing cannot go through them, and they must
    reopen what this layer writes.

## Dependencies

- Module 01 `edit-core`: types and the engine contract only; this module has
  no session of its own.
- No new runtime dependency. Inflation uses the platform
  `DecompressionStream("deflate-raw")`, available in the browser baseline
  (Safari 16.4+, current Chrome, Edge and Firefox) and in Node 18+. Writing
  needs no compression at all by default (see
  [decision 2](#decisions)); `CompressionStream("deflate-raw")` is used only
  when a caller asks for deflated output and the platform has it.

## In scope

1. **ZIP container.** Read the central directory and local headers of a
   package held in memory; inflate an entry on demand; write a package that
   copies untouched entries verbatim and appends or replaces changed ones.
2. **OPC model.** Part names, `[Content_Types].xml` (`Default` and
   `Override`), relationship parts, target resolution, the package root.
3. **XML scanner.** A non-validating, offset-preserving tokenizer that turns
   an XML part into an element tree whose every node carries its exact range
   in the part's text, plus the opaque regions (declaration, comments,
   processing instructions, CDATA, DOCTYPE) it does not interpret.
4. **Patches.** Range replacements on one part, applied together or not at
   all, verified by re-scanning the part and reading the patched node back.
   Helpers for the common shapes: replace an element's inner XML, replace or
   remove an element, insert a fragment before or after one, set or remove an
   attribute.
5. **Transactions and snapshots.** A transaction collects part patches, new
   or replaced parts, removed parts, relationship and content-type changes
   across the package and commits them atomically; a snapshot captures the
   package's changed-part overlay so an engine can implement `restore`
   cheaply.
6. **Helpers.** Add a media part from bytes with a content type and a
   relationship from its source part (deduplicated by content hash); add a
   new XML part with its content type; allocate relationship ids and unique
   part names; list the parts a part references.
7. **Macro and binary parts.** Any part the layer does not patch is copied
   byte for byte; `vbaProject.bin`, OLE objects, fonts, media and signatures
   are never parsed or executed.
8. **Limits, errors and fuzzing.** The existing `ResourceLimits` apply; new
   error codes are typed; the ZIP reader and the XML scanner join
   `scripts/fuzz-js.mjs`.

## Out of scope

- Any format semantics: slides, placeholders, paragraphs, runs, styles,
  themes, numbering. Modules 04 and 06 own them and may add their own
  helpers on top of this layer.
- Re-serializing XML from a DOM or JSON tree; DOM-based editing of any kind.
  Every byte that is not inside a patched range is the original byte.
- Writing ZIP64 archives; opening encrypted or multi-disk archives (refused
  with `unsupported-package`). Packages over the existing limits.
- XLSX editing (the viewer stays read-only for spreadsheets; the layer may
  open an XLSX, nothing uses it).
- Digital signatures (`_xmlsignatures`): signature parts are copied as they
  are; a change invalidates the signature, and the format module reports
  that with the first-change warning, as the PDF module does.
- Streaming to disk, incremental saves into the original file, or any
  server-side processing.
- Validation against the OOXML schemas; the layer checks well-formedness
  only.

## API

Internal, under `packages/viewer/src/edit/ooxml/`, not exported from the
package root (like the engine interface). The plan may refine signatures;
the responsibilities are binding.

### Package

```ts
export class OoxmlPackage {
  /** Parses the central directory; parts are inflated on first use. */
  static open(
    bytes: Uint8Array,
    options: { readonly limits: ResourceLimits; readonly signal?: AbortSignal },
  ): Promise<OoxmlPackage>;

  /** The original bytes, never mutated. */
  readonly original: Uint8Array;
  /** Absolute OPC part names in archive order, e.g. "/ppt/slides/slide1.xml". */
  readonly partNames: readonly string[];
  /** The ZIP entry behind a part: method, sizes, CRC, flags, offsets. */
  entry(name: string): ZipEntryInfo | undefined;
  has(name: string): boolean;
  /** From [Content_Types].xml: an Override, else the Default for the extension. */
  contentTypeOf(name: string): string | undefined;

  /** Current bytes of a part: the overlay's when changed, else the original's. */
  part(name: string, signal?: AbortSignal): Promise<Uint8Array>;
  /** Decoded and scanned; cached until the part changes. */
  xml(name: string, signal?: AbortSignal): Promise<XmlPart>;
  /** The relationships of a part (or of the package for "/"). */
  relationships(name: string, signal?: AbortSignal): Promise<RelationshipSet>;
  /** Resolves a relationship target against its source part to an absolute name. */
  resolve(sourcePart: string, target: string): string;

  transaction(): PackageTransaction;
  /** The changed-part overlay, cheap to take and to restore. */
  snapshot(): PackageSnapshot;
  restore(snapshot: PackageSnapshot): void;
  /** Names of the parts that differ from the original. */
  readonly changedParts: readonly string[];

  /** The package bytes: the original when nothing changed. */
  save(options?: SaveOptions, signal?: AbortSignal): Promise<Uint8Array>;
}

export interface ZipEntryInfo {
  readonly name: string; // raw entry name as stored, e.g. "ppt/slides/slide1.xml"
  readonly method: 0 | 8; // stored or deflated; anything else is refused
  readonly compressedSize: number;
  readonly uncompressedSize: number;
  readonly crc32: number;
  readonly flags: number;
  readonly localHeaderOffset: number;
}

export interface SaveOptions {
  /** How changed and new entries are written; default "store". */
  readonly compression?: "store" | "deflate";
}

export interface RelationshipSet {
  readonly sourcePart: string; // "/" for the package
  readonly partName: string | undefined; // the .rels part, when it exists
  readonly items: readonly Relationship[];
  byId(id: string): Relationship | undefined;
  byType(type: string): readonly Relationship[];
}

export interface Relationship {
  readonly id: string;
  readonly type: string;
  readonly target: string; // as written
  readonly targetMode: "Internal" | "External";
  /** Absolute part name for internal targets. */
  readonly targetPart?: string;
}
```

### XML scanner

```ts
export interface XmlPart {
  readonly name: string;
  /** The decoded text the ranges index into (UTF-16 code units). */
  readonly text: string;
  /** "utf-8" today; a part in any other encoding is not patchable. */
  readonly encoding: "utf-8";
  readonly hasBom: boolean;
  readonly declaration?: XmlRange;
  readonly root: XmlElement;
  /** Comments, processing instructions, CDATA and DOCTYPE, in document order. */
  readonly opaque: readonly XmlRange[];

  /** Depth-first search by qualified or local name, optionally under a node. */
  find(name: string, under?: XmlElement): XmlElement | undefined;
  findAll(name: string, under?: XmlElement): readonly XmlElement[];
  /** A node by its child-index path from the root, e.g. [0, 3, 1]. */
  at(path: readonly number[]): XmlElement | undefined;
  /** Concatenated character data of a subtree, entities decoded. */
  textOf(node: XmlElement): string;
  attribute(node: XmlElement, name: string): string | undefined;
}

export interface XmlRange {
  readonly kind: "declaration" | "comment" | "pi" | "cdata" | "doctype";
  readonly start: number;
  readonly end: number;
}

export interface XmlElement {
  readonly name: string; // qualified, as written: "p:sp"
  readonly prefix: string; // "p"
  readonly local: string; // "sp"
  /** Namespace URI resolved through xmlns declarations in scope, when declared. */
  readonly namespace?: string;
  readonly attributes: readonly XmlAttribute[];
  /** Range of the whole element: "<" of the start tag to ">" of the end tag or "/>". */
  readonly start: number;
  readonly end: number;
  /** Range of the content between the tags; empty and equal for a self-closing element. */
  readonly contentStart: number;
  readonly contentEnd: number;
  readonly selfClosing: boolean;
  readonly children: readonly XmlElement[];
  readonly parent?: XmlElement;
  /** Index among the parent's element children. */
  readonly index: number;
  /** Child-index path from the root; stable within one scan. */
  readonly path: readonly number[];
}

export interface XmlAttribute {
  readonly name: string; // qualified
  readonly value: string; // entities decoded
  readonly rawValue: string; // as written, without quotes
  readonly start: number; // of the attribute name
  readonly end: number; // after the closing quote
}
```

### Patches and transactions

```ts
export interface XmlPatch {
  /** Half-open range in XmlPart.text; patches of one part must not overlap. */
  readonly start: number;
  readonly end: number;
  /** Replacement text; a well-formed fragment where an element is expected. */
  readonly text: string;
  /**
   * What the patched position must read back as after the re-scan:
   * "element" — the range now holds exactly one element whose outer XML
   * equals `text`; "content" — the element that starts at `at` (an offset
   * in the unpatched text, mapped through the patches that precede it) has
   * `text` as its content; "attribute" — that element's named attribute
   * reads `value`; "removed" — the parent that starts at `parentAt` has
   * `count` element children, or the removed element's own XML no longer
   * starts where it was. Positions, not tree paths: an insertion before an
   * ancestor in the same transaction must not break the check.
   */
  readonly expect:
    | { readonly kind: "element" }
    | { readonly kind: "content"; readonly at: number }
    | {
        readonly kind: "attribute";
        readonly at: number;
        readonly name: string;
        readonly value: string | undefined;
      }
    | {
        readonly kind: "removed";
        readonly parentAt: number;
        readonly count: number;
        readonly xml: string;
      }
    | { readonly kind: "none" };
}

/** Builders for the common patches; each returns an XmlPatch with its expectation. */
export const patches: {
  replaceContent(part: XmlPart, node: XmlElement, xml: string): XmlPatch;
  replaceElement(part: XmlPart, node: XmlElement, xml: string): XmlPatch;
  removeElement(part: XmlPart, node: XmlElement): XmlPatch;
  insertBefore(part: XmlPart, node: XmlElement, xml: string): XmlPatch;
  insertAfter(part: XmlPart, node: XmlElement, xml: string): XmlPatch;
  appendChild(part: XmlPart, node: XmlElement, xml: string): XmlPatch;
  setAttribute(
    part: XmlPart,
    node: XmlElement,
    name: string,
    value: string,
  ): XmlPatch;
  removeAttribute(part: XmlPart, node: XmlElement, name: string): XmlPatch;
  /** Escapes character data for element content. */
  text(value: string): string;
  /** Escapes an attribute value, double-quoted. */
  attr(value: string): string;
};

export class PackageTransaction {
  /** Patches of one part; applied together, verified, else the whole transaction fails. */
  patch(part: string, items: readonly XmlPatch[]): void;
  /** Adds or replaces a part; `contentType` adds an Override unless a Default already covers the extension. */
  setPart(name: string, bytes: Uint8Array, contentType?: string): void;
  removePart(name: string): void;
  /** Allocates the next free rId of the source's .rels part (created when missing). */
  addRelationship(
    sourcePart: string,
    type: string,
    target: string,
    mode?: "Internal" | "External",
  ): Promise<string>;
  removeRelationship(sourcePart: string, id: string): void;
  /**
   * Stores bytes under `folder` ("/ppt/media/") with a name derived from the
   * content hash and the type's extension, once per distinct content, and
   * relates it to `sourcePart` with the image relationship type unless
   * another is given. Returns the part name and the rId.
   */
  addMedia(
    sourcePart: string,
    folder: string,
    bytes: Uint8Array,
    mimeType: string,
    relationshipType?: string,
  ): Promise<{ readonly part: string; readonly rId: string }>;
  /** A part name that does not exist yet, e.g. uniquePartName("/ppt/slides/slide", ".xml") → "/ppt/slides/slide13.xml". */
  uniquePartName(prefix: string, extension: string): string;
  /** Applies everything or nothing; returns what changed. */
  commit(signal?: AbortSignal): Promise<CommittedChange>;
}

export interface CommittedChange {
  readonly changedParts: readonly string[];
  readonly addedParts: readonly string[];
  readonly removedParts: readonly string[];
  /** Relationships whose internal target no longer exists, one per relationship. */
  readonly warnings: readonly ViewerWarning[];
}

export interface PackageSnapshot {
  /** Opaque; holds the overlay (changed and added parts, removals) by reference. */
  readonly revision: number;
}
```

### Errors

Thrown as `ViewerError` with these codes (new ones join `ViewerErrorCode`):

| Code                  | When                                                                                                                                                                                         |
| --------------------- | -------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------- |
| `invalid-file`        | No end-of-central-directory record, a central-directory entry that does not match its local header, a CRC mismatch on inflate, no `[Content_Types].xml`.                                     |
| `resource-limit`      | The existing ZIP limits (`maxZipEntryBytes`, `maxExpandedOfficeBytes`, `maxInputBytes`), plus more than 65,535 entries or a transaction larger than `maxEditOperations` patches.             |
| `unsupported-package` | ZIP64 sizes or offsets, encryption, a compression method other than stored or deflated, a multi-disk archive.                                                                                |
| `unsupported-part`    | An XML part whose bytes are not UTF-8, or whose decode→encode round trip is not byte-identical (it cannot be patched; it can still be read).                                                 |
| `malformed-xml`       | The scanner cannot parse a part: unbalanced tags, a bad attribute, an undeclared entity in a patched fragment.                                                                               |
| `invalid-patch`       | Overlapping ranges, a range outside the part, a fragment that is not well-formed, or a read-back that does not match the expectation. `details` names the part, the range and what was read. |

Format modules turn these into `OperationIssue`s or `edit-failed` errors with
`details.stage` as the core defines.

## Behaviour

### Opening

- The end-of-central-directory record is found by scanning backwards from
  the end over at most 65,557 bytes (the maximum comment length plus the
  record); the central directory is read entry by entry; each entry's local
  header is read when the entry is first used and must agree with the
  central record on name, method and sizes (flag bit 3 entries take sizes
  and CRC from the central record and keep their data descriptor).
- Entry names are decoded as UTF-8 and normalized to absolute OPC part names
  for lookup (`ppt/slides/slide1.xml` → `/ppt/slides/slide1.xml`); lookup is
  case-insensitive as OPC requires; the raw name is kept for writing.
- `[Content_Types].xml` and `/_rels/.rels` are parsed on open; every other
  part is inflated on first use and cached for the package's life. Inflation
  is metered: bytes beyond the central record's uncompressed size, or beyond
  the limits, stop the stream with `resource-limit`.
- The existing `enforceContainerLimits` keeps running before `open`; `open`
  checks the same limits again on the entries it reads, so a package that
  lies in its central directory is caught once it is used.
- Opening is cheap: a 50 MB deck opens in the time it takes to parse its
  central directory and two small parts, not to inflate every slide.

### Parts and text

- A part is decoded as UTF-8; a byte-order mark is noted and kept. Before a
  part is first patched, the decoded text is encoded back and compared with
  the original bytes; a difference (invalid UTF-8, a declared encoding the
  bytes do not follow) makes the part `unsupported-part`. This is what keeps
  string offsets honest: every patched part is one whose text and bytes map
  one to one, so a patch applied to the text and re-encoded is the same as a
  patch applied to the bytes.
- Ranges in `XmlPart` are UTF-16 code-unit indexes into `text`. Patches are
  applied to the text and the part is re-encoded as UTF-8 once per commit.
  Untouched parts are never decoded for saving.

### Scanning

- The scanner is a hand-written tokenizer over the text: the declaration;
  comments, processing instructions, CDATA sections and a DOCTYPE as opaque
  ranges; start tags with attributes (single or double quotes, entity
  references left raw in `rawValue` and decoded in `value`); end tags;
  self-closing tags; character data. It builds the element tree with exact
  ranges and resolves namespace prefixes through `xmlns` attributes in scope.
  It does not expand entities other than the five predefined ones and
  numeric references, does not read a DTD, and reports an undeclared entity
  in character data as text (parts in the wild contain none; a patched
  fragment with one is `invalid-patch`).
- `mc:AlternateContent`, extension lists, `w14:` / `a16:` and any unknown
  element are ordinary elements: they are scanned, kept, and never
  reordered or rewritten unless a patch targets them.
- A part that does not parse is `malformed-xml`; it can still be copied as
  it is, so a package with one damaged part remains editable elsewhere.
- Scanning is linear in the part's length and does not use regular
  expressions over the whole text. A 10 MB `document.xml` scans in well
  under a second; the number is recorded under Actual result.

### Patching

- Patches of one part are sorted by `start`, checked for overlap and bounds,
  and applied from the end so earlier ranges stay valid. The patched text is
  re-scanned in full; the part must still parse. Then each patch's
  expectation is checked against the new scan: the element at the patched
  position reads back as the fragment (compared as text), the parent's
  content equals the fragment, the attribute reads the value, or no element
  starts there any more. Any failure discards the part's new text and fails
  the transaction with `invalid-patch`; the package is unchanged.
- The builders produce correct ranges and expectations from a scanned node:
  `replaceContent` targets `[contentStart, contentEnd)` (a self-closing
  element is rewritten as an open–close pair), `setAttribute` targets the
  attribute's range or inserts before `>` of the start tag, `removeElement`
  targets `[start, end)` including nothing else, so surrounding whitespace is
  untouched. Callers that need different whitespace write it into the
  fragment.
- Ranges come from the current scan of the part; a patch built from a stale
  scan (the part changed since) is `invalid-patch`, because `XmlPart` carries
  the overlay revision it was scanned at.
- A transaction may touch many parts; parts are patched, replaced, added and
  removed in memory, relationship and content-type parts are regenerated
  from their models, and only then is the overlay replaced, so a failure in
  the last part leaves the first ones untouched.

### Content types and relationships

- `[Content_Types].xml` is scanned like any part, but changes to it are
  made through a small model: `Default` by extension, `Override` by part
  name. Adding a part adds an `Override` unless a `Default` for its
  extension already names the same content type; removing a part removes
  its `Override`. The model is written back as patches to the existing
  elements (a new `Override` is appended before `</Types>`), so an untouched
  file keeps its bytes.
- A relationship part is modelled as an ordered list. `addRelationship`
  allocates `rId<n>` with the smallest `n` not in use; `removeRelationship`
  removes the element. Writes are patches (`insertBefore` the end tag,
  `removeElement`), so existing relationships keep their bytes and order. A
  missing `.rels` part is created with the standard header. Lookups are
  indexed by id and by type, so a part with ten thousand relationships costs
  one scan, not one per edit.
- `resolve` applies OPC target resolution: an internal target is resolved
  against the source part's folder, `..` segments collapse, the result is
  normalized to an absolute part name; external targets are returned as
  written. `relationships` fills `targetPart` for internal targets.

### Media and new parts

- `addMedia` hashes the bytes (SHA-256), looks for a part under `folder`
  with the same hash added earlier in the session, else stores the bytes as
  `<folder>image<n>.<ext>` with the smallest free `n` and the extension
  of the MIME type (`image/png` → `png`, `image/jpeg` → `jpeg`), adds the
  content-type `Default` for that extension when missing, and relates the
  media to the source part. The bytes are stored as they are (an image
  already compresses itself; see decision 2).
- `setPart` with XML bytes for a new part (a slide, a notes part) stores the
  bytes as given; the caller builds the XML, this layer checks that it scans.
- Removed parts are dropped from the archive together with their own `.rels`
  part and their `Override`; relationships pointing at them are the caller's
  to remove (the layer reports dangling internal targets in
  `CommittedChange` warnings so a format module can decide).

### Saving

- With no changes `save()` returns the original bytes (a copy).
- Otherwise a new archive is written: entries in the original order; an
  untouched entry is copied as the exact bytes of its local header, name,
  extra field, data and data descriptor, followed later by its central record
  with only the local-header offset rewritten; a changed entry gets a fresh
  local header and central record (version needed 2.0, flags 0, method per
  `compression`, CRC-32 and sizes computed, the original name, extra field
  and modification time kept); a new entry gets the fixed time 1980-01-01
  00:00 and no extra field; removed entries are left out. The archive comment
  is kept.
- Changed and new entries are **stored** by default (decision 2); with
  `compression: "deflate"` they are deflated through
  `CompressionStream("deflate-raw")`, whose output may differ between
  engines — a host that needs byte-identical saves across browsers keeps the
  default.
- The writer never emits ZIP64; a package whose output would need it fails
  with `unsupported-package` before any bytes are written.
- The result is deterministic for a given package and transaction history on
  one engine, and byte-identical across engines with the default
  compression, which is what the core's round-trip invariants need.

### Snapshots and engines

- The package's mutable state is an overlay: changed and added parts (by
  name → bytes, plus their scanned form), removed names, the regenerated
  content-type and relationship models. The original is never mutated.
  `snapshot()` captures the overlay by reference (overlay values are
  immutable once committed), `restore(snapshot)` swaps it back — O(1).
- A format engine built on this layer keeps `original` + the overlay as its
  state: `restore({ base, batches })` reopens `base` (a checkpoint's bytes)
  as a new package and replays the batches; `materialize("show" | "save")`
  is `save()`; "no changes returns the original bytes" is the layer's own
  rule. Determinism follows from the layer's rules, provided the engine's
  fragments are deterministic (no timestamps or random ids; ids derive from
  `stateId` as the core requires).
- The layer is thread-agnostic (no DOM, no `window`) and runs in a module
  worker, on the main thread and in Node. Where a format module runs it is
  that module's decision; `pptx-edit` is expected to use a worker like the
  PDF engine so inflating and copying a large deck does not block painting.

### Macro and binary parts

- `vbaProject.bin`, `vbaData.xml`, embedded OLE objects, embedded fonts,
  thumbnails, printer settings and media are ordinary parts to this layer:
  copied verbatim unless a transaction replaces them, never parsed, never
  executed. A `.pptm` / `.docm` package keeps its macro-enabled content type
  and its macro parts; nothing here converts formats.

### Limits

- `maxInputBytes`, `maxZipEntryBytes` and `maxExpandedOfficeBytes` as they
  exist; a package with more than 65,535 entries is `resource-limit`.
- A transaction holds at most `maxEditOperations` patches across parts (500
  by default); a format module batches beyond that into several
  transactions inside one `apply`.
- Scanned parts are cached; a package keeps at most the parts an engine
  touched plus the two root parts, so memory is bounded by what the edits
  read, not by the archive.

## Work by layer

### Feature

Proposed tasks for `tasks/plan.md` Phase 7, each a commit with tests:

- **T38 ZIP reader.** Central directory, local headers, flag-bit-3 entries,
  ZIP64 extra fields read (refused as unsupported when they carry sizes),
  metered inflate through `DecompressionStream`, CRC-32 check, limits, OPC
  name normalization. Spike first: `DecompressionStream("deflate-raw")`
  under Node 22 and in the browser matrix; every corpus package's central
  directory and data-descriptor usage.
- **T39 ZIP writer.** Verbatim copy of untouched entries, stored and
  deflated changed entries, deterministic headers, no-change identity,
  rebuild-with-no-changes identity.
- **T40 OPC model.** `[Content_Types].xml`, `.rels` parts, target
  resolution, rId allocation, content-type rules, written back as patches.
- **T41 XML scanner.** Tokenizer, element tree with ranges, namespaces,
  opaque regions, entity decoding, `find` / `at` / `textOf`; fuzz target;
  the UTF-8 round-trip guard.
- **T42 Patches and transactions.** Builders, apply-from-the-end, re-scan
  and read-back verification, atomic commit across parts, snapshots and
  restore, `addMedia`, `setPart`, `uniquePartName`, dangling-target
  warnings.
- **T43 Corpus, browser reopen, docs, gate.** Every corpus PPTX and DOCX
  (and the macro-enabled fixtures of `tests/fixtures`): open → no-change
  save identical → a patch to one part → only that entry differs → every part
  well-formed → `@silurus/ooxml` reopens the bytes in the browser; the
  adversarial manifest; `npm run check`; this spec's Actual result; Linear
  proofs; ACTION-810 Done.

### Tests

- **Unit (`node:test`):** the ZIP reader on hand-built archives (stored,
  deflated, data descriptors, extra fields, a ZIP64 marker, an encrypted
  flag, a bad CRC, a lying central directory) and on the corpus; the writer's
  identities; OPC resolution cases (`../media/image1.png`, absolute targets,
  external targets, case differences); the scanner on hand-written parts
  (every construct above, `mc:AlternateContent`, CDATA, comments between
  elements, entities, namespace scoping, a self-closing root); patch builders
  and verification failures (overlap, stale scan, malformed fragment,
  read-back mismatch); transactions that fail halfway leave the package
  unchanged; media deduplication; limits and every error code.
- **Fuzz:** the ZIP reader and the scanner in `scripts/fuzz-js.mjs` with
  archive and XML seeds, the 100 ms ceiling as for the other targets.
- **Browser (Playwright):** a package patched in Node is loaded into the
  viewer through the normal adapter and renders; `DOMParser` reports no
  `parsererror` for any changed part; the macro-enabled fixture keeps its
  `vbaProject.bin` bytes after a save. The format modules add the end-to-end
  method tests.
- **Round-trip invariants:** no-change save identical; a transaction and its
  snapshot restore identical; the same transaction applied twice to fresh
  packages identical.

### Docs

- This spec's Actual result and Spike results; the roadmap's capability map
  and status.
- No public API changes: `docs/api/editing.md` gains a sentence under a new
  "PPTX and DOCX" heading only when module 04 ships. `docs/architecture.md`
  gets a paragraph on the package layer's place next to the renderers.

## Definition of done

- Opening and saving a package with no edits returns identical bytes, for
  every corpus PPTX and DOCX and the macro-enabled fixtures.
- After a patch, only the touched entries differ (compared entry by entry
  against the original archive), every XML part is well-formed, and
  `@silurus/ooxml` reopens the package in the browser.
- Unit tests cover range patching, relationship and content-type updates and
  media insertion against the corpus files; every error code has a test
  that triggers it; the fuzz targets run clean.
- `vbaProject.bin` survives a save byte for byte and is never read.
- Performance recorded under Actual result: open, scan of the largest part,
  one-part patch and save on the corpus and on a synthetic 500-slide deck.
- `npm run typecheck`, the viewer unit suite, `npm run fuzz:js`,
  `npm run check` and the browser matrix pass; the size report shows no new
  asset.

## Decisions

Approved on 2026-10-02 together with this spec.

1. **Own TypeScript implementation, no new dependency.** The ZIP container
   and the XML scanner are written in this repository: the needs are narrow
   (verbatim copy, lazy inflate, offset-preserving scan) and no library does
   byte-preserving writes; inflation uses the platform
   `DecompressionStream`, within the browser baseline. The Rust `zip` crate
   stays where it is (legacy conversion); a WASM path would add bytes and a
   build step for no gain here.
2. **Changed entries are stored, not deflated, by default.** Output is then
   byte-identical across engines (the round-trip invariants hold everywhere)
   and the writer needs no compression API. The cost is the uncompressed
   size of the parts an edit touched — a few slides or one `document.xml`,
   not the deck. `compression: "deflate"` exists for hosts that prefer
   smaller files over cross-engine identity.
3. **Ranges are string offsets guarded by a byte round trip.** The scanner
   works on decoded text; a part is patchable only when its UTF-8 decode and
   re-encode are byte-identical, so string offsets and byte offsets never
   disagree (GenOffice's `#1579`). Parts that fail the guard are read-only.
4. **Every patch is verified by re-scan and read-back.** The cost is one
   extra scan per patched part per commit; the gain is that a malformed
   fragment, a stale range or a wrong expectation can never reach a saved
   file.
5. **The package state is an overlay over immutable originals.** Snapshots
   are O(1), `restore` is a swap, and the core's checkpoint model maps onto
   it without copying.
6. **SuperDoc: ideas only, no code, no dependency.** `superdoc/docx-editor`
   is AGPL-3.0, its V2 engine is proprietary with a licence that forbids
   analysis and modification, and both are DOCX-only; web-doc is
   MIT OR Apache-2.0 and must stay usable by its hosts without an AGPL
   obligation. A line-by-line rewrite would be a derivative work of AGPL
   code. What is taken is the published architectural conclusion (a DOCX is
   a package of parts, not one editor tree), which this layer already
   follows. Its MIT-licensed neighbours — `superdoc/docx-corpus` (736K
   public `.docx` files) and `@docfonts/fallbacks` — may serve module 06 for
   robustness testing and font fallbacks, after the usual licence check.
7. **ZIP64 and encrypted packages are refused**, on open and on save, with
   `unsupported-package`; the existing size limits make ZIP64 unreachable in
   practice.
8. **The layer has no worker of its own.** It is thread-agnostic; each format
   module decides where it runs (the PDF engine's worker pattern is the
   expected choice).
9. **Dangling relationship targets warn, they do not refuse.** A transaction
   that removes a part reports relationships that still point at it in
   `CommittedChange.warnings`, so a format module that forgets to remove
   them gets the signal while one that removes them in the same transaction
   is not blocked.

## Open questions

None. The two questions of the draft were decided on 2026-10-02: changed
entries are stored (decision 2 stands), and dangling relationship targets
are reported as warnings in `CommittedChange`, not refused (decision 9).

## Spike results

Task 38, 2026-10-02, Node 22.23 on an Apple M4 Pro
(`packages/viewer/test/ooxml-zip.test.ts`):

- **Compression streams.** `DecompressionStream` and `CompressionStream`
  are globals in Node 22; `"deflate-raw"` inflates every corpus entry with a
  matching CRC. The browser matrix proves the same path in task 43.
- **Corpus packages** (`sample.pptx`, `sample.docx` from Apache POI;
  `chart-point-colors.pptx`, `oversized-inline-image.docx` from
  `tests/fixtures`): every entry is stored or deflated, none uses a data
  descriptor, one entry of `sample.pptx` carries an extra field, no archive
  has a comment, and `sample.pptx` is the only one with a stored entry (its
  JPEG thumbnail). Reading the central directory takes under 0.2 ms for all
  four; inflating every part takes 5–7 ms per package (46 parts, 107 KB for
  the deck; 18 parts, 830 KB for the image-heavy DOCX).
- **Shapes the reader refuses**, each with a test: a ZIP64 locator or entry
  count, a ZIP64 size marker, the encryption flag, a method other than stored
  or deflated, no end record, a truncated directory, a local header whose
  name, method or sizes disagree with the central record, a data descriptor
  that disagrees, a CRC mismatch, corrupt deflate data, an entry over
  `maxZipEntryBytes`, and an entry whose every header understates its size
  (caught by metering during inflation, as `resource-limit`).

## Actual result

Implemented 2026-10-02 in `packages/viewer/src/edit/ooxml/` (`zip.ts`,
`writer.ts`, `names.ts`, `opc.ts`, `xml.ts`, `patch.ts`, `transaction.ts`,
`package.ts`), tasks 38–43 of `tasks/todo.md`:

- **Container.** `parseZip` reads the end record and the central directory;
  local records are located on first use and checked against the directory
  (data descriptors included); `inflateEntry` meters the platform
  `DecompressionStream` against the declared size and checks the CRC-32.
  `writeZip` copies untouched entries' local and central records byte for
  byte (the offset rewritten), writes changed entries stored — deflated on
  request — with their name, extra field, time and attributes kept, appends
  new entries with the fixed 1980 time, drops removed ones, keeps the
  comment, and refuses ZIP64 output before writing.
- **Package.** `OoxmlPackage` holds the original, lazily inflated parts, an
  overlay of replaced, added and removed parts, O(1) snapshots, scanned
  parts cached per overlay revision, content types and relationships as
  models, and `save()` that returns the original bytes without changes.
- **Scanner.** A hand-written tokenizer over the decoded text; every element
  carries the exact range of its tags, content and attributes; namespaces
  are resolved in scope; declaration, comments, PIs, CDATA and DOCTYPE are
  opaque ranges; predefined and numeric entities are decoded; non-UTF-8
  parts are `unsupported-part`, unparsable ones `malformed-xml` with the
  offset.
- **Patches.** Builders for content, element, insertion, append and
  attribute changes; patches apply from the end, the part is re-scanned and
  every expectation checked by element positions mapped through the
  preceding patches (a refinement of the draft's tree paths: an insertion
  before an ancestor in the same transaction must not break the check).
- **Transactions.** Part patches, replaced, added and removed parts,
  relationship adds and removes (`rId` allocation fills the smallest gap),
  content-type Overrides and Defaults, media stored once per SHA-256 and
  related from the source part, unique part names; everything is computed
  and verified in memory, then committed to the overlay at once; a removed
  part takes its own `.rels` part and Override along; relationships still
  pointing at a removed part come back as `fidelity-degraded` warnings with
  `details.reason: "dangling-relationship"` (decision 9). A transaction
  older than the package's revision is an `edit-conflict`.
- **Differences from the draft API**, all recorded in the API section:
  `addRelationship` and `addMedia` return promises (they read the
  relationships part); patch expectations use positions, not paths;
  `CommittedChange.warnings` carries the dangling-target warnings.
- **Tests** (`ooxml-zip`, `ooxml-writer`, `ooxml-xml`, `ooxml-opc`,
  `ooxml-patch`, `ooxml-package`; 29 tests): every accepted and refused
  archive shape; no-change rebuild byte-identical for hand-built archives
  and every corpus package; one changed part leaves every other entry's
  local and central records untouched; every XML part of every corpus
  package scans with ranges that reproduce the source; every builder's
  read-back; overlapping, stale, malformed and mismatching patches;
  content types and relationships written as patches that keep their bytes;
  media deduplication; snapshots; a transaction whose last part fails; a
  macro-enabled package whose `vbaProject.bin` survives untouched and is
  never read; the full cycle on every corpus and fixture package;
  `tests/e2e/ooxml-package.spec.ts` loads packages patched by the layer into
  the viewer, renders them, finds the edit in the extracted text and parses
  every changed part with the browser's `DOMParser`. The ZIP reader and the
  scanner are `fuzz:js` targets. Five cases joined the adversarial manifest.
- **Performance**, Node 22 on an Apple M4 Pro, one text replaced in one
  part:

  | Package                     |      Size |   Open |        Scan of the part | Commit |   Save |
  | --------------------------- | --------: | -----: | ----------------------: | -----: | -----: |
  | sample.docx (19 parts)      |  14,860 B | 0.5 ms |   `document.xml` 5.1 ms | 1.1 ms | 0.5 ms |
  | sample.pptx (46 parts)      |  39,083 B | 0.1 ms |     `slide1.xml` 0.3 ms | 0.5 ms | 0.1 ms |
  | chart-point-colors.pptx     |  41,693 B | 0.1 ms |   `slide1.xml` < 0.1 ms | 0.4 ms | 0.1 ms |
  | oversized-inline-image.docx |  37,126 B | 0.1 ms | `document.xml` < 0.1 ms | 0.3 ms | 0.1 ms |
  | synthetic 500-slide deck    | 364,496 B | 1.0 ms |   `slide250.xml` 0.9 ms | 1.2 ms | 1.4 ms |

  The largest corpus part, the 438 KB `stylesWithEffects.xml` of the image
  DOCX, scans in 7.6 ms. A saved package grows by the stored size of the
  parts an edit touched (decision 2): 1.5–2 KB on these files.
