# Module 07. `ai-edit` — AI tooling over the edit sessions

**Status:** Done 2026-10-02 (T59–T64; approved the same day with the
recommended answers to the open questions, decisions 8–10), Linear
ACTION-858 under ACTION-723 (web-doc); release pending. Everything here is additive to the
`EditSession` contract of module 01 and to the three format sessions:
nothing a host or the Operators shell uses today changes shape.

## Goal

Let an AI agent read and change a PDF, PPTX or DOCX file through the same
operations a person uses, safely and with a review path: a compact outline
the model can reason about, a way to turn "the paragraph that starts with
Revenue" into an element id, tool definitions an LLM can call directly
(with the session validating and executing them), suggestion mode where the
format has it (Word tracked changes), and checkpoints so a host can roll
back an agent's turn in one call. The host (Operators, a CLI, an MCP server)
owns the model, the prompt and the loop; web-doc owns reading, validation,
application and the receipts.

## Requirement sources

- [`../00-roadmap.md`](../00-roadmap.md): capability map row 07 ("document
  outline, target resolution from text or citations, tool schemas,
  suggestion mode (tracked changes where the format has them),
  checkpoints"); fixed decision 1 (one operation vocabulary for the UI and
  the AI); the hard limits.
- [`./01-edit-core.md`](./01-edit-core.md): out of scope "AI-specific
  helpers such as outlines and target resolution (module 07)"; R2
  checkpoints retained inside `maxEditCheckpointBytes`; `ApplyOptions.
  author` reserved for formats that keep authorship; `session.schemas`,
  `dryRun`, `expectedRevision`, `expectedSessionId`, `findText()` with
  `ranges`, `OperationIssue` with `operationIndex` and `path`.
- [`../../api/editing.md`](../../api/editing.md) "Guidance for AI clients":
  read before writing, send `expectedRevision`, dry-run first, batch,
  treat issues as structured feedback. This module turns that guidance
  into API.
- Linear ACTION-723: "AI agent editing of these files. The same operation
  API is designed for it (`ai-edit` module), and it will be a follow-up
  ticket."
- Research of 2026-10-01 (ideas only, no code; GenOffice attributed in
  `THIRD_PARTY_NOTICES.md`, SuperDoc never copied):
  - GenOffice (Apache-2.0): the model sees a numbered skeleton of blocks,
    never the file's XML; a write against stale indexes is refused by an
    optimistic "document changed since it was read" check; AI edits land
    as Word tracked changes authored by the AI; every agent turn has a
    snapshot for a one-click rollback.
  - SuperDoc (AGPL, ideas only): an agent preset of two tools, inspect and
    act, over a validated list of actions, with post-checks on the
    receipts; `changeMode: 'direct' | 'tracked'`.
- ECMA-376 Part 1 (WordprocessingML) 17.13.5: `w:ins` and `w:del` wrap
  runs with `w:id`, `w:author` and `w:date`; deleted text is `w:delText`;
  a deleted paragraph mark is `w:rPr/w:del` on the mark; property changes
  are `w:rPrChange` and `w:pPrChange` holding the previous properties.

## Dependencies

- Modules 01–06 done and released (web-doc 0.7.0).
- No new runtime dependency. Fuzzy matching reuses the viewer's citation
  search scoring (`src/search`); JSON Schema generation reuses the
  per-operation schemas every session already exports.

## In scope

### Outline and description

```ts
interface OutlineOptions extends ReadOptions {
  readonly pageRange?: readonly [number, number];
  readonly kinds?: readonly string[];
  /** Characters of text kept per node; the rest is cut and `truncated` says so. Default 160. */
  readonly maxTextChars?: number;
}

interface OutlineNode {
  readonly id: string; // the element id, as every operation takes it
  readonly ordinal: string; // "3", "3.2" (a cell paragraph under its table): the reading-order path
  readonly kind: string; // the format's element kind
  readonly pageIndex: number; // −1 while the page is not laid out (DOCX without a cached page)
  readonly text?: string; // first `maxTextChars` characters of the element's text
  readonly textLength: number;
  readonly truncated: boolean;
  readonly label?: string; // what a person calls it: "Title 1", "Heading 2", "Table (3×4)", "Picture"
  readonly operations: readonly string[];
  readonly children?: readonly OutlineNode[]; // a table's cell paragraphs, a group's children
}

interface DescribeOptions extends OutlineOptions {
  /** Characters the description may take; the tail is cut and `truncated` says so. Default 50 000. */
  readonly maxChars?: number;
}

interface DocumentDescription {
  readonly format: EditableFormat;
  readonly pageCount: number;
  readonly elementCount: number;
  readonly text: string; // one line per node: `[sld2:7] slide 2 shape "Title 1": Quarterly review`
  readonly truncated: boolean;
}

interface EditSession {
  getOutline(options?: OutlineOptions): Promise<ReadResult<OutlineNode>>;
  describe(options?: DescribeOptions): Promise<ReadItem<DocumentDescription>>;
}
```

- `getOutline()` lists the body elements in reading order with the ids the
  operations take, a short text, a human label and the operations each one
  accepts. It is `getElements()` reshaped for a prompt: no geometry, no
  styles, nested where the format nests (table cells, groups).
- `describe()` renders the outline as plain text, one line per node, in a
  fixed grammar the docs state, so a host can paste it into a prompt
  without its own formatter. Page sizes and the format go first. The
  budget is characters, not tokens: hosts count tokens themselves.
- Both are reads: envelopes with `sessionId` and `revision`, queued behind
  changes like every read, cancellable with `signal`.

### Target resolution

```ts
interface TargetQuery {
  /** Text to find; whitespace-insensitive, case-insensitive, tolerant to small differences. */
  readonly text?: string;
  /** A citation as the Operators viewer gets them: the passage and the page it is expected on. */
  readonly citation?: { readonly text: string; readonly pageNumber?: number };
  readonly pageIndex?: number;
  readonly kinds?: readonly string[];
  /** Restrict to an element and its descendants (a table, a group). */
  readonly within?: string;
  readonly maxResults?: number; // default 5
}

interface TargetCandidate {
  readonly elementId: string;
  readonly range?: TextRange; // the matched part, for ranged operations
  readonly pageIndex: number;
  readonly score: number; // 1 exact, down to 0.5 for the loosest accepted match
  readonly reason: "exact" | "normalized" | "fuzzy" | "kind-only";
  readonly snippet: string; // the matched text with a little context
}

interface EditSession {
  resolveTargets(query: TargetQuery, options?: ReadOptions): Promise<ReadResult<TargetCandidate>>;
}
```

- Resolution runs in three passes and stops at the first that yields:
  exact `findText()`, then a normalized match (whitespace folded, case and
  quotes folded, NFKC), then a fuzzy match with the viewer's citation
  scoring (the one `search()` uses for a citation request) over the
  elements' text, bounded by `pageIndex`, `kinds` and `within`. A query
  with `kinds` only ("the first table on slide 3") is a kind lookup in
  reading order.
- Every candidate carries a `range` when the match is a part of the
  element's text, so `replaceText` and `setTextStyle` can take it as it is.
- A `citation` query is the viewer's citation request in element terms:
  `pageNumber` is a hint, not a filter, as in `search()`.

### Checkpoints

```ts
interface EditCheckpoint {
  readonly id: string; // 22 URL-safe characters, unique in the session
  readonly label?: string;
  readonly revision: number; // the state it names
  readonly createdAt: string; // ISO 8601
}

interface EditSession {
  createCheckpoint(label?: string): Promise<EditCheckpoint>;
  listCheckpoints(): readonly EditCheckpoint[];
  /** Back to the checkpoint's state; one undo step, like `reset()` is. */
  restoreCheckpoint(id: string, options?: HistoryOptions): Promise<EditReceipt>;
  dropCheckpoint(id: string): void;
}
```

- A checkpoint pins a state the host names, typically before an agent's
  turn. Module 01 already retains bytes for some states; a named
  checkpoint is pinned: never evicted by `maxEditCheckpointBytes` while it
  exists, counted against the budget, and rebuilt by replay from the
  original when the budget cannot hold its bytes (the content is the same
  either way; module 01's guarantee).
- `restoreCheckpoint()` is a change: it lands as one history entry
  (undoable), bumps the revision, emits `editstatechange` and
  `documentchange` with every page the restore touched, and reports the
  elements the restore removed and recreated in its receipt. A checkpoint
  survives `undo()`, `redo()` and `reset()`; `reset()` drops none.
- `maxEditCheckpoints` (default 20): creating one past the limit rejects
  with `resource-limit`.

### Tool definitions and execution

```ts
interface ToolDefinition {
  readonly name: string;
  readonly description: string;
  readonly inputSchema: JsonSchema; // draft 2020-12
}

interface ToolSet {
  readonly version: number; // raised when a tool's shape changes incompatibly
  readonly format: EditableFormat;
  readonly definitions: readonly ToolDefinition[];
}

interface ToolCall {
  readonly name: string;
  readonly arguments: unknown; // the model's JSON, validated against the tool's schema
}

interface ToolResult {
  readonly ok: boolean;
  /** JSON for the model: an outline, candidates, a receipt, or the issues of a refused call. */
  readonly content: unknown;
  /** One or two sentences for the model and the chat: "Replaced the title of slide 2." */
  readonly text: string;
  readonly issues?: readonly OperationIssue[];
}

interface ToolCallOptions {
  readonly expectedRevision?: number; // checked before any tool that writes
  readonly changeMode?: "direct" | "tracked";
  readonly author?: string;
  readonly signal?: AbortSignal;
}

interface EditSession {
  readonly tools: ToolSet;
  callTool(call: ToolCall, options?: ToolCallOptions): Promise<ToolResult>;
}
```

The tool set of every format (names are binding):

| Tool                  | Arguments                                                                    | What it does                                                                                                          |
| --------------------- | ---------------------------------------------------------------------------- | --------------------------------------------------------------------------------------------------------------------- |
| `document_describe`   | `OutlineOptions` + `maxChars`                                                | `describe()`; the usual first call of a turn.                                                                         |
| `document_outline`    | `OutlineOptions`                                                             | `getOutline()`, as JSON.                                                                                              |
| `document_find`       | `TargetQuery`                                                                | `resolveTargets()`.                                                                                                   |
| `document_inspect`    | `{ id }`                                                                     | `getElement()` with text, style and table rows, geometry left out.                                                    |
| `document_preview`    | `{ operations: Operation[] }`                                                | `apply()` with `dryRun: true`: the receipt an apply would give, or the issues.                                        |
| `document_apply`      | `{ operations: Operation[], label? }`                                        | `apply()`; `expectedRevision` from `ToolCallOptions`; `changeMode` as below.                                          |
| `document_undo`       | `{}`                                                                         | `undo()`.                                                                                                             |
| `document_checkpoint` | `{ action: "create" \| "restore" \| "list", id?, label? }`                   | The checkpoint methods.                                                                                               |

- `operations` validate against the session's operation schemas as one
  `oneOf`, so a model gets the same `invalid-operation` issues as any
  client, with `operationIndex` and `path`. Binary payloads are asset
  references (`asset:<sha-256>`) that the host registered with
  `addAsset()`; the schema says so, a raw base64 payload over a tool call
  is refused with `unknown-asset` guidance in `text`.
- `callTool()` never throws for a model's mistake: a bad name, bad
  arguments, a refused batch or a conflict come back as `ok: false` with
  `issues` and a `text` the model can act on. It throws only for the
  session's own errors (`lifecycle-error`, `aborted`).
- `ToolResult.text` is generated from the receipt by `describeReceipt()`
  (exported): what changed, where, and what it created, in the format's
  words ("Added a text box on page 1 (p0:n2)", "Deleted slide 3").
- The definitions are plain JSON Schema; a host maps them to its model
  provider's tool format (OpenAI `parameters`, Anthropic `input_schema`)
  without a web-doc dependency on either.

### Suggestion mode: tracked changes in DOCX

```ts
interface ApplyOptions {
  readonly changeMode?: "direct" | "tracked"; // default "direct"
  readonly author?: string; // module 01's reserved field, now read: the `w:author` of tracked changes
}
```

- With `changeMode: "tracked"` the DOCX engine writes the change as a
  revision instead of replacing content: `replaceText` wraps the removed
  runs in `w:del` (text as `w:delText`) and the new runs in `w:ins`;
  `insertParagraph` writes the paragraphs inside `w:ins` with an inserted
  paragraph mark; `deleteElement` of a paragraph marks its runs and its
  mark deleted (the paragraph stays until Word accepts); `setTableCell` is
  a tracked `replaceText` of the cell's first paragraph; `setTextStyle`
  and `setParagraphStyle` write `w:rPrChange` / `w:pPrChange` with the
  previous properties. `moveElement`, `insertTable`, `insertImage` and
  `deleteElement` of a table or picture have no tracked form in this
  module and are refused in tracked mode (`unsupported-change-mode`).
- `w:author` is `ApplyOptions.author`, required in tracked mode (decision
  8); `w:date` is the batch's `timestamp` (module 01); `w:id` is unique in
  the part. Word and
  Pages show the result as suggestions to accept or reject.
- A paragraph holding a tracked change is read-only for direct edits
  (module 06's rule) and stays so after the session's own tracked edit:
  the next tracked edit of the same paragraph is refused with
  `invalid-target` until the change is accepted or rejected in Word. This
  keeps the engine's paragraph rebuild local and the revisions unambiguous;
  lifting it is a later module.
- PDF and PPTX have no tracked changes: `changeMode: "tracked"` is refused
  with `unsupported-change-mode`, and the host reviews with checkpoints
  (`document_checkpoint` before the turn, `restore` to reject). The tool
  set's `document_apply` description says which mode the format has.
- Reading: a tracked paragraph's `text` stays the accepted text (inserted
  runs in, deleted runs out, as module 06 reads today); `readOnlyReason:
  "tracked-changes"` marks it. A new read, `getRevisions(elementId)`, lists
  the paragraph's revisions (`ins` / `del` / `rPrChange` / `pPrChange`,
  author, date, text) so a host can show what the AI proposed.

### Limits and performance

| Limit                | Default | Meaning                                             |
| -------------------- | ------- | --------------------------------------------------- |
| `maxOutlineNodes`    | 5 000   | Nodes a `getOutline()` returns; more is `truncated` |
| `maxDescribeChars`   | 200 000 | Upper bound of `DescribeOptions.maxChars`           |
| `maxEditCheckpoints` | 20      | Named checkpoints alive at once                     |

`describe()` on a 500-page DOCX or a 500-slide deck resolves within the
read budget of the format (one body scan, cached pages only for geometry,
no layout), measured in the browser matrix and recorded in the docs.

## Out of scope

- Running a model, prompts, an agent loop, MCP or CLI surfaces: hosts.
- Tracked changes in PPTX and PDF (the formats have none); comments as a
  suggestion channel (a later module).
- Accepting or rejecting tracked changes inside web-doc (Word's job for
  now; a later module may add `acceptRevision`).
- Semantic edits ("rewrite this section in a friendlier tone"): the model
  writes text, web-doc applies it.
- Summaries, embeddings, retrieval over the document's text.

## Behaviour

### Reads and revisions

Outline, description and target resolution are reads in the session queue:
a read queued behind an `apply()` describes the document after it, and the
envelope's `revision` is what a model hands back as `expectedRevision`. A
`callTool` with a stale `expectedRevision` returns `ok: false` with an
`edit-conflict` issue and a `text` that says to describe again.

### Ids in prompts

Element ids are the stable ids of the formats (`p0:o3`, `sld3:7`,
`p:1A000000`, `tbl:…`): never renumbered in a session, never reused after a
deletion, the same after undo and redo. The description prints them in
brackets at the start of every line so a model copies them verbatim; the
tool schemas describe the id as an opaque string.

### Tracked changes and the display copy

A tracked DOCX edit changes the saved bytes exactly like a direct one does
elsewhere in the paragraph (one paragraph rebuilt, the rest byte-identical),
so the viewer reopens it through module 05's display pre-pass; the 0.88
renderer draws inserted runs and leaves deleted runs out, which is what
Word shows with "No markup". Showing revision marks in the viewer is the
renderer's job and out of scope.

### Checkpoints and history

A checkpoint is a pinned history state. Restoring it appends a history
entry whose "after" is the checkpoint's content; `undo()` then returns to
the state before the restore. Module 01's checkpoint retention keeps the
bytes when the budget allows; otherwise the restore replays from the
original up to the checkpoint's revision (its batches are in the history
or folded into the starting point, which keeps their effect). Saved
checkpoints do not survive a reload: they are session state, like the
history.

## Work by layer

### Feature

- `src/edit/ai/outline.ts`: `getOutline` over `getElements()` of every
  format (nesting from `parentId`, labels from the format's element,
  ordinals in reading order), `describe` with the fixed line grammar.
- `src/edit/ai/targets.ts`: the three-pass resolution over `findText()` and
  the elements' text, reusing `src/search` scoring.
- `src/edit/session.ts`: named checkpoints over the R2 retention (pins,
  eviction order, `restoreCheckpoint` as a history entry, the limit).
- `src/edit/ai/tools.ts`: `ToolSet` built from `session.schemas`,
  `callTool` dispatch with validation, `describeReceipt`.
- `src/edit/docx/tracked.ts`: tracked writers for the five operations
  (`w:ins`, `w:del`, `w:delText`, `w:rPrChange`, `w:pPrChange`, ids,
  author, date), `getRevisions`; `changeMode` plumbing through the worker.
- `src/index.ts`: the new types; `ApplyOptions.changeMode`.

### Tests

- Unit: outline nesting and ordinals per format (PDF, PPTX, DOCX fixtures
  of the existing suites), `describe` grammar and truncation, the three
  resolution passes with scores, checkpoint create/list/restore/drop and
  eviction with pins, `callTool` for every tool and every refusal path,
  `describeReceipt` wording, tracked DOCX writers at the XML level (ECMA
  element order, ids, author, date) and `getRevisions`.
- Browser (`tests/e2e/edit-ai.spec.ts`, in the matrix): a scripted agent
  turn over each format's fixture — describe, find, preview, apply,
  checkpoint, restore — through `callTool` only; a tracked DOCX edit saved
  and reopened shows the accepted text; `describe()` latency on the 500-
  page and 500-slide synthetic documents.
- Fixtures: `npm run fixtures:docx` gains a tracked-changes document for the
  manual Word/Pages check (suggestions visible, accept and reject work).

### Docs

- `docs/api/ai-editing.md`: the outline grammar, the tools with an example
  turn, checkpoints, tracked changes, limits; `docs/api/editing.md`
  "Guidance for AI clients" points to it; roadmap row 07 and the decisions
  log; `THIRD_PARTY_NOTICES.md` (GenOffice ideas).

## Definition of done

- `getOutline()`, `describe()`, `resolveTargets()`, the checkpoint methods,
  `tools` and `callTool()` work on PDF, PPTX and DOCX sessions, in the
  worker-backed sessions of the viewer and through the tool path only.
- A tracked DOCX edit opens in Word and Pages as a suggestion with the
  given author; accepting it yields the direct edit's text; rejecting
  restores the original; untouched paragraphs stay byte-identical.
- The browser agent turn passes in the matrix; `describe()` on the 500-
  page and 500-slide documents is recorded in the docs.
- Unit tests, the browser matrix, `npm run check`, the size and licence
  gates are green; the docs listed above are written; a web-doc release is
  published and ACTION-723 gets its AI follow-up ticket closed.

## Decisions

1. **Hosts own the model; web-doc owns the tools.** The module ships
   definitions and an executor, never a prompt or a loop, so Operators,
   a CLI and an MCP server share one validated path.
2. **The description is a line grammar, not Markdown or HTML.** One line
   per element with the id first keeps the mapping back to operations
   exact and the budget predictable; a host wanting prose renders the
   outline itself.
3. **Target resolution degrades in named steps** (exact, normalized,
   fuzzy, kind-only) and reports which step matched, so a host can decide
   when to ask the person instead of trusting a loose match.
4. **Named checkpoints are pinned history states, restored as a change.**
   No second history, no bytes outside module 01's retention; a restore is
   undoable like any batch.
5. **Suggestion mode is Word tracked changes, nothing synthetic for PDF and
   PPTX.** Those formats review through checkpoints; inventing a revision
   layer the file cannot carry would not survive a save.
6. **A paragraph with a tracked change is read-only until Word decides.**
   The paragraph rebuild stays local and the revisions unambiguous; a
   later module may add accept/reject inside web-doc.
7. **Tool arguments carry assets by reference.** Registered through
   `addAsset()` by the host; a model never pastes base64 into a call.
8. **Tracked changes name their author explicitly** (2026-10-02, Leonid):
   no default; a tracked batch without `author` is refused with a
   `required` issue at `/author`, so a file never carries an anonymous
   suggestion.
9. **No preview is required before an apply** (2026-10-02, Leonid): the
   model decides; the docs recommend `document_preview` for batches above
   one operation, and `document_apply`'s description says so.
10. **`maxEditCheckpoints` stays 20** (2026-10-02, Leonid), raised through
    `ResourceLimits` by a host that needs more.

## Open questions

None: the three questions of the draft were decided on 2026-10-02
(decisions 8–10).

## Actual result

Done 2026-10-02 on `feat/ai-edit` (T59–T64, ACTION-858); every method of
the spec works on PDF, PPTX and DOCX sessions, in the worker-backed sessions
of the viewer and through the tool path only. Deviations from the draft:

- `getOutline()` resolves to `OutlineResult`, a `ReadResult<OutlineNode>`
  with `nodeCount` and `truncated` added, so the node limit is reported
  without a second call. `OutlineNode` gains `readOnlyReason` and `hidden`.
- The description's header names the format, the page count and the
  element count; page sizes are not part of it, since a session does not
  know them (the viewer's `DocumentInfo.pageSizes` does).
- `describeReceipt()` takes the format, the operations and the receipt.
- `restoreCheckpoint()` reuses the checkpoint's state id for its history
  entry, so `dirty` is exact after restoring a saved state; its receipt's
  `createdIds` and `removedIds` are the net of the batches between the two
  states; unknown ids reject with `invalid-operation` (the tool answers
  `unknown-checkpoint`).
- `documentchange` gains the reason `restore`.
- A paragraph with changed properties only (`w:rPrChange`, `w:pPrChange`)
  stays editable, as module 06 decided; insertions, deletions and moves on
  runs or on the paragraph mark lock it. A tracked change that would touch
  a hyperlink, content control or field is refused
  (`unsupported-change-mode`), since a revision cannot wrap them.
- The write mode (`changeMode`, `author`, `timestamp`) travels with the
  engine batch and the history entry, so replays reproduce tracked bytes.
- `describe()` latency (Chromium): 8 ms on 500 DOCX pages, 53 ms on 500
  slides, 72 ms on 500 PDF pages; recorded in `docs/api/ai-editing.md`.
- The manual Word and Pages check of `artifacts/docx-fixtures/tracked-changes.docx`
  (`npm run fixtures:docx`) is Leonid's, like the module 06 fixtures.
