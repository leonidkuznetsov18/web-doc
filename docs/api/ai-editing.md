# AI editing

The [editing API](./editing.md) is a JSON protocol a model can drive: ids,
operations, receipts and issues are all plain data. This page covers what an
edit session adds for an AI client on top of it: a document outline and a
plain-text description shaped for a prompt. The package ships no model, no
prompt and no agent loop; a host owns those and calls the session.

Every method here is a read in the session's queue: it runs after the calls
made before it, describes the document as it is then, and carries the usual
envelope (`sessionId`, `revision`). The `revision` is what a model hands back
as `expectedRevision` with its next batch.

## Outline

```ts
const session = await viewer.edit();
const outline = await session.getOutline({ pageRange: [0, 2] });
for (const node of outline.items)
  console.log(node.ordinal, node.id, node.kind, node.label, node.text);
```

`getOutline()` is `getElements()` reshaped for a prompt: the body elements in
reading order, nested where the format nests, with the ids the operations
take, a short text, a human label and the operations each one accepts. No
geometry and no styles; a model that needs them calls `getElement(id)`.

| Field            | Meaning                                                                                                                      |
| ---------------- | ---------------------------------------------------------------------------------------------------------------------------- |
| `id`             | The element id, as every operation takes it.                                                                                 |
| `ordinal`        | The reading-order path: `"3"`, `"3.2"` (the second cell paragraph of the third element). Assigned before any filter applies. |
| `kind`           | The format's element kind.                                                                                                   |
| `pageIndex`      | 0-based; `-1` while the page is not laid out (a DOCX page the viewer has not cached).                                        |
| `text`           | The first `maxTextChars` characters of the element's text (default 160); absent for an element without text.                 |
| `textLength`     | Length of the whole text.                                                                                                    |
| `truncated`      | `text` is shorter than the element's text.                                                                                   |
| `label`          | What a person calls it; see below.                                                                                           |
| `readOnlyReason` | Why the element only accepts insertions next to it (a DOCX paragraph with tracked changes, a section break).                 |
| `hidden`         | Listed and editable, but not drawn (a hidden PowerPoint shape).                                                              |
| `operations`     | Names of the operations that accept the element as their target.                                                             |
| `children`       | A table's cell paragraphs, a group's members, a paragraph's inline pictures.                                                 |

Labels come from the format: a deck shape's name as the selection pane shows
it (`Title 1`), else its placeholder (`Title`, `Subtitle`, `Content`); a
Word paragraph's style (`Heading 1`, `List Paragraph`) or its list membership
(`List item (level 1)`); a table's size everywhere (`Table (3×4)`, appended
to a deck table's name). An element without a telling name has no label.

Options:

- `pageRange`: inclusive 0-based range. A node outside it is left out
  unless a descendant is inside; the container then stays as its parent.
- `kinds`: kinds to keep, with the same container rule.
- `maxTextChars`: characters of text per node; default 160.
- `signal`: cancels the read like any other.

The result adds `nodeCount` (nodes returned, nested ones included) and
`truncated` to the envelope: `maxOutlineNodes` (default 5 000) cuts the
outline in reading order, subtrees included, and `truncated` says so.

## Description

```ts
const { item } = await session.describe({ maxChars: 20_000 });
prompt.push(item.text);
```

`describe()` renders the outline as plain text in a fixed grammar, so a host
pastes it into a prompt without a formatter of its own. One line per node,
the id first so a model copies it verbatim:

```text
pptx: 12 slides, 87 elements
[sld1:2] slide 1 shape "Title 1": Quarterly review
[sld1:3] slide 1 group "Group 2"
  [sld1:4] slide 1 shape "Note 3": Inside the group
[sld1:5] slide 1 table "Table 4 (2×3)": Region | Q1 | Q2 ⏎ North | 120 | 130
[p:1A000000] page 1 paragraph "Heading 1": Introduction
[p:1A000003] paragraph (read-only: tracked-changes): A suggested sentence…
```

The grammar:

```text
<format>: <n> pages|slides, <m> elements
[<id>] page|slide <number> <kind> "<label>" (<flags>): <text>
```

- The header names the format, the page count (`slides` for a deck) and the
  number of elements the description covers.
- The page is 1-based and left out while the node has none (`pageIndex` is
  `-1`); the label and the flags are left out while the node has none; the
  text and its colon are left out while it is empty.
- Flags are `read-only: <reason>` and `hidden`, comma-separated.
- The text is one line: tabs (cell separators) become `|`, line breaks
  `⏎`, other control characters a space; an inline picture or a field in
  a Word paragraph is its placeholder character (U+FFFC); a text cut by
  `maxTextChars` ends in `…`.
- Children are indented two spaces per level.
- Double quotes inside a label become single quotes.

The budget is characters, not tokens; hosts count tokens themselves.
`maxChars` (default 50 000, at most `maxDescribeChars`, default 200 000)
cuts whole lines from the end; the last line then reads `… (<n> more
lines)`. When `maxOutlineNodes` cut the outline instead, it reads
`… (outline cut at <n> nodes)`. `DocumentDescription.truncated` is set in
both cases, `elementCount` counts the nodes the outline holds, `pageCount`
is the session's.

`DescribeOptions` takes every `OutlineOptions` field, so a description of one
page or of the tables only is one call.

## Target resolution

```ts
const { items } = await session.resolveTargets({
  citation: { text: "covers three regions", pageNumber: 2 },
});
const best = items[0];
if (best && best.score >= 0.9)
  await session.replaceText({
    target: best.elementId,
    ...(best.range ? { range: best.range } : {}),
    text: "covers four regions",
  });
```

A model names what it wants to edit by quoting it, by citing it with a page,
or by kind. `resolveTargets()` turns that into element ids and ranges in
named passes and stops at the first that yields, so a host knows how far to
trust a candidate and when to ask the person instead.

| Query field  | Meaning                                                                                                                                                              |
| ------------ | -------------------------------------------------------------------------------------------------------------------------------------------------------------------- |
| `text`       | Text to find; whitespace-insensitive, case-insensitive, tolerant to small differences.                                                                               |
| `citation`   | A passage and the 1-based page it is expected on, as the viewer's `search()` takes them. The page is a hint, not a filter: candidates are ordered by distance to it. |
| `pageIndex`  | Only elements of that page (0-based).                                                                                                                                |
| `kinds`      | Only elements of those kinds.                                                                                                                                        |
| `within`     | Only an element and its descendants (a table's cells, a group's members).                                                                                            |
| `maxResults` | Default 5.                                                                                                                                                           |

The passes, with the `reason` and `score` a candidate reports:

| Pass         | Score      | What matched                                                                                                                  |
| ------------ | ---------- | ----------------------------------------------------------------------------------------------------------------------------- |
| `exact`      | 1          | The engine's `findText()`: the characters as the document holds them.                                                         |
| `normalized` | 0.9        | The elements' text with whitespace runs folded to one space, case folded, curly quotes and dashes made ASCII, NFKC applied.   |
| `fuzzy`      | 0.85 – 0.5 | The viewer's citation matching (Fuse.js candidates, contiguous edit alignment) over the elements' text; lower for more edits. |
| `kind-only`  | 0.5        | No text in the query: the elements of `kinds` in reading order.                                                               |

A candidate carries `elementId` (the first element the match touches),
`pageIndex`, a `snippet` (the matched text with a little context, on one
line) and a `range` when the match is a part of the element's text, in the
form `replaceText` and `setTextStyle` take; a match that covers an element's
whole text has no range, so the operation takes the element as it is. The
folded and fuzzy passes run over the elements' text joined in reading order,
so a passage spanning two paragraphs or two PDF text objects resolves to a
range from the first element into the last. A Word table contributes its
cells, never its own text, so a match names the cell paragraph the
operations take.

Without text and without `kinds` the result is empty; with text that
nothing resembles it is empty as well, never an error.

## Checkpoints

```ts
const checkpoint = await session.createCheckpoint("before the agent's turn");
// … the agent applies a few batches …
if (!accepted) await session.restoreCheckpoint(checkpoint.id);
session.dropCheckpoint(checkpoint.id);
```

A checkpoint pins a state the host names, typically before an agent's turn,
so the whole turn can be reviewed and rejected at once. It is session state:
gone with the session, like the history.

| Method                            | What it does                                                                                                                                            |
| --------------------------------- | ------------------------------------------------------------------------------------------------------------------------------------------------------- |
| `createCheckpoint(label?)`        | Pins the current state; queued like a change, so it names the state after the calls before it. Rejects with `resource-limit` past `maxEditCheckpoints`. |
| `listCheckpoints()`               | Every checkpoint alive, in creation order: `{ id, label?, revision, createdAt }`.                                                                       |
| `restoreCheckpoint(id, options?)` | Back to the checkpoint's content as one history entry. Takes `expectedRevision` like `undo()`. Rejects with `invalid-operation` for an unknown id.      |
| `dropCheckpoint(id)`              | Forgets it; unknown ids are ignored.                                                                                                                    |

A restore is a change: it bumps the revision, emits `editstatechange` and
`documentchange` with reason `restore` and every page of the larger of the
two states, and returns a receipt whose `removedIds` and `createdIds` are
the net of the batches between the two states. `undo()` then returns to the
state before the restore, `redo()` to the checkpoint again; restoring the
state the session is in is a no-op. A checkpoint survives `undo()`, `redo()`
and `reset()`, and can be restored from a branch the history has dropped (a
new change after an undo). Restoring a checkpoint of a state that was saved
makes `dirty` false again, since the content is the same.

The session already retains the bytes of some states for fast undo (the
editing API's checkpoints, within `maxEditCheckpointBytes`). A named
checkpoint pins its state's bytes: they are never evicted while a checkpoint
names them, they count against the budget, and when the budget cannot hold
them the restore replays the checkpoint's batches from the original instead.
The content is the same either way.

## Tools

```ts
const session = await viewer.edit();
// Hand the definitions to the model provider: OpenAI `parameters`,
// Anthropic `input_schema`; the schemas are plain JSON Schema 2020-12.
const tools = session.tools.definitions.map((tool) => ({
  name: tool.name,
  description: tool.description,
  input_schema: tool.inputSchema,
}));
// Run what the model calls; `content` goes back to the model, `text` to the chat.
const result = await session.callTool(
  { name: call.name, arguments: call.input },
  { expectedRevision: revision, signal },
);
```

`session.tools` is the editing API as a model provider lists tools: the same
eight names on every format, the operations of the session's format in
`document_apply`. `callTool()` validates the call against the tool's schema,
runs it on the session and answers with JSON for the model and a sentence
for the chat. A model's mistake never throws: an unknown tool, bad
arguments, a refused batch or a stale revision come back as `ok: false` with
`issues` in the shape `apply()` reports and a `text` the model can act on.
Only the session's own errors throw (`lifecycle-error`, `aborted`).

| Tool                  | Arguments                                                  | What it does                                                                                  |
| --------------------- | ---------------------------------------------------------- | --------------------------------------------------------------------------------------------- |
| `document_describe`   | `OutlineOptions` + `maxChars`                              | `describe()`; the usual first call of a turn. `content` is the description with the revision. |
| `document_outline`    | `OutlineOptions`                                           | `getOutline()`, as JSON (`nodes`, `nodeCount`, `truncated`, `revision`).                      |
| `document_find`       | `TargetQuery`                                              | `resolveTargets()`; `content.candidates`, the best first.                                     |
| `document_inspect`    | `{ id }`                                                   | `getElement()` with text, styles and table rows, geometry left out.                           |
| `document_preview`    | `{ operations }`                                           | `apply()` with `dryRun: true`: the receipt an apply would give, or the issues.                |
| `document_apply`      | `{ operations, label? }`                                   | `apply()`; `expectedRevision`, `changeMode` and `author` come from `ToolCallOptions`.         |
| `document_undo`       | `{}`                                                       | `undo()`.                                                                                     |
| `document_checkpoint` | `{ action: "create" \| "restore" \| "list", id?, label? }` | The checkpoint methods; `restore` takes `expectedRevision`.                                   |

The `operations` of a batch validate against the session's operation
schemas as one `oneOf`, so a model gets the same `invalid-operation` issues
as any client, with `operationIndex` and `path`. Binary payloads travel as
asset references: the host registers the bytes with `addAsset()` and the
tool schema narrows `data` to `asset:<sha-256>`; a base64 payload in a tool
call is refused with `unknown-asset` and guidance in `text`. A preview is
not required before an apply (decision 9); `document_apply`'s description
recommends one for batches above one operation, and says which change mode
the format has.

`ToolCallOptions` carries what the host decides per turn: `expectedRevision`
(checked by every tool that changes the document), `changeMode` and `author`
for `document_apply` (see tracked changes), and a `signal`.

### Receipts in words

`describeReceipt(format, operations, receipt)` (exported) is what
`document_apply` and `document_preview` put in `text`: one sentence per
operation in the format's words, with the id each creating operation got
when the receipt names one per creation, then what else the receipt
reports.

```text
Added a text box on page 1 (p0:n1). Deleted p0:o3. Removed p0:o3. Revision 3; changed pages 1, 2; 2 pages.
Preview, nothing changed: Added a paragraph after p:1A000000 (p:2B000000). Would repaint 8 pages (1–8); 8 pages after.
Deleted slide 3. Moved slide 1 to 4. Removed sld3:2, sld3:3. Revision 3; changed slides 1, 2, 3, 4; 4 slides.
```

### An agent turn

1. `document_checkpoint { action: "create", label: "turn" }` so the whole
   turn can be rejected at once.
2. `document_describe {}`: the document with ids; keep `content.revision`.
3. `document_find { text: "…" }` for what the person quoted; take the best
   candidate's `elementId` and `range`.
4. `document_preview { operations }` for a batch above one operation, then
   `document_apply { operations }` with `expectedRevision` from step 2.
5. On `ok: false`, read `issues` (`operationIndex`, `path`, `code`) and
   `text`, fix the batch and send it again; on `edit-conflict`, describe
   again first.
6. To reject the turn: `document_checkpoint { action: "restore", id }`.

## Limits

| Limit                | Default | Meaning                                                  |
| -------------------- | ------- | -------------------------------------------------------- |
| `maxOutlineNodes`    | 5 000   | Nodes a `getOutline()` returns; more is cut and reported |
| `maxDescribeChars`   | 200 000 | Upper bound of `DescribeOptions.maxChars`                |
| `maxEditCheckpoints` | 20      | Named checkpoints alive at once                          |

All three join `ResourceLimits` and are raised through `ViewerClient.create({ limits })`.
