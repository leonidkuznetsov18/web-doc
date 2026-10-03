import type { ResourceLimits, ViewerWarning } from "../../contracts.js";
import { ViewerError } from "../../errors.js";
import { AssetStore } from "../assets.js";
import type {
  BatchMode,
  EditEngine,
  EngineBatch,
  EngineChange,
  MaterializedDocument,
  MaterializeOptions,
  RestoreTarget,
} from "../engine.js";
import { OoxmlPackage } from "../ooxml/package.js";
import { patches, type XmlPatch } from "../ooxml/patch.js";
import {
  checkOperations,
  freezeOperations,
  invalidOperationError,
  parseReference,
} from "../operations.js";
import type {
  EditFindOptions,
  EditOperation,
  ElementQuery,
  OperationIssue,
  PagePoint,
  TextTarget,
} from "../types.js";
import { NO_PAGE, toElement } from "./elements.js";
import { docxHandlers } from "./handlers.js";
import { freshParagraphId, paragraphsOf } from "./ids.js";
import { DocxModel, type AnyRecord } from "./model.js";
import { resolveTextStyle, type DocxStyles } from "./style.js";
import {
  issueCollector,
  type DocxOperationContext,
  type TrackedChange,
} from "./operations.js";
import { docxOperationSchemas } from "./schemas.js";
import { revisionsOf } from "./tracked.js";
import type {
  DocxElement,
  DocxFields,
  DocxReplaceTextOperation,
  DocxOperation,
  DocxRevision,
  DocxTextStyle,
} from "./types.js";
import {
  runsCovering,
  sharedStyle,
  spanFits,
  type TextSpan,
} from "../range-style.js";
import { attributeProblem, namespacePatches } from "./write.js";

/*
 * The DOCX edit engine: the package layer under a block index of the body
 * story. It runs inside the OOXML edit worker in the browser and directly
 * in Node tests. It never lays out: elements carry no geometry, and the
 * session joins the renderer's runs on the main thread.
 *
 * Paragraph ids: a paragraph with a `w14:paraId` keeps it; one without is
 * numbered from its position when the document opens, and the engine then
 * tracks those ids by document order (`#unauthored`), writing a
 * `w14:paraId` on every paragraph it rebuilds or creates. The shown copy
 * carries an id on every paragraph, so the viewer's runs name the same
 * paragraphs; the saved file carries ids only where the session wrote.
 */

/**
 * The part a shown copy carries with the ids the engine owns, so a copy
 * restored as a base knows which `w14:paraId` values to leave out of a
 * saved file. An XML part: its content type is the one every package
 * declares for the extension, so `[Content_Types].xml` stays untouched.
 */
const UNAUTHORED_PART = "/webdoc/unauthored.xml";
const UNAUTHORED_NS = "urn:web-doc:docx-edit";

function unauthoredPartXml(ids: Iterable<string>): Uint8Array {
  const items = [...ids].map((id) => `<p id="${id}"/>`).join("");
  return new TextEncoder().encode(
    `<?xml version="1.0" encoding="UTF-8" standalone="yes"?><unauthored xmlns="${UNAUTHORED_NS}">${items}</unauthored>`,
  );
}

function unauthoredIdsOf(bytes: Uint8Array): string[] {
  const text = new TextDecoder().decode(bytes);
  return [...text.matchAll(/<p id="([0-9A-F]{8})"\/>/g)].map((m) => m[1]!);
}

/** Reads the DOCX session adds on top of the core, served by the engine and the worker client alike. */
export interface DocxEngineReads {
  previewDraft(
    fields: DocxFields<DocxReplaceTextOperation>,
    signal: AbortSignal,
  ): Promise<DocxDraftDocument>;
  previewText(
    fields: DocxFields<DocxReplaceTextOperation>,
    signal: AbortSignal,
  ): Promise<Uint8Array>;
  revisions(id: string, signal: AbortSignal): Promise<readonly DocxRevision[]>;
  textStyle(
    id: string,
    span: TextSpan | undefined,
    signal: AbortSignal,
  ): Promise<Partial<DocxTextStyle> | undefined>;
  textColors(
    id: string,
    spans: readonly TextSpan[],
    signal: AbortSignal,
  ): Promise<readonly string[] | undefined>;
}

/** Internal worker read: display bytes and target metadata belong to the same isolated draft. */
export interface DocxDraftDocument {
  readonly bytes: Uint8Array;
  readonly paragraph?: DocxElement;
}

/** Word's automatic text colour on a white page. */
const AUTOMATIC_COLOR = "#000000";

/** The tracked-change record of a batch, when it writes revisions. */
function trackedOf(mode: BatchMode): TrackedChange | undefined {
  if (mode.changeMode !== "tracked") return undefined;
  return {
    author: mode.author ?? "",
    ...(mode.timestamp === undefined ? {} : { date: mode.timestamp }),
  };
}

/** ISO 8601 as `xsd:dateTime` takes it; what `w:date` carries. */
const DATE_TIME =
  /^\d{4}-\d{2}-\d{2}T\d{2}:\d{2}:\d{2}(?:\.\d+)?(?:Z|[+-]\d{2}:\d{2})?$/;

/**
 * What a tracked batch must carry before any revision is written: an author
 * the file can hold (the core requires one too; the engine used on its own
 * does the same) and a date the file can hold.
 */
function trackedModeIssues(
  tracked: TrackedChange | undefined,
): OperationIssue[] {
  if (!tracked) return [];
  const issues: OperationIssue[] = [];
  const author = attributeProblem(tracked.author);
  if (tracked.author.trim().length === 0)
    issues.push({
      operationIndex: -1,
      path: "/author",
      code: "required",
      message: "Tracked changes name their author; pass ApplyOptions.author",
    });
  else if (author)
    issues.push({
      operationIndex: -1,
      path: "/author",
      code: "invalid-value",
      message: `The author holds ${author}`,
    });
  if (tracked.date !== undefined) {
    const date = attributeProblem(tracked.date);
    if (
      date ||
      !DATE_TIME.test(tracked.date) ||
      Number.isNaN(Date.parse(tracked.date))
    )
      issues.push({
        operationIndex: -1,
        path: "/timestamp",
        code: "invalid-value",
        message: date
          ? `The timestamp holds ${date}`
          : "The timestamp must be an ISO 8601 date-time",
      });
  }
  return issues;
}

export class DocxEditEngine implements EditEngine, DocxEngineReads {
  readonly schemas = docxOperationSchemas;
  readonly #original: Uint8Array;
  readonly #limits: ResourceLimits;
  readonly #assets = new AssetStore();
  #pkg: OoxmlPackage;
  #model: Promise<DocxModel> | undefined;
  /** The styles and theme, read once: no operation changes them. */
  #styles: DocxStyles | undefined;
  /** Ids of the paragraphs without `w14:paraId`, in document order; undefined until read from the bytes. */
  #unauthored: string[] | undefined;
  #disposed = false;
  #stateId = 0;

  private constructor(
    original: Uint8Array,
    pkg: OoxmlPackage,
    limits: ResourceLimits,
  ) {
    this.#original = original;
    this.#pkg = pkg;
    this.#limits = limits;
  }

  /** Opens the package and reads the block index, so a broken document fails here. */
  static async open(
    bytes: Uint8Array,
    limits: ResourceLimits,
    signal?: AbortSignal,
  ): Promise<DocxEditEngine> {
    const pkg = await OoxmlPackage.open(bytes, {
      limits,
      ...(signal ? { signal } : {}),
    });
    const engine = new DocxEditEngine(bytes, pkg, limits);
    await engine.model(signal);
    return engine;
  }

  /** The package behind the engine, for tests and the operations. */
  get package(): OoxmlPackage {
    return this.#pkg;
  }

  /** The engine has no pages of its own; the renderer counts them. */
  get pageCount(): number {
    return 0;
  }

  /** The block index at the current revision. */
  model(signal?: AbortSignal): Promise<DocxModel> {
    this.#assertAlive();
    const revision = this.#pkg.revision;
    if (!this.#model) return this.#startModel(signal);
    return this.#model.then((current) =>
      current.revision === revision ? current : this.#startModel(signal),
    );
  }

  #startModel(signal?: AbortSignal): Promise<DocxModel> {
    const pending = DocxModel.load(
      this.#pkg,
      signal,
      this.#unauthored,
      this.#styles,
    ).then((model) => {
      this.#unauthored ??= [...model.unauthoredIds];
      this.#styles ??= model.styles;
      return model;
    });
    this.#model = pending;
    pending.catch(() => {
      if (this.#model === pending) this.#model = undefined;
    });
    return pending;
  }

  async validate(
    operations: readonly EditOperation[],
    signal: AbortSignal,
    mode: BatchMode = {},
  ): Promise<readonly OperationIssue[]> {
    const issues: OperationIssue[] = [];
    const tracked = trackedOf(mode);
    const modeIssues = trackedModeIssues(tracked);
    if (modeIssues.length > 0) return modeIssues;
    const base = await this.#context(0, signal, 0, new Set(), tracked);
    for (const [index, operation] of operations.entries()) {
      const context = { ...base, operationIndex: index };
      const handler = docxHandlers.get(operation.op);
      if (!handler) {
        issues.push({
          operationIndex: index,
          path: "/op",
          code: "unknown-operation",
          message: `Unknown operation ${operation.op}`,
        });
        continue;
      }
      // A same-batch reference names an element that does not exist yet:
      // its target is checked when the batch is applied, the rest now.
      const referenced = referenceFields(operation);
      const forward = referenced.find((entry) => entry.reference >= index);
      if (forward) {
        issues.push({
          operationIndex: index,
          path: `/${forward.field}`,
          code: "unknown-target",
          message: `"${forward.value}" must refer to an earlier operation`,
        });
        continue;
      }
      const collect = issueCollector(index, issues);
      const skipped = referenced.map((entry) => `/${entry.field}`);
      await handler.validate(
        operation as DocxOperation,
        context,
        referenced.length === 0
          ? collect
          : (path, code, message) => {
              if (
                !skipped.some(
                  (prefix) => path === prefix || path.startsWith(`${prefix}/`),
                ) &&
                !path.startsWith("/range")
              )
                collect(path, code, message);
            },
      );
    }
    return issues;
  }

  /** Plain operation arrays, as the unit tests pass them, become the next batch. */
  async apply(
    input: EngineBatch | readonly EditOperation[],
    signal: AbortSignal,
  ): Promise<EngineChange> {
    const batch: EngineBatch = Array.isArray(input)
      ? { stateId: this.#nextStateId(), operations: input }
      : (input as EngineBatch);
    this.#stateId = Math.max(this.#stateId, batch.stateId);
    const tracked = trackedOf(batch);
    const modeIssues = trackedModeIssues(tracked);
    if (modeIssues.length > 0) throw invalidOperationError(modeIssues);
    const snapshot = this.#pkg.snapshot();
    const unauthored = this.#unauthored ? [...this.#unauthored] : undefined;
    const createdIds: string[] = [];
    const removedIds: string[] = [];
    const warnings: ViewerWarning[] = [];
    const createdByOperation: string[][] = [];
    const issued = new Set<string>();
    const remappedIds: Record<string, string> = {};
    let reflowFrom: string | undefined;
    try {
      for (const [index, raw] of batch.operations.entries()) {
        throwIfAborted(signal);
        const operation = resolveReferences(raw, createdByOperation, index);
        const handler = docxHandlers.get(operation.op);
        if (!handler)
          throw new ViewerError(
            "invalid-operation",
            `Unknown operation ${operation.op}`,
          );
        const context = await this.#context(
          batch.stateId,
          signal,
          index,
          issued,
          tracked,
        );
        const issues: OperationIssue[] = [];
        await handler.validate(
          operation,
          context,
          issueCollector(index, issues),
        );
        if (issues.length > 0) throw invalidOperationError(issues);
        const result = await handler.apply(operation, context);
        this.#model = undefined;
        const gone = new Set([
          ...(result.stamped ?? []),
          ...(result.removedParagraphIds ?? []),
        ]);
        if (gone.size > 0 && this.#unauthored)
          this.#unauthored = this.#unauthored.filter((id) => !gone.has(id));
        createdByOperation.push([...result.createdIds]);
        createdIds.push(...result.createdIds);
        removedIds.push(...(result.removedIds ?? []));
        warnings.push(...result.warnings);
        reflowFrom ??= result.reflowFrom;
        // A table renamed twice in one batch maps its first id to its last.
        for (const [from, to] of Object.entries(result.remappedIds ?? {})) {
          const origin =
            Object.entries(remappedIds).find(
              ([, value]) => value === from,
            )?.[0] ?? from;
          remappedIds[origin] = to;
        }
      }
      // The index after the batch; an inconsistency surfaces here and
      // rolls the batch back like any other failure.
      await this.model(signal);
    } catch (error) {
      // Everything the batch did, ids included, is undone.
      this.#pkg.restore(snapshot);
      this.#unauthored = unauthored;
      this.#model = undefined;
      throw error;
    }
    this.#pkg.release(snapshot);
    return {
      createdIds,
      removedIds,
      // A flow document reflows from the first changed paragraph; the host
      // turns it into pages.
      changedPages: [],
      ...(reflowFrom === undefined ? {} : { reflowFrom }),
      ...(Object.keys(remappedIds).length > 0 ? { remappedIds } : {}),
      warnings,
    };
  }

  #nextStateId(): number {
    this.#stateId += 1;
    return this.#stateId;
  }

  async #context(
    stateId: number,
    signal: AbortSignal | undefined,
    operationIndex = 0,
    issued: Set<string> = new Set(),
    tracked?: TrackedChange,
  ): Promise<DocxOperationContext> {
    const model = await this.model(signal);
    const taken = new Set([...model.takenIds, ...issued]);
    let count = 0;
    let revisions = 0;
    return {
      pkg: this.#pkg,
      model,
      limits: this.#limits,
      assets: this.#assets,
      stateId,
      operationIndex,
      freshParagraphId: () => {
        const id = freshParagraphId(stateId, operationIndex, count, taken);
        count += 1;
        issued.add(id);
        taken.add(id);
        return id;
      },
      ...(tracked ? { tracked } : {}),
      nextRevisionId: () => {
        revisions += 1;
        return model.maxRevisionId + revisions;
      },
    };
  }

  /** The revisions of a paragraph; none for another element or an unknown id. */
  async revisions(
    id: string,
    signal: AbortSignal,
  ): Promise<readonly DocxRevision[]> {
    const model = await this.model(signal);
    const record = model.byId.get(id);
    if (!record || record.kind !== "paragraph") return [];
    return revisionsOf(model.document, record.node);
  }

  /** Display-only bytes of an isolated draft; the live package is never written. */
  async previewText(
    fields: DocxFields<DocxReplaceTextOperation>,
    signal: AbortSignal,
  ): Promise<Uint8Array> {
    return (await this.previewDraft(fields, signal)).bytes;
  }

  async previewDraft(
    fields: DocxFields<DocxReplaceTextOperation>,
    signal: AbortSignal,
  ): Promise<DocxDraftDocument> {
    const input = [{ ...fields, op: "replaceText" as const }];
    const issues = checkOperations(input, this.schemas);
    if (issues.length > 0) throw invalidOperationError(issues);
    const operations = freezeOperations(input);
    // A separate package keeps draft reads out of live ids, history and save bytes.
    const shown = await this.materializeDocument("show", {}, signal);
    const preview = await DocxEditEngine.open(
      shown.bytes,
      this.#limits,
      signal,
    );
    try {
      await preview.apply(operations, signal);
      const bytes = (await preview.materializeDocument("show", {}, signal))
        .bytes;
      const paragraph = await preview.getElement(fields.target, signal);
      return { bytes, ...(paragraph ? { paragraph } : {}) };
    } finally {
      await preview.dispose();
    }
  }

  /**
   * The style a span of a paragraph shows: what every run it covers shares;
   * the paragraph mark's for an empty paragraph. None for another element,
   * an unknown id or a span past the text.
   */
  async textStyle(
    id: string,
    span: TextSpan | undefined,
    signal: AbortSignal,
  ): Promise<Partial<DocxTextStyle> | undefined> {
    const model = await this.model(signal);
    const record = model.byId.get(id);
    if (!record || record.kind !== "paragraph") return undefined;
    const { text, items } = record.text;
    const range = span ?? { start: 0, end: text.length };
    if (!spanFits(range, text.length)) return undefined;
    const covered = runsCovering(items, range);
    const mark = record.pPr?.children.find((child) => child.local === "rPr");
    return sharedStyle(
      covered.length > 0
        ? covered.map((item) =>
            resolveTextStyle(model.styles, record.pPr, item.rPr),
          )
        : [resolveTextStyle(model.styles, record.pPr, mark)],
    );
  }

  /**
   * The `#RRGGBB` each span of a paragraph is drawn in: its first run's
   * colour, a theme colour resolved through the theme and Word's automatic
   * colour as black. None for another element or an unknown id.
   */
  async textColors(
    id: string,
    spans: readonly TextSpan[],
    signal: AbortSignal,
  ): Promise<readonly string[] | undefined> {
    const model = await this.model(signal);
    const record = model.byId.get(id);
    if (!record || record.kind !== "paragraph") return undefined;
    const { text, items } = record.text;
    const mark = record.pPr?.children.find((child) => child.local === "rPr");
    return spans.map((span) => {
      const first = spanFits(span, text.length)
        ? runsCovering(items, span)[0]
        : undefined;
      const { color } = resolveTextStyle(
        model.styles,
        record.pPr,
        first ? first.rPr : mark,
      );
      if (typeof color !== "string") {
        const value = model.styles.themeColor(color.theme);
        return value ? `#${value}` : AUTOMATIC_COLOR;
      }
      return color === "auto" ? AUTOMATIC_COLOR : color.slice(0, 7);
    });
  }

  async materialize(
    purposeOrSignal: "show" | "save" | AbortSignal = "show",
    options: MaterializeOptions = {},
    signal: AbortSignal = new AbortController().signal,
  ): Promise<Uint8Array> {
    return (await this.materializeDocument(purposeOrSignal, options, signal))
      .bytes;
  }

  /**
   * `save` is the package as the session changed it: ids written only on
   * the paragraphs the session rebuilt or created. `show` also stamps every
   * other paragraph with the id the engine knows it by, in a copy, so the
   * display pre-pass and the viewer's runs name the engine's paragraphs
   * after edits that moved paragraphs around. Without changes both are
   * the original bytes.
   */
  async materializeDocument(
    purposeOrSignal: "show" | "save" | AbortSignal = "show",
    _options: MaterializeOptions = {},
    signal: AbortSignal = new AbortController().signal,
  ): Promise<MaterializedDocument> {
    this.#assertAlive();
    const purpose =
      purposeOrSignal instanceof AbortSignal ? "show" : purposeOrSignal;
    const own =
      purposeOrSignal instanceof AbortSignal ? purposeOrSignal : signal;
    const bytes = await this.#pkg.save({}, own);
    const fromShownBase = this.#pkg.has(UNAUTHORED_PART);
    if (purpose === "save")
      return {
        bytes: fromShownBase ? await this.#unstamped(bytes, own) : bytes,
        warnings: [],
      };
    if (this.#pkg.changedParts.length === 0 && !fromShownBase)
      return { bytes, warnings: [] };
    return { bytes: await this.#stamped(bytes, own), warnings: [] };
  }

  /**
   * A copy of `bytes` with a `w14:paraId` on every paragraph of the main
   * part and the engine's own ids listed in a part of their own.
   */
  async #stamped(bytes: Uint8Array, signal: AbortSignal): Promise<Uint8Array> {
    const model = await this.model(signal);
    const copy = await OoxmlPackage.open(bytes, {
      limits: this.#limits,
      signal,
    });
    const part = await copy.xml(model.mainPart, signal);
    const owned = this.#unauthored ?? [];
    // Ids already written (by an edit, or by the shown base this state
    // came from) stay; the rest go on the unmarked paragraphs in order.
    const attributed = new Set<string>();
    for (const paragraph of paragraphsOf(part)) {
      const id = part.attribute(paragraph, "w14:paraId");
      if (id) attributed.add(id.toUpperCase());
    }
    const queue = owned.filter((id) => !attributed.has(id));
    const items: XmlPatch[] = [];
    for (const paragraph of paragraphsOf(part)) {
      if (part.attribute(paragraph, "w14:paraId")) continue;
      const id = queue.shift();
      if (id === undefined)
        throw new ViewerError(
          "internal",
          "The document has more unmarked paragraphs than the session knows",
        );
      items.push(patches.setAttribute(part, paragraph, "w14:paraId", id));
    }
    const list = owned.length > 0 ? unauthoredPartXml(owned) : undefined;
    if (items.length === 0) {
      // Already the shown form of this state: nothing to write.
      const current = copy.has(UNAUTHORED_PART)
        ? await copy.part(UNAUTHORED_PART, signal)
        : undefined;
      if (
        current === undefined
          ? list === undefined
          : list !== undefined && sameBytes(current, list)
      )
        return bytes;
    }
    const transaction = copy.transaction();
    if (items.length > 0)
      transaction.patch(part, [...namespacePatches(part), ...items]);
    if (list) transaction.setPart(UNAUTHORED_PART, list, "application/xml");
    else if (copy.has(UNAUTHORED_PART)) transaction.removePart(UNAUTHORED_PART);
    await transaction.commit(signal);
    return copy.save({}, signal);
  }

  /**
   * A copy of `bytes` without the ids a shown base brought: the engine's
   * own `w14:paraId` values go, the ones an edit wrote stay, and the list
   * part goes with them.
   */
  async #unstamped(
    bytes: Uint8Array,
    signal: AbortSignal,
  ): Promise<Uint8Array> {
    const model = await this.model(signal);
    const copy = await OoxmlPackage.open(bytes, {
      limits: this.#limits,
      signal,
    });
    const part = await copy.xml(model.mainPart, signal);
    const owned = model.unauthoredSet;
    const items: XmlPatch[] = [];
    for (const paragraph of paragraphsOf(part)) {
      const id = part.attribute(paragraph, "w14:paraId")?.toUpperCase();
      if (id && owned.has(id))
        items.push(patches.removeAttribute(part, paragraph, "w14:paraId"));
    }
    const transaction = copy.transaction();
    if (items.length > 0) transaction.patch(part, items);
    if (copy.has(UNAUTHORED_PART)) transaction.removePart(UNAUTHORED_PART);
    await transaction.commit(signal);
    return copy.save({}, signal);
  }

  async restore(
    input: RestoreTarget | readonly (readonly EditOperation[])[],
    signal: AbortSignal,
  ): Promise<void> {
    const target: RestoreTarget = Array.isArray(input)
      ? {
          batches: (input as readonly (readonly EditOperation[])[]).map(
            (operations, index) => ({ stateId: index + 1, operations }),
          ),
        }
      : (input as RestoreTarget);
    this.#assertAlive();
    this.#pkg = await OoxmlPackage.open(target.base ?? this.#original, {
      limits: this.#limits,
      signal,
    });
    // The original has the ids its positions give; a base is a shown copy
    // with every paragraph marked and the engine's own ids listed.
    this.#unauthored = undefined;
    if (this.#pkg.has(UNAUTHORED_PART))
      this.#unauthored = unauthoredIdsOf(
        await this.#pkg.part(UNAUTHORED_PART, signal),
      );
    this.#model = undefined;
    await this.model(signal);
    for (const batch of target.batches) await this.apply(batch, signal);
  }

  async putAsset(
    id: string,
    data: Uint8Array,
    _signal: AbortSignal,
  ): Promise<void> {
    this.#assets.set(id, data);
  }

  /**
   * Every element of the body story, in document order, without geometry.
   * `pageIndex` and `intersects` are the session's to apply once the runs
   * are joined; `kinds` filters here.
   */
  async getElements(
    query: ElementQuery,
    signal: AbortSignal,
  ): Promise<readonly DocxElement[]> {
    const model = await this.model(signal);
    const out: DocxElement[] = [];
    for (const record of model.records) {
      if (query.kinds && !query.kinds.includes(record.kind)) continue;
      out.push(toElement(model, record));
    }
    return out;
  }

  async getElement(
    id: string,
    signal: AbortSignal,
  ): Promise<DocxElement | undefined> {
    const record = await this.locate(id, signal);
    return record ? toElement(await this.model(signal), record) : undefined;
  }

  /** The record an element id names, at the current revision. */
  async locate(
    id: string,
    signal?: AbortSignal,
  ): Promise<AnyRecord | undefined> {
    return (await this.model(signal)).byId.get(id);
  }

  /** Hit-testing needs the renderer's geometry; the session answers it. */
  async elementsAt(
    _pageIndex: number,
    _point: PagePoint,
    _signal: AbortSignal,
  ): Promise<readonly DocxElement[]> {
    return [];
  }

  /**
   * Matches in paragraph text, in document order, without rectangles or
   * pages: the session adds those from the renderer's runs.
   */
  async findText(
    query: string,
    options: EditFindOptions,
    signal: AbortSignal,
  ): Promise<readonly TextTarget[]> {
    if (query.length === 0) return [];
    const model = await this.model(signal);
    // A case-insensitive regular expression keeps offsets on the original
    // text, where a lower-cased copy can change length.
    const needle = new RegExp(
      query.replace(/[.*+?^${}()|[\]\\]/g, "\\$&"),
      options.caseSensitive ? "g" : "gi",
    );
    const limit = options.maxResults ?? Number.POSITIVE_INFINITY;
    const targets: TextTarget[] = [];
    for (const record of model.records) {
      if (record.kind !== "paragraph") continue;
      const text = record.text.text;
      if (!text) continue;
      needle.lastIndex = 0;
      while (targets.length < limit) {
        const match = needle.exec(text);
        if (!match) break;
        const at = match.index;
        const end = at + match[0].length;
        if (match[0].length === 0) needle.lastIndex += 1;
        targets.push({
          pageIndex: NO_PAGE,
          text: text.slice(at, end),
          rects: [],
          elementIds: [record.elementId],
          ranges: [
            {
              start: { elementId: record.elementId, offset: at },
              end: { elementId: record.elementId, offset: end },
            },
          ],
        });
      }
      if (targets.length >= limit) break;
    }
    return targets;
  }

  async dispose(): Promise<void> {
    this.#disposed = true;
    this.#model = undefined;
  }

  #assertAlive(): void {
    if (this.#disposed)
      throw new ViewerError("lifecycle-error", "The DOCX engine was disposed");
  }
}

const REFERENCE_FIELDS = ["target", "before", "after"] as const;

function referenceFields(
  operation: EditOperation,
): { field: string; value: string; reference: number }[] {
  const out: { field: string; value: string; reference: number }[] = [];
  for (const field of REFERENCE_FIELDS) {
    const value = (operation as unknown as Record<string, unknown>)[field];
    if (typeof value !== "string") continue;
    const reference = parseReference(value);
    if (reference !== undefined) out.push({ field, value, reference });
  }
  return out;
}

/** Replaces `"$<n>"` references with the first id operation `n` created. */
function resolveReferences(
  operation: EditOperation,
  created: readonly (readonly string[])[],
  index: number,
): DocxOperation {
  let resolved = operation as unknown as Record<string, unknown>;
  for (const { field, value, reference } of referenceFields(operation)) {
    const id = reference < index ? created[reference]?.[0] : undefined;
    if (!id)
      throw new ViewerError(
        "invalid-operation",
        `Operation ${index} refers to "${value}", which created nothing`,
        { details: { operationIndex: index, target: value } },
      );
    resolved = { ...resolved, [field]: id };
  }
  return resolved as unknown as DocxOperation;
}

function sameBytes(a: Uint8Array, b: Uint8Array): boolean {
  return a.length === b.length && a.every((byte, index) => byte === b[index]);
}

function throwIfAborted(signal: AbortSignal | undefined): void {
  if (signal?.aborted)
    throw new ViewerError("aborted", "The operation was aborted");
}
