import type { ResourceLimits } from "../../contracts.js";
import { ViewerError } from "../../errors.js";
import { isDirectoryEntry, partKey, partNameOf } from "./names.js";
import { writeZip, type WriteOverlay } from "./writer.js";
import { decodePart, scanXml, type XmlPart } from "./xml.js";
import {
  parseContentTypes,
  parseRelationships,
  type ContentTypes,
  type RelationshipSet,
} from "./opc.js";
import { PackageTransaction } from "./transaction.js";
import { relationshipsPartOf, resolveTarget } from "./names.js";
import {
  inflateEntry,
  parseZip,
  type ZipArchive,
  type ZipEntry,
} from "./zip.js";

/*
 * An OOXML package held in memory: the original bytes, the ZIP directory,
 * and a cache of the parts that were read. Parts are inflated on first use.
 * Changes live in an overlay — replaced, added and removed parts — that
 * saving applies over the archive; the original is never touched.
 */

export interface SaveOptions {
  /** How changed and new entries are written; default "store". */
  readonly compression?: "store" | "deflate";
}

/** The overlay at a point in time; opaque to callers, O(1) to take and restore. */
export interface PackageSnapshot {
  readonly revision: number;
}

interface Overlay extends WriteOverlay {
  readonly revision: number;
}

export interface OpenOptions {
  readonly limits: ResourceLimits;
  readonly signal?: AbortSignal;
}

export const CONTENT_TYPES_PART = "/[Content_Types].xml";

export class OoxmlPackage {
  /** The original bytes, never mutated. */
  readonly original: Uint8Array;
  /** Absolute part names in archive order; directory markers are left out. */
  readonly partNames: readonly string[];
  readonly #archive: ZipArchive;
  readonly #limits: ResourceLimits;
  readonly #entries = new Map<string, ZipEntry>();
  readonly #parts = new Map<string, Promise<Uint8Array>>();
  #overlay: Overlay = {
    revision: 0,
    changed: new Map(),
    added: new Map(),
    removed: new Set(),
  };
  readonly #snapshots = new Map<number, Overlay>();
  /** Scanned parts by key, valid for the overlay revision they were scanned at. */
  readonly #xml = new Map<string, Promise<XmlPart>>();

  private constructor(archive: ZipArchive, limits: ResourceLimits) {
    this.original = archive.bytes;
    this.#archive = archive;
    this.#limits = limits;
    const names: string[] = [];
    for (const entry of archive.entries) {
      if (isDirectoryEntry(entry.name)) continue;
      const name = partNameOf(entry.name);
      const key = partKey(name);
      if (this.#entries.has(key))
        throw new ViewerError("invalid-file", `Duplicate package part ${name}`);
      this.#entries.set(key, entry);
      names.push(name);
    }
    this.partNames = Object.freeze(names);
    if (!this.has(CONTENT_TYPES_PART))
      throw new ViewerError(
        "invalid-file",
        "The package has no [Content_Types].xml",
      );
  }

  /** Parses the central directory; parts are inflated on first use. */
  static async open(
    bytes: Uint8Array,
    options: OpenOptions,
  ): Promise<OoxmlPackage> {
    if (options.signal?.aborted)
      throw new ViewerError("aborted", "Opening the package was aborted");
    return new OoxmlPackage(parseZip(bytes, options.limits), options.limits);
  }

  /** The archive behind the package, for the writer. */
  get archive(): ZipArchive {
    return this.#archive;
  }

  get limits(): ResourceLimits {
    return this.#limits;
  }

  /** Whether a part exists in the current state, overlay included. */
  has(name: string): boolean {
    const key = partKey(name);
    if (this.#overlay.removed.has(key)) return false;
    return this.#entries.has(key) || this.#overlay.added.has(key);
  }

  /** Part names in the current state: archive order, then additions. */
  get currentPartNames(): readonly string[] {
    const names = this.partNames.filter((name) => this.has(name));
    for (const [, change] of this.#overlay.added) names.push(change.name);
    return names;
  }

  /** Names of the parts that differ from the original: changed, added and removed. */
  get changedParts(): readonly string[] {
    return [
      ...[...this.#overlay.changed.values()].map((change) => change.name),
      ...[...this.#overlay.added.values()].map((change) => change.name),
      ...[...this.#overlay.removed]
        .map((key) => this.#entries.get(key)?.name ?? key)
        .map(partNameOf),
    ];
  }

  get revision(): number {
    return this.#overlay.revision;
  }

  /** From [Content_Types].xml: an Override, else the Default for the extension. */
  async contentTypeOf(name: string): Promise<string | undefined> {
    return (await this.contentTypes()).typeOf(name);
  }

  /** The content-type model of the current state. */
  async contentTypes(): Promise<ContentTypes> {
    return parseContentTypes(await this.xml(CONTENT_TYPES_PART));
  }

  /** The relationships of a part, or of the package for "/". */
  async relationships(
    name: string,
    signal?: AbortSignal,
  ): Promise<RelationshipSet> {
    const source = name === "/" || name === "" ? "/" : partNameOf(name);
    const relsName = relationshipsPartOf(source);
    if (!this.has(relsName))
      return parseRelationships(source, undefined, undefined);
    return parseRelationships(
      source,
      relsName,
      await this.xml(relsName, signal),
    );
  }

  /** Resolves a relationship target against its source part to an absolute name. */
  resolve(sourcePart: string, target: string): string {
    return resolveTarget(sourcePart, target);
  }

  transaction(): PackageTransaction {
    return new PackageTransaction(this);
  }

  /** The ZIP entry behind a part. */
  entry(name: string): ZipEntry | undefined {
    return this.#entries.get(partKey(name));
  }

  /** The current bytes of a part: the overlay's when changed, else the original's. */
  part(name: string, signal?: AbortSignal): Promise<Uint8Array> {
    const key = partKey(name);
    if (this.#overlay.removed.has(key))
      return Promise.reject(this.#missing(name));
    const change =
      this.#overlay.changed.get(key) ?? this.#overlay.added.get(key);
    if (change) return Promise.resolve(change.bytes.slice());
    return this.originalPart(name, signal);
  }

  /**
   * A part decoded and scanned, cached until the part changes. A part that
   * is not UTF-8 is `unsupported-part`; one that does not parse is
   * `malformed-xml`.
   */
  xml(name: string, signal?: AbortSignal): Promise<XmlPart> {
    const key = partKey(name);
    const revision = this.#overlay.revision;
    const cached = this.#xml.get(key);
    if (cached)
      return cached.then((part) =>
        part.revision === revision
          ? part
          : this.#rescan(name, key, revision, signal),
      );
    return this.#rescan(name, key, revision, signal);
  }

  #rescan(
    name: string,
    key: string,
    revision: number,
    signal: AbortSignal | undefined,
  ): Promise<XmlPart> {
    const pending = this.part(name, signal).then((bytes) => {
      const partName = partNameOf(name);
      const { text } = decodePart(partName, bytes);
      return scanXml(partName, text, revision);
    });
    this.#xml.set(key, pending);
    pending.catch(() => this.#xml.delete(key));
    return pending;
  }

  /** The original bytes of a part, inflated once and cached; ignores the overlay. */
  originalPart(name: string, signal?: AbortSignal): Promise<Uint8Array> {
    const key = partKey(name);
    const entry = this.#entries.get(key);
    if (!entry) return Promise.reject(this.#missing(name));
    let pending = this.#parts.get(key);
    if (!pending) {
      pending = inflateEntry(this.#archive, entry, this.#limits, signal);
      this.#parts.set(key, pending);
      pending.catch(() => this.#parts.delete(key));
    }
    return pending.then((bytes) => bytes.slice());
  }

  /**
   * Replaces the overlay with one that has `changes` applied: replaced parts
   * (existing names) and added parts (new names) in `set`, and `remove`.
   * Used by transactions; the previous overlay stays reachable through its
   * snapshot.
   */
  applyOverlay(changes: {
    readonly set: ReadonlyMap<string, Uint8Array>;
    readonly remove: ReadonlySet<string>;
  }): PackageSnapshot {
    const changed = new Map(this.#overlay.changed);
    const added = new Map(this.#overlay.added);
    const removed = new Set(this.#overlay.removed);
    for (const name of changes.remove) {
      const key = partKey(name);
      changed.delete(key);
      added.delete(key);
      if (this.#entries.has(key)) removed.add(key);
    }
    for (const [name, bytes] of changes.set) {
      const key = partKey(name);
      removed.delete(key);
      const change = { name: partNameOf(name), bytes: bytes.slice() };
      if (this.#entries.has(key)) changed.set(key, change);
      else added.set(key, change);
    }
    const revision = this.#overlay.revision + 1;
    this.#overlay = Object.freeze({ revision, changed, added, removed });
    this.#snapshots.set(revision, this.#overlay);
    return { revision };
  }

  /** The overlay as it is now, by reference; restore it later in O(1). */
  snapshot(): PackageSnapshot {
    this.#snapshots.set(this.#overlay.revision, this.#overlay);
    return { revision: this.#overlay.revision };
  }

  restore(snapshot: PackageSnapshot): void {
    const overlay = this.#snapshots.get(snapshot.revision);
    if (!overlay)
      throw new ViewerError("lifecycle-error", "Unknown package snapshot", {
        details: { revision: snapshot.revision },
      });
    this.#overlay = overlay;
  }

  /** The package bytes: the original (copied) when nothing changed. */
  async save(
    options: SaveOptions = {},
    signal?: AbortSignal,
  ): Promise<Uint8Array> {
    const { changed, added, removed } = this.#overlay;
    if (changed.size === 0 && added.size === 0 && removed.size === 0)
      return this.original.slice();
    return writeZip(this.#archive, this.#overlay, {
      ...(options.compression ? { compression: options.compression } : {}),
      ...(signal ? { signal } : {}),
    });
  }

  #missing(name: string): ViewerError {
    return new ViewerError("invalid-file", `No part ${partNameOf(name)}`, {
      details: { part: partNameOf(name) },
    });
  }
}
