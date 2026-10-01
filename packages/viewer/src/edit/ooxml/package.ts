import type { ResourceLimits } from "../../contracts.js";
import { ViewerError } from "../../errors.js";
import { isDirectoryEntry, partKey, partNameOf } from "./names.js";
import {
  inflateEntry,
  parseZip,
  type ZipArchive,
  type ZipEntry,
} from "./zip.js";

/*
 * An OOXML package held in memory: the original bytes, the ZIP directory,
 * and a cache of the parts that were read. Parts are inflated on first use.
 * This is the reading half (task 38); the overlay, transactions and saving
 * follow in the later tasks of the module.
 */

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

  has(name: string): boolean {
    return this.#entries.has(partKey(name));
  }

  /** The ZIP entry behind a part. */
  entry(name: string): ZipEntry | undefined {
    return this.#entries.get(partKey(name));
  }

  /** The original bytes of a part, inflated once and cached. */
  part(name: string, signal?: AbortSignal): Promise<Uint8Array> {
    const key = partKey(name);
    const entry = this.#entries.get(key);
    if (!entry)
      return Promise.reject(
        new ViewerError("invalid-file", `No part ${partNameOf(name)}`, {
          details: { part: partNameOf(name) },
        }),
      );
    let pending = this.#parts.get(key);
    if (!pending) {
      pending = inflateEntry(this.#archive, entry, this.#limits, signal);
      this.#parts.set(key, pending);
      pending.catch(() => this.#parts.delete(key));
    }
    return pending.then((bytes) => bytes.slice());
  }
}
