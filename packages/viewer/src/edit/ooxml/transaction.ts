import type { ViewerWarning } from "../../contracts.js";
import { ViewerError } from "../../errors.js";
import {
  extensionOf,
  partKey,
  partNameOf,
  relationshipsPartOf,
} from "./names.js";
import {
  CONTENT_TYPES_NAMESPACE,
  extensionForMime,
  IMAGE_RELATIONSHIP_TYPE,
  parseContentTypes,
  parseRelationships,
  RELATIONSHIPS_NAMESPACE,
  sourceOfRelationshipsPart,
  type RelationshipSet,
} from "./opc.js";
import type { OoxmlPackage } from "./package.js";
import { applyPatches, patches, type XmlPatch } from "./patch.js";
import { encodePart, escapeAttribute, scanXml, type XmlPart } from "./xml.js";

/*
 * A transaction collects part patches, replaced, added and removed parts,
 * relationship and content-type changes, and commits them to the package
 * overlay atomically: every part is computed and verified in memory first,
 * and only then does the overlay change. Relationship and content-type
 * parts are written as patches too, so their untouched bytes stay.
 */

export interface CommittedChange {
  readonly changedParts: readonly string[];
  readonly addedParts: readonly string[];
  readonly removedParts: readonly string[];
  /** One per relationship whose internal target no longer exists. */
  readonly warnings: readonly ViewerWarning[];
}

const CONTENT_TYPES = "/[Content_Types].xml";
const XML_HEADER = '<?xml version="1.0" encoding="UTF-8" standalone="yes"?>\n';

interface RelationshipAdd {
  readonly id: string;
  readonly type: string;
  readonly target: string;
  readonly mode: "Internal" | "External";
}

export class PackageTransaction {
  readonly #pkg: OoxmlPackage;
  readonly #revision: number;
  readonly #patches = new Map<
    string,
    { readonly part: XmlPart; readonly items: XmlPatch[] }
  >();
  readonly #set = new Map<
    string,
    {
      readonly name: string;
      readonly bytes: Uint8Array;
      readonly contentType?: string;
    }
  >();
  readonly #removed = new Set<string>();
  readonly #relationshipAdds = new Map<string, RelationshipAdd[]>();
  readonly #relationshipRemoves = new Map<string, Set<string>>();
  readonly #relationshipSets = new Map<string, Promise<RelationshipSet>>();
  readonly #mediaByHash = new Map<string, string>();
  #committed = false;

  constructor(pkg: OoxmlPackage) {
    this.#pkg = pkg;
    this.#revision = pkg.revision;
  }

  /** Patches of one scanned part; the scan must be of the current revision. */
  patch(part: XmlPart, items: readonly XmlPatch[]): void {
    this.#open();
    if (part.revision !== this.#revision)
      throw new ViewerError(
        "invalid-patch",
        `Part ${part.name} was scanned at another revision`,
        {
          details: {
            part: part.name,
            scannedAt: part.revision,
            current: this.#revision,
          },
        },
      );
    const key = partKey(part.name);
    if (this.#set.has(key) || this.#removed.has(key))
      throw new ViewerError(
        "invalid-patch",
        `Part ${part.name} is replaced or removed in this transaction`,
        {
          details: { part: part.name },
        },
      );
    const existing = this.#patches.get(key);
    if (existing) existing.items.push(...items);
    else this.#patches.set(key, { part, items: [...items] });
  }

  /** Adds or replaces a part; `contentType` registers it unless a Default covers the extension. */
  setPart(name: string, bytes: Uint8Array, contentType?: string): void {
    this.#open();
    const partName = partNameOf(name);
    const key = partKey(partName);
    if (this.#patches.has(key))
      throw new ViewerError(
        "invalid-patch",
        `Part ${partName} is patched in this transaction`,
        {
          details: { part: partName },
        },
      );
    this.#removed.delete(key);
    this.#set.set(key, {
      name: partName,
      bytes: bytes.slice(),
      ...(contentType ? { contentType } : {}),
    });
  }

  removePart(name: string): void {
    this.#open();
    const key = partKey(name);
    this.#patches.delete(key);
    this.#set.delete(key);
    this.#removed.add(key);
  }

  /** Allocates the next free rId of the source's .rels part, pending adds included. */
  async addRelationship(
    sourcePart: string,
    type: string,
    target: string,
    mode: "Internal" | "External" = "Internal",
  ): Promise<string> {
    this.#open();
    const source = sourcePart === "/" ? "/" : partNameOf(sourcePart);
    const set = await this.#relationships(source);
    const pending = this.#relationshipAdds.get(source) ?? [];
    const taken = new Set(pending.map((add) => add.id));
    const id = set.nextId(taken);
    pending.push({ id, type, target, mode });
    this.#relationshipAdds.set(source, pending);
    return id;
  }

  removeRelationship(sourcePart: string, id: string): void {
    this.#open();
    const source = sourcePart === "/" ? "/" : partNameOf(sourcePart);
    const pending = this.#relationshipAdds.get(source);
    if (pending) {
      const index = pending.findIndex((add) => add.id === id);
      if (index >= 0) {
        pending.splice(index, 1);
        return;
      }
    }
    const removes = this.#relationshipRemoves.get(source) ?? new Set<string>();
    removes.add(id);
    this.#relationshipRemoves.set(source, removes);
  }

  /**
   * Stores media once per distinct content under `folder`, registers the
   * extension's content type when missing, and relates it to `sourcePart`.
   */
  async addMedia(
    sourcePart: string,
    folder: string,
    bytes: Uint8Array,
    mimeType: string,
    relationshipType: string = IMAGE_RELATIONSHIP_TYPE,
  ): Promise<{ readonly part: string; readonly rId: string }> {
    this.#open();
    const extension = extensionForMime(mimeType);
    const hash = `${folder}|${extension}|${await sha256(bytes)}`;
    let part = this.#mediaByHash.get(hash);
    if (!part) {
      part = this.uniquePartName(
        `${partNameOf(folder).replace(/\/?$/, "/")}image`,
        `.${extension}`,
      );
      this.#mediaByHash.set(hash, part);
      const types = await this.#contentTypes();
      const needsDefault =
        !types.defaults.has(extension) &&
        ![...this.#set.values()].some(
          (entry) => entry.contentType && extensionOf(entry.name) === extension,
        );
      this.setPart(part, bytes, needsDefault ? mimeType : undefined);
    }
    const source = sourcePart === "/" ? "/" : partNameOf(sourcePart);
    const rId = await this.addRelationship(
      source,
      relationshipType,
      relativeTarget(source, part),
    );
    return { part, rId };
  }

  /** A part name that does not exist yet, in the package or in this transaction. */
  uniquePartName(prefix: string, extension: string): string {
    const base = partNameOf(prefix);
    for (let n = 1; ; n += 1) {
      const candidate = `${base}${n}${extension}`;
      const key = partKey(candidate);
      if (this.#removed.has(key)) return candidate;
      if (!this.#pkg.has(candidate) && !this.#set.has(key)) return candidate;
    }
  }

  /** Applies everything or nothing; returns what changed. */
  async commit(signal?: AbortSignal): Promise<CommittedChange> {
    this.#open();
    if (signal?.aborted)
      throw new ViewerError("aborted", "The transaction was aborted");
    if (this.#pkg.revision !== this.#revision)
      throw new ViewerError(
        "edit-conflict",
        "The package changed since the transaction started",
        {
          details: { expected: this.#revision, actual: this.#pkg.revision },
        },
      );
    const set = new Map<string, Uint8Array>();
    const remove = new Set<string>();
    const next = this.#revision + 1;

    // 1. Patched parts, verified by re-scan.
    for (const { part, items } of this.#patches.values()) {
      const patched = applyPatches(part, items, next);
      set.set(part.name, encodePart(patched.text));
    }
    // 2. Replaced and added parts; XML parts must scan.
    for (const entry of this.#set.values()) {
      if (/\.(xml|rels)$/i.test(entry.name)) {
        const text = new TextDecoder("utf-8", { fatal: true }).decode(
          entry.bytes,
        );
        scanXml(entry.name, text, next);
      }
      set.set(entry.name, entry.bytes);
    }
    // 3. Removed parts take their own .rels part along.
    for (const key of this.#removed) {
      const name =
        this.#pkg.currentPartNames.find(
          (candidate) => partKey(candidate) === key,
        ) ?? key;
      if (!this.#pkg.has(name)) continue;
      remove.add(name);
      const rels = relationshipsPartOf(name);
      if (this.#pkg.has(rels)) remove.add(rels);
    }
    // 4. Relationship parts, as patches or new parts.
    const relationshipSources = new Set([
      ...this.#relationshipAdds.keys(),
      ...this.#relationshipRemoves.keys(),
    ]);
    for (const source of relationshipSources) {
      const adds = this.#relationshipAdds.get(source) ?? [];
      const removes =
        this.#relationshipRemoves.get(source) ?? new Set<string>();
      if (adds.length === 0 && removes.size === 0) continue;
      const relsName = relationshipsPartOf(source);
      const existing = await this.#relationships(source);
      const fragments = adds.map(
        (add) =>
          `<Relationship Id=${escapeAttribute(add.id)} Type=${escapeAttribute(add.type)} Target=${escapeAttribute(add.target)}${add.mode === "External" ? ' TargetMode="External"' : ""}/>`,
      );
      if (
        existing.partName &&
        this.#pkg.has(existing.partName) &&
        !remove.has(existing.partName)
      ) {
        const part = await this.#pkg.xml(existing.partName, signal);
        const items: XmlPatch[] = [];
        for (const id of removes) {
          const node = existing.byId(id)?.node;
          if (!node)
            throw new ViewerError(
              "invalid-patch",
              `No relationship ${id} in ${existing.partName}`,
              {
                details: { part: existing.partName, id },
              },
            );
          items.push(patches.removeElement(part, node));
        }
        // One patch per element: each must read back as one element.
        for (const fragment of fragments)
          items.push(patches.appendChild(part, part.root, fragment));
        const patched = applyPatches(part, items, next);
        set.set(existing.partName, encodePart(patched.text));
      } else {
        if (removes.size > 0)
          throw new ViewerError(
            "invalid-patch",
            `No relationships part for ${source}`,
            { details: { part: relsName } },
          );
        set.set(
          relsName,
          encodePart(
            `${XML_HEADER}<Relationships xmlns="${RELATIONSHIPS_NAMESPACE}">${fragments.join("")}</Relationships>`,
          ),
        );
      }
    }
    // 5. Content types: Overrides for new parts with a type, Defaults for media, removals.
    const types = await this.#contentTypes();
    const typesPart = await this.#pkg.xml(CONTENT_TYPES, signal);
    const typeItems: XmlPatch[] = [];
    const additions: string[] = [];
    const seenDefaults = new Set<string>();
    for (const entry of this.#set.values()) {
      if (!entry.contentType) continue;
      const extension = extensionOf(entry.name);
      const isMedia = !/\.(xml|rels)$/i.test(entry.name) && extension !== "";
      if (
        isMedia &&
        !types.defaults.has(extension) &&
        !seenDefaults.has(extension)
      ) {
        seenDefaults.add(extension);
        additions.push(
          `<Default Extension=${escapeAttribute(extension)} ContentType=${escapeAttribute(entry.contentType)}/>`,
        );
      } else if (
        !isMedia &&
        types.defaults.get(extension) !== entry.contentType
      ) {
        const node = types.overrideNode(entry.name);
        if (node)
          typeItems.push(
            patches.setAttribute(
              typesPart,
              node,
              "ContentType",
              entry.contentType,
            ),
          );
        else
          additions.push(
            `<Override PartName=${escapeAttribute(entry.name)} ContentType=${escapeAttribute(entry.contentType)}/>`,
          );
      }
    }
    for (const name of remove) {
      const node = types.overrideNode(name);
      if (node) typeItems.push(patches.removeElement(typesPart, node));
    }
    for (const addition of additions)
      typeItems.push(patches.appendChild(typesPart, typesPart.root, addition));
    if (typeItems.length > 0) {
      const patched = applyPatches(typesPart, typeItems, next);
      set.set(CONTENT_TYPES, encodePart(patched.text));
    }
    // 6. Dangling relationship targets after the removals, as warnings.
    const warnings: ViewerWarning[] = [];
    if (remove.size > 0) {
      const removedKeys = new Set([...remove].map(partKey));
      for (const name of this.#pkg.currentPartNames) {
        const source = sourceOfRelationshipsPart(name);
        if (source === undefined || remove.has(name)) continue;
        const relsPart = set.has(name)
          ? scanXml(name, new TextDecoder().decode(set.get(name)!), next)
          : await this.#pkg.xml(name, signal);
        const rels = parseRelationships(source, name, relsPart);
        for (const item of rels.items)
          if (item.targetPart && removedKeys.has(partKey(item.targetPart)))
            warnings.push({
              code: "fidelity-degraded",
              message: `Relationship ${item.id} of ${source} points at the removed part ${item.targetPart}`,
              details: {
                source,
                id: item.id,
                target: item.targetPart,
                reason: "dangling-relationship",
              },
            });
      }
    }
    if (signal?.aborted)
      throw new ViewerError("aborted", "The transaction was aborted");
    const added = [...set.keys()].filter((name) => !this.#pkg.has(name));
    const changed = [...set.keys()].filter((name) => this.#pkg.has(name));
    this.#pkg.applyOverlay({ set, remove });
    this.#committed = true;
    return {
      changedParts: changed,
      addedParts: added,
      removedParts: [...remove],
      warnings,
    };
  }

  #open(): void {
    if (this.#committed)
      throw new ViewerError(
        "lifecycle-error",
        "The transaction was already committed",
      );
  }

  #relationships(source: string): Promise<RelationshipSet> {
    let pending = this.#relationshipSets.get(source);
    if (!pending) {
      pending = (async () => {
        const relsName = relationshipsPartOf(source);
        if (!this.#pkg.has(relsName) || this.#removed.has(partKey(relsName)))
          return parseRelationships(source, undefined, undefined);
        return parseRelationships(
          source,
          relsName,
          await this.#pkg.xml(relsName),
        );
      })();
      this.#relationshipSets.set(source, pending);
    }
    return pending;
  }

  async #contentTypes() {
    return parseContentTypes(await this.#pkg.xml(CONTENT_TYPES));
  }
}

/** "../media/image1.png" for a target seen from a source part. */
export function relativeTarget(sourcePart: string, targetPart: string): string {
  if (sourcePart === "/") return partNameOf(targetPart).slice(1);
  const from = partNameOf(sourcePart).split("/").slice(1, -1);
  const to = partNameOf(targetPart).split("/").slice(1);
  let common = 0;
  while (
    common < from.length &&
    common < to.length - 1 &&
    from[common]!.toLowerCase() === to[common]!.toLowerCase()
  )
    common += 1;
  const up = from.length - common;
  return [...Array.from({ length: up }, () => ".."), ...to.slice(common)].join(
    "/",
  );
}

async function sha256(bytes: Uint8Array): Promise<string> {
  const digest = await crypto.subtle.digest(
    "SHA-256",
    bytes.slice() as Uint8Array<ArrayBuffer>,
  );
  return Array.from(new Uint8Array(digest), (byte) =>
    byte.toString(16).padStart(2, "0"),
  ).join("");
}

export { CONTENT_TYPES_NAMESPACE };
