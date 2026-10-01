import { extensionOf, partKey, partNameOf, resolveTarget } from "./names.js";
import type { XmlElement, XmlPart } from "./xml.js";

/*
 * The Open Packaging Conventions parts the layer understands as models:
 * [Content_Types].xml (Default by extension, Override by part name) and
 * relationship parts (ordered, with ids, types and resolved targets). Both
 * are read from scanned parts and written back as patches by transactions,
 * so untouched entries keep their bytes.
 */

export const CONTENT_TYPES_NAMESPACE =
  "http://schemas.openxmlformats.org/package/2006/content-types";
export const RELATIONSHIPS_NAMESPACE =
  "http://schemas.openxmlformats.org/package/2006/relationships";
export const RELATIONSHIPS_CONTENT_TYPE =
  "application/vnd.openxmlformats-package.relationships+xml";
export const IMAGE_RELATIONSHIP_TYPE =
  "http://schemas.openxmlformats.org/officeDocument/2006/relationships/image";

export interface Relationship {
  readonly id: string;
  readonly type: string;
  /** As written. */
  readonly target: string;
  readonly targetMode: "Internal" | "External";
  /** Absolute part name for internal targets. */
  readonly targetPart?: string;
  /** The element in the scanned .rels part, for patches. */
  readonly node?: XmlElement;
}

export class RelationshipSet {
  readonly #byId = new Map<string, Relationship>();
  readonly #byType = new Map<string, Relationship[]>();

  constructor(
    /** "/" for the package. */
    readonly sourcePart: string,
    /** The .rels part, when it exists. */
    readonly partName: string | undefined,
    readonly items: readonly Relationship[],
  ) {
    for (const item of items) {
      this.#byId.set(item.id, item);
      const list = this.#byType.get(item.type);
      if (list) list.push(item);
      else this.#byType.set(item.type, [item]);
    }
  }

  byId(id: string): Relationship | undefined {
    return this.#byId.get(id);
  }

  byType(type: string): readonly Relationship[] {
    return this.#byType.get(type) ?? [];
  }

  /** The first free "rId<n>", counting from 1 and skipping `taken`. */
  nextId(taken: ReadonlySet<string> = new Set()): string {
    for (let n = 1; ; n += 1) {
      const id = `rId${n}`;
      if (!this.#byId.has(id) && !taken.has(id)) return id;
    }
  }
}

/** Reads a relationships part; `part` undefined means the source has none yet. */
export function parseRelationships(
  sourcePart: string,
  partName: string | undefined,
  part: XmlPart | undefined,
): RelationshipSet {
  if (!part) return new RelationshipSet(sourcePart, partName, []);
  const items: Relationship[] = [];
  for (const node of part.root.children) {
    if (node.local !== "Relationship") continue;
    const id = part.attribute(node, "Id");
    const type = part.attribute(node, "Type");
    const target = part.attribute(node, "Target");
    if (id === undefined || type === undefined || target === undefined)
      continue;
    const external = part.attribute(node, "TargetMode") === "External";
    items.push({
      id,
      type,
      target,
      targetMode: external ? "External" : "Internal",
      ...(external ? {} : { targetPart: resolveTarget(sourcePart, target) }),
      node,
    });
  }
  return new RelationshipSet(sourcePart, partName, items);
}

export interface ContentTypes {
  /** Lower-case extension → content type. */
  readonly defaults: ReadonlyMap<string, string>;
  /** Part-name key → content type. */
  readonly overrides: ReadonlyMap<string, string>;
  typeOf(partName: string): string | undefined;
  /** The Override element of a part, for patches. */
  overrideNode(partName: string): XmlElement | undefined;
  defaultNode(extension: string): XmlElement | undefined;
}

export function parseContentTypes(part: XmlPart): ContentTypes {
  const defaults = new Map<string, string>();
  const overrides = new Map<string, string>();
  const defaultNodes = new Map<string, XmlElement>();
  const overrideNodes = new Map<string, XmlElement>();
  for (const node of part.root.children) {
    const contentType = part.attribute(node, "ContentType");
    if (contentType === undefined) continue;
    if (node.local === "Default") {
      const extension = part.attribute(node, "Extension")?.toLowerCase();
      if (extension === undefined) continue;
      defaults.set(extension, contentType);
      defaultNodes.set(extension, node);
    } else if (node.local === "Override") {
      const name = part.attribute(node, "PartName");
      if (name === undefined) continue;
      overrides.set(partKey(name), contentType);
      overrideNodes.set(partKey(name), node);
    }
  }
  return {
    defaults,
    overrides,
    typeOf(partName) {
      return (
        overrides.get(partKey(partName)) ?? defaults.get(extensionOf(partName))
      );
    },
    overrideNode(partName) {
      return overrideNodes.get(partKey(partName));
    },
    defaultNode(extension) {
      return defaultNodes.get(extension.toLowerCase());
    },
  };
}

/** The file extension conventionally used for a media type. */
export function extensionForMime(mimeType: string): string {
  switch (mimeType.toLowerCase()) {
    case "image/png":
      return "png";
    case "image/jpeg":
    case "image/jpg":
      return "jpeg";
    case "image/gif":
      return "gif";
    case "image/bmp":
      return "bmp";
    case "image/tiff":
      return "tiff";
    case "image/svg+xml":
      return "svg";
    case "image/x-emf":
    case "image/emf":
      return "emf";
    case "image/x-wmf":
    case "image/wmf":
      return "wmf";
    default:
      return "bin";
  }
}

/** The relationships part a .rels file belongs to, or undefined for other parts. */
export function sourceOfRelationshipsPart(
  partName: string,
): string | undefined {
  const name = partNameOf(partName);
  const match = /^(.*\/)_rels\/([^/]*)\.rels$/.exec(name);
  if (!match) return undefined;
  const [, folder, base] = match;
  if (folder === "/" && base === "") return "/";
  return `${folder}${base}`;
}
