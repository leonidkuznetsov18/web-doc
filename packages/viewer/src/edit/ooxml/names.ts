/*
 * OPC part names (ECMA-376 Part 2 §9.1): absolute, "/"-separated, compared
 * case-insensitively. ZIP entry names carry them without the leading slash.
 */

/** "/ppt/slides/slide1.xml" for a ZIP entry name or a loosely written name. */
export function partNameOf(entryName: string): string {
  const name = entryName.replaceAll("\\", "/");
  return name.startsWith("/") ? name : `/${name}`;
}

/** The ZIP entry name of a part: no leading slash. */
export function entryNameOf(partName: string): string {
  return partNameOf(partName).slice(1);
}

/** The key two part names are compared by. */
export function partKey(name: string): string {
  return partNameOf(name).toLowerCase();
}

/** Whether a ZIP entry is a directory marker rather than a part. */
export function isDirectoryEntry(entryName: string): boolean {
  return entryName.endsWith("/");
}

/** "/ppt/slides/" for "/ppt/slides/slide1.xml"; "/" at the root. */
export function folderOf(partName: string): string {
  const name = partNameOf(partName);
  return name.slice(0, name.lastIndexOf("/") + 1);
}

/** The extension without the dot, lower-case; "" when there is none. */
export function extensionOf(partName: string): string {
  const name = partNameOf(partName);
  const slash = name.lastIndexOf("/");
  const dot = name.lastIndexOf(".");
  return dot > slash ? name.slice(dot + 1).toLowerCase() : "";
}

/** The relationships part of a part ("/" for the package): "/ppt/slides/_rels/slide1.xml.rels". */
export function relationshipsPartOf(partName: string): string {
  if (partName === "/" || partName === "") return "/_rels/.rels";
  const name = partNameOf(partName);
  const slash = name.lastIndexOf("/");
  return `${name.slice(0, slash + 1)}_rels/${name.slice(slash + 1)}.rels`;
}

/**
 * Resolves a relationship target against its source part: relative targets
 * against the source's folder, absolute ones as they are; "." and ".."
 * segments collapse. The result is an absolute part name.
 */
export function resolveTarget(sourcePart: string, target: string): string {
  const base =
    sourcePart === "/" || sourcePart === "" ? "/" : folderOf(sourcePart);
  const raw = target.replaceAll("\\", "/");
  const combined = raw.startsWith("/") ? raw : `${base}${raw}`;
  const segments: string[] = [];
  for (const segment of combined.split("/")) {
    if (segment === "" || segment === ".") continue;
    if (segment === "..") segments.pop();
    else segments.push(segment);
  }
  return `/${segments.join("/")}`;
}
