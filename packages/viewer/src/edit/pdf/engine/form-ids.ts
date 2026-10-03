import { ViewerError } from "../../../errors.js";
import { OBJECT_FORM, type ObjectRecord } from "./elements.js";
import { MAX_FORM_DEPTH } from "./forms.js";
import { isPageKey } from "./page-keys.js";
import type { Pdfium } from "./pdfium.js";

/** Identity belongs to this drawing, not to its potentially shared XObject. */
const MARK_NAME = "WebDocFormIds";
const MARK_PARAM = "ids";
const MAX_MARK_BYTES = 1_048_576;
const MAX_ID_LENGTH = 256;

/** Only identities and the native structure are persisted, never rendering inputs. */
function identity(record: ObjectRecord): ObjectRecord {
  return {
    id: record.id,
    type: record.type,
    ...(record.children ? { children: record.children.map(identity) } : {}),
  };
}

export function writeFormIds(
  pdfium: Pdfium,
  document: number,
  object: number,
  record: ObjectRecord,
): void {
  const raw = JSON.stringify(identity(record));
  if ((raw.length + 1) * 2 > MAX_MARK_BYTES)
    throw new ViewerError("edit-failed", "The form identity tree is too large");
  const { lib } = pdfium;
  for (
    let index = lib.FPDFPageObj_CountMarks(object) - 1;
    index >= 0;
    index -= 1
  ) {
    const mark = lib.FPDFPageObj_GetMark(object, index);
    if (
      pdfium.readWideStringOut(
        (buffer, bytes, out) =>
          lib.FPDFPageObjMark_GetName(mark, buffer, bytes, out),
        256,
      ) === MARK_NAME
    )
      lib.FPDFPageObj_RemoveMark(object, mark);
  }
  const mark = lib.FPDFPageObj_AddMark(object, MARK_NAME);
  if (!mark || !pdfium.setMarkString(document, object, mark, MARK_PARAM, raw))
    throw new ViewerError(
      "edit-failed",
      "PDFium could not preserve form identities",
    );
}

/** Foreign, damaged, stale or colliding metadata never changes object ownership. */
export function readFormIds(
  pdfium: Pdfium,
  object: number,
  pageKey: string,
  unavailable: (id: string) => boolean,
): ObjectRecord | undefined {
  const { lib } = pdfium;
  for (let index = 0; index < lib.FPDFPageObj_CountMarks(object); index += 1) {
    const mark = lib.FPDFPageObj_GetMark(object, index);
    if (
      pdfium.readWideStringOut(
        (buffer, bytes, out) =>
          lib.FPDFPageObjMark_GetName(mark, buffer, bytes, out),
        256,
      ) !== MARK_NAME
    )
      continue;
    const raw = pdfium.readWideStringOut(
      (buffer, bytes, out) =>
        lib.FPDFPageObjMark_GetParamStringValue(
          mark,
          MARK_PARAM,
          buffer,
          bytes,
          out,
        ),
      MAX_MARK_BYTES,
    );
    try {
      const value: unknown = JSON.parse(raw);
      return validated(value, object, 0, new Set());
    } catch {
      // Metadata is untrusted; the caller retains native path identities.
      return undefined;
    }
  }
  return undefined;

  function validated(
    value: unknown,
    object: number,
    depth: number,
    used: Set<string>,
    sourceKey?: string,
  ): ObjectRecord | undefined {
    if (
      !value ||
      typeof value !== "object" ||
      !("id" in value) ||
      !("type" in value)
    )
      return undefined;
    const { id: storedId, type } = value;
    if (
      typeof storedId !== "string" ||
      !/^[^:]+:[^\s]+$/.test(storedId) ||
      type !== lib.FPDFPageObj_GetType(object)
    )
      return undefined;
    const owner = sourceKey ?? storedId.slice(0, storedId.indexOf(":"));
    if (!isPageKey(owner) || !storedId.startsWith(`${owner}:`))
      return undefined;
    // Like ordinary WebDoc marks, a drawing imported onto another page
    // keeps its suffix but belongs to that page for target dispatch.
    const id = storedId.startsWith(`${pageKey}:`)
      ? storedId
      : `${pageKey}:${storedId}`;
    if (id.length > MAX_ID_LENGTH || unavailable(id) || used.has(id))
      return undefined;
    used.add(id);
    if (type !== OBJECT_FORM || depth === MAX_FORM_DEPTH)
      return "children" in value ? undefined : { id, type };
    if (
      !("children" in value) ||
      !Array.isArray(value.children) ||
      value.children.length !== lib.FPDFFormObj_CountObjects(object)
    )
      return undefined;
    const children: ObjectRecord[] = [];
    for (let index = 0; index < value.children.length; index += 1) {
      const child = validated(
        value.children[index],
        lib.FPDFFormObj_GetObject(object, index),
        depth + 1,
        used,
        owner,
      );
      if (!child) return undefined;
      children.push(child);
    }
    return { id, type, children };
  }
}
