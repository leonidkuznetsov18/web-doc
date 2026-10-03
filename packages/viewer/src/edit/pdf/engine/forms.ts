import { ViewerError } from "../../../errors.js";
import { readObjects, type PdfObjectBytes } from "./compact.js";
import { objectMatrix, OBJECT_FORM } from "./elements.js";
import { concat, invert, type Matrix } from "./geometry.js";
import type { Pdfium } from "./pdfium.js";

/*
 * Objects inside Form XObjects. PDFium parses a form's objects but never
 * writes changes to them back: only the page's own content is regenerated.
 * So an edit inside a form rewrites the form, copy on write: the objects of
 * each form on the way to the edited one move onto a stand-in page, the edit
 * runs there, and each stand-in becomes a new Form XObject drawn in the old
 * form's place with the old form's matrix. Only that one drawing of the form
 * changes; other drawings of the same XObject keep the original.
 *
 * The objects keep their own matrices and clip paths, which already carry
 * the form's /Matrix and /BBox, so the new form needs neither. The clip path
 * the old form was drawn with moves to the new one. What a form's
 * dictionary says beyond that, such as a transparency group or
 * optional content, is not carried over; `rewritesFaithfully` in the
 * document checks by rendering that a rewrite leaves the page looking the
 * same. The objects' graphics states are written back by name, and those
 * names belong to the old forms' resources, which the stand-ins do not
 * have; `graphicsStatesResolve` finds a rewrite that would leave them
 * dangling, and the document refuses it.
 */

/** A form's objects in drawing order. */
export function formChildren(pdfium: Pdfium, form: number): number[] {
  const { lib } = pdfium;
  return Array.from(
    { length: Math.max(0, lib.FPDFFormObj_CountObjects(form)) },
    (_, index) => lib.FPDFFormObj_GetObject(form, index),
  );
}

/** The object `path` leads to: a page object, then objects of the forms it passes. */
export function objectAt(
  pdfium: Pdfium,
  page: number,
  path: readonly number[],
): number {
  const { lib } = pdfium;
  let object = lib.FPDFPage_GetObject(page, path[0]!);
  for (const index of path.slice(1))
    object = object ? lib.FPDFFormObj_GetObject(object, index) : 0;
  if (!object)
    throw new ViewerError("internal", "No object at this path", {
      details: { path },
    });
  return object;
}

/** What maps the space of the innermost form of `forms` onto the page. */
export function formsMatrix(
  pdfium: Pdfium,
  page: number,
  forms: readonly number[],
): Matrix {
  let matrix: Matrix = [1, 0, 0, 1, 0, 0];
  forms.forEach((_, depth) => {
    const form = objectAt(pdfium, page, forms.slice(0, depth + 1));
    matrix = concat(objectMatrix(pdfium, form), matrix);
  });
  return matrix;
}

/**
 * Runs `use` on a stand-in page holding the objects of the innermost form
 * of `forms`, then writes the chain of forms back as new XObjects in place
 * of the old ones. The caller regenerates `page` afterwards; if `use` or the
 * rewrite throws, the caller must close `page` without regenerating it.
 * `inspect` sees the object numbers of the stand-ins' page dictionaries
 * once they are written, before they are deleted.
 */
export function rewriteForms<T>(
  pdfium: Pdfium,
  document: number,
  page: number,
  forms: readonly number[],
  use: (holder: number) => T,
  inspect?: (standIns: readonly number[]) => void,
): T {
  const { lib } = pdfium;
  const chain = forms.map((_, depth) => {
    const form = objectAt(pdfium, page, forms.slice(0, depth + 1));
    if (lib.FPDFPageObj_GetType(form) !== OBJECT_FORM)
      throw new ViewerError(
        "internal",
        "The path does not lead through forms",
        {
          details: { forms },
        },
      );
    return form;
  });
  const first = lib.FPDF_GetPageCount(document);
  const holders: number[] = [];
  const standIns: number[] = [];
  try {
    for (const form of chain) {
      const holder = lib.FPDFPage_New(document, first + holders.length, 1, 1);
      if (!holder)
        throw new ViewerError("edit-failed", "PDFium could not add a page", {
          details: { stage: "apply", reason: "form-rewrite" },
        });
      holders.push(holder);
      standIns.push(
        lib.EPDFDoc_GetPageObjectNumberByIndex(
          document,
          first + holders.length - 1,
        ),
      );
      for (const child of formChildren(pdfium, form)) {
        lib.FPDFFormObj_RemoveObject(form, child);
        lib.FPDFPage_InsertObject(holder, child);
      }
    }
    const result = use(holders.at(-1)!);
    let rewritten = 0;
    for (let depth = chain.length - 1; depth >= 0; depth -= 1) {
      const holder = holders[depth]!;
      const old = chain[depth]!;
      if (rewritten) swap(pdfium, holder, forms[depth + 1]!, rewritten);
      rewritten = asForm(pdfium, document, holder, first + depth, old);
      const [a, b, c, d, e, f] = objectMatrix(pdfium, old);
      lib.FPDFPageObj_Transform(rewritten, a, b, c, d, e, f);
    }
    swap(pdfium, page, forms[0]!, rewritten);
    inspect?.(standIns);
    return result;
  } finally {
    // Innermost first, so the pages before keep their indexes.
    for (let depth = holders.length - 1; depth >= 0; depth -= 1) {
      lib.FPDF_ClosePage(holders[depth]!);
      lib.FPDFPage_Delete(document, first + depth);
    }
  }
}

/** Puts `object` where the object at `index` of `holder` was, and destroys that one. */
function swap(
  pdfium: Pdfium,
  holder: number,
  index: number,
  object: number,
): void {
  const { lib } = pdfium;
  const old = lib.FPDFPage_GetObject(holder, index);
  lib.FPDFPage_RemoveObject(holder, old);
  lib.FPDFPageObj_Destroy(old);
  lib.FPDFPage_InsertObjectAtIndex(holder, object, index);
}

/**
 * A new form object drawing what the stand-in page at `pageIndex` holds, in
 * the stand-in's space, cut by the clip path `old` was drawn with. The box
 * covers every object generously: what the old form's box cut off is cut
 * off by the clip paths the objects carry.
 */
function asForm(
  pdfium: Pdfium,
  document: number,
  holder: number,
  pageIndex: number,
  old: number,
): number {
  const { lib } = pdfium;
  const [left, bottom, right, top] = holderBounds(pdfium, holder);
  const margin = Math.max(right - left, top - bottom) / 2 + 10;
  const origin = { x: left - margin, y: bottom - margin };
  lib.FPDFPage_SetMediaBox(
    holder,
    origin.x,
    origin.y,
    right + margin,
    top + margin,
  );
  if (!lib.FPDFPage_GenerateContent(holder))
    throw new ViewerError("edit-failed", "PDFium could not rewrite a form", {
      details: { stage: "apply", reason: "form-rewrite" },
    });
  const clip = lib.FPDFPageObj_GetClipPath(old);
  if (clip && lib.FPDFClipPath_CountPaths(clip) > 0) {
    // The clip is in the space the form is drawn into; the stand-in holds
    // the form's own space. The old object goes away, so it is changed.
    const [a, b, c, d, e, f] = invert(objectMatrix(pdfium, old));
    lib.FPDFPageObj_TransformClipPath(old, a, b, c, d, e, f);
    lib.FPDFPage_InsertClipPath(holder, lib.FPDFPageObj_GetClipPath(old));
  }
  const xobject = lib.FPDF_NewXObjectFromPage(document, document, pageIndex);
  const form = xobject ? lib.FPDF_NewFormObjectFromXObject(xobject) : 0;
  if (xobject) lib.FPDF_CloseXObject(xobject);
  if (!form)
    throw new ViewerError("edit-failed", "PDFium could not rewrite a form", {
      details: { stage: "apply", reason: "form-rewrite" },
    });
  // PDFium moves a page's box to the origin when it makes the XObject.
  lib.FPDFPageObj_Transform(form, 1, 0, 0, 1, origin.x, origin.y);
  return form;
}

function holderBounds(
  pdfium: Pdfium,
  holder: number,
): readonly [number, number, number, number] {
  const { lib } = pdfium;
  let box: [number, number, number, number] | undefined;
  for (let index = 0; index < lib.FPDFPage_CountObjects(holder); index += 1) {
    const bounds = pdfium.readNumbers(4, "float", ([l, b, r, t]) =>
      lib.FPDFPageObj_GetBounds(
        lib.FPDFPage_GetObject(holder, index),
        l!,
        b!,
        r!,
        t!,
      ),
    );
    if (!bounds) continue;
    const [left, bottom, right, top] = bounds as [
      number,
      number,
      number,
      number,
    ];
    box = box
      ? [
          Math.min(box[0], left),
          Math.min(box[1], bottom),
          Math.max(box[2], right),
          Math.max(box[3], top),
        ]
      : [left, bottom, right, top];
  }
  return box ?? [0, 0, 1, 1];
}

/** A `/Name gs` operator in content PDFium wrote; its strings are hex, so none can fake one. */
const SET_GRAPHICS_STATE = /\/([^\s/[\]()<>{}%]+)\s+gs(?=\s|$)/g;

/**
 * Whether every graphics state the stand-in pages' content sets by name is
 * one their resources hold, in `bytes`, a file PDFium wrote after a rewrite.
 * A name the old form's resources defined would be left dangling inside the
 * new form: readers then drop what it set, and some report an error.
 */
export async function graphicsStatesResolve(
  bytes: Uint8Array,
  standIns: readonly number[],
): Promise<boolean> {
  const objects = readObjects(bytes);
  for (const number of standIns) {
    const page = objects.get(number)?.value;
    if (!page) return false;
    const held = new Set(
      [
        ...(/\/ExtGState\s*<<([^>]*)>>/.exec(page)?.[1] ?? "").matchAll(
          /\/([^\s/<>[\]]+)\s+\d+\s+\d+\s+R/g,
        ),
      ].map((match) => match[1]!),
    );
    for (const reference of contentReferences(page, objects)) {
      const stream = objects.get(reference);
      if (!stream?.data || stream.value === undefined) return false;
      const content = /\/FlateDecode/.test(stream.value)
        ? await inflate(stream.data)
        : stream.data;
      if (!content) return false;
      for (const match of latin1(content).matchAll(SET_GRAPHICS_STATE))
        if (!held.has(match[1]!)) return false;
    }
  }
  return true;
}

/**
 * Object numbers of a page dictionary's content streams: one stream, or an
 * array of them, direct or itself an object, as clip paths leave it.
 */
function contentReferences(
  page: string,
  objects: ReadonlyMap<number, PdfObjectBytes>,
): number[] {
  const contents =
    /\/Contents\s*(\[[^\]]*\]|\d+\s+\d+\s+R)/.exec(page)?.[1] ?? "";
  const references = (text: string): number[] =>
    [...text.matchAll(/(\d+)\s+\d+\s+R/g)].map((match) => Number(match[1]));
  return references(contents).flatMap((reference) => {
    const array = objects.get(reference);
    return !array?.data && array?.value?.trimStart().startsWith("[")
      ? references(array.value)
      : [reference];
  });
}

/** Flate data through the platform's zlib stream, or nothing where there is none. */
async function inflate(data: Uint8Array): Promise<Uint8Array | undefined> {
  if (typeof DecompressionStream !== "function") return undefined;
  try {
    const stream = new Blob([data.slice()])
      .stream()
      .pipeThrough(new DecompressionStream("deflate"));
    return new Uint8Array(await new Response(stream).arrayBuffer());
  } catch {
    return undefined;
  }
}

function latin1(bytes: Uint8Array): string {
  let text = "";
  for (let offset = 0; offset < bytes.length; offset += 0x8000)
    text += String.fromCharCode(...bytes.subarray(offset, offset + 0x8000));
  return text;
}
