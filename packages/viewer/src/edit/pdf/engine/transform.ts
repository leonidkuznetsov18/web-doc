import type { PageRect } from "../../types.js";
import type {
  DeleteElementOperation,
  MoveElementOperation,
  PdfElement,
  ResizeElementOperation,
} from "../types.js";
import { pageToUser, type PageGeometry } from "./geometry.js";
import type {
  ElementLocation,
  Issue,
  OperationContext,
  OperationHandler,
  OperationResult,
} from "./operations.js";
import {
  changed,
  rebuildTextBox,
  removeObjects,
  validateRect,
  type TextBoxSpec,
} from "./text-box.js";
import { changedTable, rebuildTable, tableSpecOf } from "./tables.js";

/*
 * Geometry operations that apply to any element. Plain objects are moved and
 * stretched through PDFium matrices in user space; web-doc's own composites
 * are rebuilt from their stored inputs so the mark stays true.
 */

export const moveElement: OperationHandler<MoveElementOperation> = {
  validate(operation, context, issue) {
    const target = anyTarget(operation.target, context, issue);
    if (!target) return;
    if (target.element.kind === "paragraph") {
      issue(
        "/target",
        "unsupported-target",
        "Imported paragraphs cannot be moved",
      );
      return;
    }
    if ((operation.to === undefined) === (operation.by === undefined)) {
      issue("", "one-of", "Give exactly one of `to` and `by`");
      return;
    }
    const delta = moveDelta(operation, target.element.bounds);
    validateRect(
      shifted(target.element.bounds, delta.dx, delta.dy),
      context.geometry(target.location.pageIndex),
      issue,
      operation.to ? "/to" : "/by",
      target.element.bounds,
    );
  },
  apply(operation, context) {
    const { location, element } = anyTarget(operation.target, context)!;
    const { dx, dy } = moveDelta(operation, element.bounds);
    const spec = textBoxSpec(location);
    if (spec)
      return changed(
        location,
        rebuildTextBox(context, location, {
          ...spec,
          rect: shifted(spec.rect, dx, dy),
        }),
      );
    const table = tableSpecOf(location.record.mark);
    if (table)
      return changedTable(
        location,
        rebuildTable(context, location, {
          ...table,
          at: { x: table.at.x + dx, y: table.at.y + dy },
        }),
      );
    const geometry = context.geometry(location.pageIndex);
    const from = pageToUser(geometry, 0, 0);
    const to = pageToUser(geometry, dx, dy);
    transformObjects(context, location, [
      1,
      0,
      0,
      1,
      to.x - from.x,
      to.y - from.y,
    ]);
    return changed(location, { overflow: false });
  },
};

export const resizeElement: OperationHandler<ResizeElementOperation> = {
  validate(operation, context, issue) {
    const target = anyTarget(operation.target, context, issue);
    if (!target) return;
    if (
      target.element.kind === "table" ||
      target.element.kind === "paragraph"
    ) {
      issue(
        "/target",
        "unsupported-target",
        "Tables and imported paragraphs cannot be resized",
      );
      return;
    }
    validateRect(
      operation.rect,
      context.geometry(target.location.pageIndex),
      issue,
      "/rect",
      target.element.bounds,
    );
  },
  apply(operation, context) {
    const { location, element } = anyTarget(operation.target, context)!;
    const spec = textBoxSpec(location);
    if (spec)
      return changed(
        location,
        rebuildTextBox(context, location, { ...spec, rect: operation.rect }),
      );
    const geometry = context.geometry(location.pageIndex);
    const { bounds } = element;
    // Scale about the top-left corner in page space, then move it into place.
    const scaleX = operation.rect.width / bounds.width;
    const scaleY = operation.rect.height / bounds.height;
    const [userScaleX, userScaleY] =
      geometry.rotation % 2 === 0 ? [scaleX, scaleY] : [scaleY, scaleX];
    const anchor = pageToUser(geometry, bounds.x, bounds.y);
    transformObjects(context, location, [
      userScaleX,
      0,
      0,
      userScaleY,
      anchor.x - userScaleX * anchor.x,
      anchor.y - userScaleY * anchor.y,
    ]);
    const from = pageToUser(geometry, 0, 0);
    const to = pageToUser(
      geometry,
      operation.rect.x - bounds.x,
      operation.rect.y - bounds.y,
    );
    transformObjects(context, location, [
      1,
      0,
      0,
      1,
      to.x - from.x,
      to.y - from.y,
    ]);
    return changed(location, { overflow: false });
  },
};

export const deleteElement: OperationHandler<DeleteElementOperation> = {
  validate(operation, context, issue) {
    anyTarget(operation.target, context, issue);
  },
  apply(operation, context) {
    const { location } = anyTarget(operation.target, context)!;
    const paragraph = context.paragraph(operation.target)?.paragraph;
    removeObjects(context, location);
    return {
      ...changed(location, { overflow: false }),
      removedIds: [
        ...new Set([
          operation.target,
          ...(paragraph?.id === operation.target ? paragraph.memberIds : []),
        ]),
      ],
    };
  },
};

/** Any element, by id, with its location; reports a missing one. */
function anyTarget(
  target: string,
  context: OperationContext,
  issue: Issue = () => {},
):
  | { readonly location: ElementLocation; readonly element: PdfElement }
  | undefined {
  const location = context.locate(target);
  const element = location && context.element(target);
  if (!location || !element) {
    issue("/target", "unknown-target", `No element ${target}`);
    return undefined;
  }
  return { location, element };
}

function textBoxSpec(location: ElementLocation): TextBoxSpec | undefined {
  return location.record.mark?.kind === "textBox"
    ? (location.record.mark as unknown as TextBoxSpec)
    : undefined;
}

function moveDelta(
  operation: MoveElementOperation,
  bounds: PageRect,
): { readonly dx: number; readonly dy: number } {
  return operation.to
    ? { dx: operation.to.x - bounds.x, dy: operation.to.y - bounds.y }
    : { dx: operation.by!.dx, dy: operation.by!.dy };
}

function shifted(rect: PageRect, dx: number, dy: number): PageRect {
  return { ...rect, x: rect.x + dx, y: rect.y + dy };
}

/** Concatenates a user-space matrix onto every object of an element. */
function transformObjects(
  context: OperationContext,
  location: ElementLocation,
  [a, b, c, d, e, f]: readonly [number, number, number, number, number, number],
): void {
  const { lib } = context.pdfium;
  context.withPage(location.pageIndex, (page) => {
    for (const index of location.indexes)
      lib.FPDFPageObj_Transform(
        lib.FPDFPage_GetObject(page, index),
        a,
        b,
        c,
        d,
        e,
        f,
      );
  });
}
