import type { PagePoint } from "../../types.js";
import type {
  InsertShapeOperation,
  PdfFill,
  PdfStroke,
  SetShapeStyleOperation,
} from "../types.js";
import { OBJECT_PATH } from "./elements.js";
import { pageToUser, type PageGeometry } from "./geometry.js";
import type {
  ElementLocation,
  Issue,
  OperationContext,
  OperationHandler,
  OperationResult,
} from "./operations.js";
import { parseColor, validateRect } from "./text-box.js";

/*
 * Vector shapes as path objects. Geometry is given in page space and turned
 * into user space point by point, so shapes land where the host expects on
 * rotated and cropped pages too.
 */

/** FPDFPath_SetDrawMode fill modes. */
const FILL_NONE = 0;
const FILL_ALTERNATE = 1;

/** Bézier control-point ratio that approximates a quarter circle. */
const KAPPA = 0.5522847498;

/** Round joins keep rectangle corners tidy at wide strokes. */
const LINE_JOIN_ROUND = 1;

export const insertShape: OperationHandler<InsertShapeOperation> = {
  validate(operation, context, issue) {
    if (operation.pageIndex >= context.pageCount) {
      issue("/pageIndex", "unknown-target", `No page ${operation.pageIndex}`);
      return;
    }
    const geometry = context.geometry(operation.pageIndex);
    if (operation.shape === "line") {
      if (!operation.from || !operation.to) {
        issue("", "required", "A line needs `from` and `to`");
        return;
      }
      if ("fill" in operation && operation.fill !== undefined)
        issue("/fill", "unsupported-style", "A line cannot be filled");
      if (!operation.stroke)
        issue("/stroke", "required", "A line needs a stroke");
      for (const [name, point] of [
        ["from", operation.from],
        ["to", operation.to],
      ] as const)
        validateRect(
          { ...point, width: 0.001, height: 0.001 },
          geometry,
          issue,
          `/${name}`,
        );
      return;
    }
    if (!operation.rect) {
      issue("/rect", "required", `A ${operation.shape} needs a rect`);
      return;
    }
    validateRect(operation.rect, geometry, issue);
    if (!operation.stroke && !operation.fill)
      issue("", "required", "Give a stroke, a fill, or both");
  },
  apply(operation, context) {
    const { lib } = context.pdfium;
    const geometry = context.geometry(operation.pageIndex);
    const id = context.newId(operation.pageIndex);
    const path =
      operation.shape === "line"
        ? linePath(context, geometry, operation.from, operation.to)
        : operation.shape === "rectangle"
          ? rectanglePath(context, geometry, operation.rect)
          : ellipsePath(context, geometry, operation.rect);
    applyStyle(
      context,
      path,
      operation.stroke,
      operation.shape === "line" ? undefined : operation.fill,
    );
    context.withPage(operation.pageIndex, (page) => {
      lib.FPDFPage_InsertObject(page, path);
    });
    context.appendObjects(operation.pageIndex, [{ id, type: OBJECT_PATH }]);
    return {
      createdIds: [id],
      changedPages: [operation.pageIndex],
      warnings: [],
    };
  },
};

export const setShapeStyle: OperationHandler<SetShapeStyleOperation> = {
  validate(operation, context, issue) {
    const target = shapeTarget(operation.target, context, issue);
    if (!target) return;
    const style = target.element.shapeStyle ?? {};
    const stroke =
      operation.stroke === undefined ? style.stroke : operation.stroke;
    const fill = operation.fill === undefined ? style.fill : operation.fill;
    if (!stroke && !fill)
      issue("", "required", "A shape keeps a stroke, a fill, or both");
  },
  apply(operation, context) {
    const { lib } = context.pdfium;
    const { location, element } = shapeTarget(operation.target, context)!;
    const style = element.shapeStyle ?? {};
    const stroke =
      operation.stroke === undefined ? style.stroke : operation.stroke;
    const fill = operation.fill === undefined ? style.fill : operation.fill;
    context.withPage(location.pageIndex, (page) => {
      applyStyle(
        context,
        lib.FPDFPage_GetObject(page, location.indexes[0]!),
        stroke ?? undefined,
        fill ?? undefined,
      );
    });
    return result(location);
  },
};

function shapeTarget(
  target: string,
  context: OperationContext,
  issue: Issue = () => {},
) {
  const location = context.locate(target);
  const element = location && context.element(target);
  if (!location || !element) {
    issue("/target", "unknown-target", `No element ${target}`);
    return undefined;
  }
  if (element.kind !== "shape") {
    issue(
      "/target",
      "unsupported-target",
      "Only shapes have a stroke and fill",
    );
    return undefined;
  }
  return { location, element };
}

function applyStyle(
  context: OperationContext,
  path: number,
  stroke: PdfStroke | undefined,
  fill: PdfFill | undefined,
): void {
  const { lib } = context.pdfium;
  if (stroke) {
    const [r, g, b] = parseColor(stroke.color);
    lib.FPDFPageObj_SetStrokeColor(path, r, g, b, 255);
    lib.FPDFPageObj_SetStrokeWidth(path, stroke.width);
  }
  if (fill) {
    const [r, g, b] = parseColor(fill.color);
    lib.FPDFPageObj_SetFillColor(path, r, g, b, 255);
  }
  lib.FPDFPath_SetDrawMode(
    path,
    fill ? FILL_ALTERNATE : FILL_NONE,
    Boolean(stroke),
  );
}

function linePath(
  context: OperationContext,
  geometry: PageGeometry,
  from: PagePoint,
  to: PagePoint,
): number {
  const { lib } = context.pdfium;
  const start = pageToUser(geometry, from.x, from.y);
  const end = pageToUser(geometry, to.x, to.y);
  const path = lib.FPDFPageObj_CreateNewPath(start.x, start.y);
  lib.FPDFPath_LineTo(path, end.x, end.y);
  return path;
}

function rectanglePath(
  context: OperationContext,
  geometry: PageGeometry,
  rect: { x: number; y: number; width: number; height: number },
): number {
  const { lib } = context.pdfium;
  const corners = [
    pageToUser(geometry, rect.x, rect.y),
    pageToUser(geometry, rect.x + rect.width, rect.y),
    pageToUser(geometry, rect.x + rect.width, rect.y + rect.height),
    pageToUser(geometry, rect.x, rect.y + rect.height),
  ];
  const path = lib.FPDFPageObj_CreateNewPath(corners[0]!.x, corners[0]!.y);
  for (const corner of corners.slice(1))
    lib.FPDFPath_LineTo(path, corner.x, corner.y);
  lib.FPDFPath_Close(path);
  lib.FPDFPageObj_SetLineJoin(path, LINE_JOIN_ROUND);
  return path;
}

function ellipsePath(
  context: OperationContext,
  geometry: PageGeometry,
  rect: { x: number; y: number; width: number; height: number },
): number {
  const { lib } = context.pdfium;
  const cx = rect.x + rect.width / 2;
  const cy = rect.y + rect.height / 2;
  const rx = rect.width / 2;
  const ry = rect.height / 2;
  const at = (x: number, y: number) => pageToUser(geometry, x, y);
  // Four quarter arcs, clockwise in page space from the right-most point.
  const arcs: [PagePoint, PagePoint, PagePoint][] = [
    [
      { x: cx + rx, y: cy + ry * KAPPA },
      { x: cx + rx * KAPPA, y: cy + ry },
      { x: cx, y: cy + ry },
    ],
    [
      { x: cx - rx * KAPPA, y: cy + ry },
      { x: cx - rx, y: cy + ry * KAPPA },
      { x: cx - rx, y: cy },
    ],
    [
      { x: cx - rx, y: cy - ry * KAPPA },
      { x: cx - rx * KAPPA, y: cy - ry },
      { x: cx, y: cy - ry },
    ],
    [
      { x: cx + rx * KAPPA, y: cy - ry },
      { x: cx + rx, y: cy - ry * KAPPA },
      { x: cx + rx, y: cy },
    ],
  ];
  const start = at(cx + rx, cy);
  const path = lib.FPDFPageObj_CreateNewPath(start.x, start.y);
  for (const [c1, c2, end] of arcs) {
    const u1 = at(c1.x, c1.y);
    const u2 = at(c2.x, c2.y);
    const u3 = at(end.x, end.y);
    lib.FPDFPath_BezierTo(path, u1.x, u1.y, u2.x, u2.y, u3.x, u3.y);
  }
  lib.FPDFPath_Close(path);
  return path;
}

function result(location: ElementLocation): OperationResult {
  return { createdIds: [], changedPages: [location.pageIndex], warnings: [] };
}
