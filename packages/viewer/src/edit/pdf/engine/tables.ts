import type { PagePoint } from "../../types.js";
import type {
  InsertTableOperation,
  PdfTableStyle,
  SetTableCellOperation,
} from "../types.js";
import {
  MARK_NAME,
  MARK_PARAM,
  OBJECT_PATH,
  OBJECT_TEXT,
  type MarkParams,
  type ObjectRecord,
} from "./elements.js";
import type { FontRequest } from "./fonts.js";
import { displayedSize, pageToUser } from "./geometry.js";
import type {
  ElementLocation,
  Issue,
  OperationContext,
  OperationHandler,
  OperationResult,
} from "./operations.js";
import { layoutText, type LayoutLine } from "./text-layout.js";
import {
  parseColor,
  placeUpright,
  removeObjects,
  setText,
  validateRect,
  type DrawnTextBox,
} from "./text-box.js";

/*
 * Tables are parametric like text boxes: the grid is one stroked path, an
 * optional header fill one more, and every cell line a text object. The
 * inputs live in the `WebDoc` mark of the path objects; the text objects
 * carry only the table's id, so a large table does not repeat its contents.
 */

export const MAX_ROWS = 100;
export const MAX_COLUMNS = 20;

/** FPDFPath_SetDrawMode fill modes. */
const FILL_NONE = 0;
const FILL_ALTERNATE = 1;
const LINE_HEIGHT = 1.2;

/** A table's inputs with every default filled in; stored in its head mark. */
export type TableSpec = {
  readonly kind: "table";
  readonly id: string;
  readonly at: PagePoint;
  readonly rows: readonly (readonly string[])[];
  /** Absolute column widths in points. */
  readonly columnWidths: readonly number[];
  readonly style: {
    readonly fontFamily: string;
    readonly fontSize: number;
    readonly color: string;
    readonly borderColor: string;
    readonly borderWidth: number;
    readonly cellPadding: number;
    readonly headerFill: string | null;
  };
};

export const insertTable: OperationHandler<InsertTableOperation> = {
  validate(operation, context, issue) {
    if (operation.pageIndex >= context.pageCount) {
      issue("/pageIndex", "unknown-target", `No page ${operation.pageIndex}`);
      return;
    }
    if (!validateRows(operation.rows, issue)) return;
    const columns = operation.rows[0]!.length;
    if (
      operation.columnWidths !== undefined &&
      operation.columnWidths.length !== columns
    ) {
      issue(
        "/columnWidths",
        "range",
        `Give one width per column; the table has ${columns}`,
      );
      return;
    }
    const style = resolveTableStyle(operation.style ?? {});
    const widths = resolveColumns(
      operation.width,
      columns,
      operation.columnWidths,
    );
    validateRect(
      {
        x: operation.at.x,
        y: operation.at.y,
        width: operation.width,
        height: 0.001,
      },
      context.geometry(operation.pageIndex),
      issue,
      "/at",
    );
    if (Math.min(...widths) - 2 * style.cellPadding < 1)
      issue(
        "/width",
        "range",
        "Every column must be wider than twice the cell padding",
      );
    validateTableFont(
      style,
      tableText(operation.rows),
      context,
      issue,
      "/rows",
    );
  },

  apply(operation, context) {
    const id = context.newId(operation.pageIndex);
    const columns = operation.rows[0]!.length;
    const spec: TableSpec = {
      kind: "table",
      id,
      at: operation.at,
      rows: operation.rows,
      columnWidths: resolveColumns(
        operation.width,
        columns,
        operation.columnWidths,
      ),
      style: resolveTableStyle(operation.style ?? {}),
    };
    const drawn = drawTable(context, operation.pageIndex, spec);
    return {
      createdIds: [id],
      changedPages: [operation.pageIndex],
      warnings: tableWarnings(drawn, id),
    };
  },
};

export const setTableCell: OperationHandler<SetTableCellOperation> = {
  validate(operation, context, issue) {
    const target = tableTarget(operation.target, context, issue);
    if (!target) return;
    const { spec } = target;
    if (operation.row >= spec.rows.length) {
      issue("/row", "range", `The table has ${spec.rows.length} rows`);
      return;
    }
    const columns = spec.rows[0]!.length;
    if (operation.column >= columns) {
      issue("/column", "range", `The table has ${columns} columns`);
      return;
    }
    validateTableFont(spec.style, operation.text, context, issue, "/text");
  },
  apply(operation, context) {
    const { location, spec } = tableTarget(operation.target, context)!;
    const rows = spec.rows.map((row, rowIndex) =>
      rowIndex === operation.row
        ? row.map((cell, columnIndex) =>
            columnIndex === operation.column ? operation.text : cell,
          )
        : row,
    );
    return changedTable(
      location,
      rebuildTable(context, location, { ...spec, rows }),
    );
  },
};

/** Lays the table out and inserts its objects, appended or at `insertAt`. */
export function drawTable(
  context: OperationContext,
  pageIndex: number,
  spec: TableSpec,
  insertAt?: number,
): DrawnTextBox {
  const { pdfium, measurer } = context;
  const { lib } = pdfium;
  const { style } = spec;
  const font = context.fonts.resolve(
    pdfium,
    context.document,
    tableFontRequest(style, tableText(spec.rows)),
  );
  const { ascent } = measurer.metrics(font.handle, style.fontSize);
  const layout = layoutTable(spec, ascent, (text) =>
    measurer.advance(font.handle, style.fontSize, text),
  );
  const geometry = context.geometry(pageIndex);
  const overflow =
    spec.at.y + layout.height > displayedSize(geometry).height + 0.001;
  const head = JSON.stringify(spec);
  const memberMark: MarkParams = { kind: "table", id: spec.id };
  const member = JSON.stringify(memberMark);
  const at = (x: number, y: number) => pageToUser(geometry, x, y);
  const width = spec.columnWidths.reduce((total, w) => total + w, 0);
  const right = spec.at.x + width;
  const bottom = spec.at.y + layout.height;
  const [textR, textG, textB] = parseColor(style.color);

  context.withPage(pageIndex, (page) => {
    const records: ObjectRecord[] = [];
    let index = insertAt ?? -1;
    const place = (
      object: number,
      type: number,
      mark: MarkParams,
      params: string,
    ): void => {
      const handle = lib.FPDFPageObj_AddMark(object, MARK_NAME);
      lib.FPDFPageObjMark_SetStringParam(
        context.document,
        object,
        handle,
        MARK_PARAM,
        params,
      );
      if (insertAt === undefined) lib.FPDFPage_InsertObject(page, object);
      else lib.FPDFPage_InsertObjectAtIndex(page, object, index++);
      records.push({ id: spec.id, type, mark });
    };

    if (style.headerFill) {
      const headerBottom = spec.at.y + layout.rows[0]!.height;
      const corners = [
        at(spec.at.x, spec.at.y),
        at(right, spec.at.y),
        at(right, headerBottom),
        at(spec.at.x, headerBottom),
      ];
      const fill = lib.FPDFPageObj_CreateNewPath(corners[0]!.x, corners[0]!.y);
      for (const corner of corners.slice(1))
        lib.FPDFPath_LineTo(fill, corner.x, corner.y);
      lib.FPDFPath_Close(fill);
      const [r, g, b] = parseColor(style.headerFill);
      lib.FPDFPageObj_SetFillColor(fill, r, g, b, 255);
      lib.FPDFPath_SetDrawMode(fill, FILL_ALTERNATE, false);
      place(fill, OBJECT_PATH, spec, head);
    }

    // One path for the whole grid: the frame, then the inner lines.
    const frame = [
      at(spec.at.x, spec.at.y),
      at(right, spec.at.y),
      at(right, bottom),
      at(spec.at.x, bottom),
    ];
    const grid = lib.FPDFPageObj_CreateNewPath(frame[0]!.x, frame[0]!.y);
    for (const corner of frame.slice(1))
      lib.FPDFPath_LineTo(grid, corner.x, corner.y);
    lib.FPDFPath_Close(grid);
    for (const column of layout.columns.slice(1)) {
      const top = at(column.x, spec.at.y);
      const end = at(column.x, bottom);
      lib.FPDFPath_MoveTo(grid, top.x, top.y);
      lib.FPDFPath_LineTo(grid, end.x, end.y);
    }
    for (const row of layout.rows.slice(1)) {
      const start = at(spec.at.x, row.y);
      const end = at(right, row.y);
      lib.FPDFPath_MoveTo(grid, start.x, start.y);
      lib.FPDFPath_LineTo(grid, end.x, end.y);
    }
    const [r, g, b] = parseColor(style.borderColor);
    lib.FPDFPageObj_SetStrokeColor(grid, r, g, b, 255);
    lib.FPDFPageObj_SetStrokeWidth(grid, style.borderWidth);
    lib.FPDFPath_SetDrawMode(grid, FILL_NONE, true);
    place(grid, OBJECT_PATH, spec, head);

    layout.rows.forEach((row) => {
      row.cells.forEach((lines, columnIndex) => {
        const column = layout.columns[columnIndex]!;
        for (const line of lines) {
          if (!line.text) continue;
          const object = lib.FPDFPageObj_CreateTextObj(
            context.document,
            font.handle,
            style.fontSize,
          );
          setText(pdfium, object, line.text);
          lib.FPDFPageObj_SetFillColor(object, textR, textG, textB, 255);
          placeUpright(
            pdfium,
            object,
            geometry,
            column.x + style.cellPadding + line.x,
            row.y + style.cellPadding + line.baseline,
          );
          place(object, OBJECT_TEXT, memberMark, member);
        }
      });
    });

    if (insertAt === undefined) context.appendObjects(pageIndex, records);
    else context.spliceObjects(pageIndex, insertAt, 0, records);
  });
  return {
    overflow,
    ...(font.substitution ? { substitution: font.substitution } : {}),
  };
}

/** Redraws a table from `spec` at the same place in the drawing order. */
export function rebuildTable(
  context: OperationContext,
  location: ElementLocation,
  spec: TableSpec,
): DrawnTextBox {
  const first = removeObjects(context, location);
  return drawTable(context, location.pageIndex, spec, first);
}

/** The location and stored inputs of a table, or the issue that stops the edit. */
export function tableTarget(
  target: string,
  context: OperationContext,
  issue: Issue = () => {},
):
  { readonly location: ElementLocation; readonly spec: TableSpec } | undefined {
  const location = context.locate(target);
  if (!location) {
    issue("/target", "unknown-target", `No element ${target}`);
    return undefined;
  }
  const spec = tableSpecOf(location.record.mark);
  if (!spec) {
    issue(
      "/target",
      "unsupported-target",
      "Only tables created by web-doc have cells",
    );
    return undefined;
  }
  return { location, spec };
}

/** The spec a mark carries when it is a table's head mark. */
export function tableSpecOf(
  mark: MarkParams | undefined,
): TableSpec | undefined {
  return mark?.kind === "table" && Array.isArray(mark.rows)
    ? (mark as unknown as TableSpec)
    : undefined;
}

export function changedTable(
  location: ElementLocation,
  drawn: DrawnTextBox,
): OperationResult {
  return {
    createdIds: [],
    changedPages: [location.pageIndex],
    warnings: tableWarnings(drawn, location.record.id),
  };
}

/** Every cell's text, for font coverage checks. */
export function tableText(rows: readonly (readonly string[])[]): string {
  return rows.flat().join("\n");
}

export function tableFontRequest(
  style: TableSpec["style"],
  text: string,
): FontRequest {
  return { family: style.fontFamily, bold: false, italic: false, text };
}

export function resolveTableStyle(style: PdfTableStyle): TableSpec["style"] {
  return {
    fontFamily: style.fontFamily ?? "Helvetica",
    fontSize: style.fontSize ?? 10,
    color: style.color ?? "#000000",
    borderColor: style.borderColor ?? "#000000",
    borderWidth: style.borderWidth ?? 0.75,
    cellPadding: style.cellPadding ?? 4,
    headerFill: style.headerFill ?? null,
  };
}

interface TableLayout {
  readonly columns: readonly { readonly x: number; readonly width: number }[];
  readonly rows: readonly {
    readonly y: number;
    readonly height: number;
    readonly cells: readonly (readonly LayoutLine[])[];
  }[];
  readonly height: number;
}

/** Column edges and row heights: each row is as tall as its tallest cell. */
function layoutTable(
  spec: TableSpec,
  ascent: number,
  advance: (text: string) => number,
): TableLayout {
  const { cellPadding, fontSize } = spec.style;
  const columns: { x: number; width: number }[] = [];
  let x = spec.at.x;
  for (const width of spec.columnWidths) {
    columns.push({ x, width });
    x += width;
  }
  const rows: TableLayout["rows"][number][] = [];
  let y = spec.at.y;
  for (const cells of spec.rows) {
    const layouts = cells.map((text, columnIndex) =>
      layoutText({
        text,
        width: columns[columnIndex]!.width - 2 * cellPadding,
        height: Number.POSITIVE_INFINITY,
        fontSize,
        lineHeight: LINE_HEIGHT,
        align: "left",
        ascent,
        advance,
      }),
    );
    const height =
      Math.max(...layouts.map((layout) => layout.height)) + 2 * cellPadding;
    rows.push({ y, height, cells: layouts.map((layout) => layout.lines) });
    y += height;
  }
  return { columns, rows, height: y - spec.at.y };
}

/** Column widths in points from optional relative weights. */
function resolveColumns(
  width: number,
  columns: number,
  weights: readonly number[] | undefined,
): number[] {
  const relative = weights ?? Array.from({ length: columns }, () => 1);
  const total = relative.reduce((sum, weight) => sum + weight, 0);
  return relative.map((weight) => (weight / total) * width);
}

function validateRows(
  rows: readonly (readonly string[])[],
  issue: Issue,
): boolean {
  if (rows.length === 0 || rows.length > MAX_ROWS) {
    issue("/rows", "range", `A table has 1 to ${MAX_ROWS} rows`);
    return false;
  }
  const columns = rows[0]!.length;
  if (columns === 0 || columns > MAX_COLUMNS) {
    issue("/rows/0", "range", `A table has 1 to ${MAX_COLUMNS} columns`);
    return false;
  }
  let ok = true;
  rows.forEach((row, index) => {
    if (row.length !== columns) {
      issue(
        `/rows/${index}`,
        "range",
        `Row ${index} has ${row.length} cells; the first row has ${columns}`,
      );
      ok = false;
    }
  });
  return ok;
}

function validateTableFont(
  style: TableSpec["style"],
  text: string,
  context: OperationContext,
  issue: Issue,
  textPath: string,
): void {
  const problem = context.fonts.problem(tableFontRequest(style, text));
  if (problem)
    issue(
      problem.path === "/text" ? textPath : problem.path,
      problem.code,
      problem.message,
    );
}

function tableWarnings(
  drawn: DrawnTextBox,
  elementId: string,
): OperationResult["warnings"] {
  return [
    ...(drawn.overflow
      ? [
          {
            code: "fidelity-degraded" as const,
            message: "The table runs past the bottom of the page",
            details: { elementId },
          },
        ]
      : []),
    ...(drawn.substitution
      ? [
          {
            code: "font-substitution" as const,
            message: drawn.substitution,
            details: { elementId },
          },
        ]
      : []),
  ];
}
