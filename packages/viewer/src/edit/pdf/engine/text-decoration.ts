import type { Pdfium } from "./pdfium.js";

/** A native underline follows the actual glyph origins, without re-laying out text. */
interface UnderlineGeometry {
  readonly matrix: readonly number[];
  readonly rect: {
    readonly x: number;
    readonly y: number;
    readonly width: number;
    readonly height: number;
  };
  readonly color: readonly number[];
}

const UNDERLINE_OFFSET_EM = 0.12;
const UNDERLINE_THICKNESS_EM = 0.05;
const NATIVE_PRECISION = 0.02;

export function underlineGeometry(
  pdfium: Pdfium,
  object: number,
  textPage: number,
): UnderlineGeometry | undefined {
  const { lib } = pdfium;
  const matrix = pdfium.readNumbers(
    6,
    "float",
    ([p]) => p !== undefined && lib.FPDFPageObj_GetMatrix(object, p),
  );
  const size = pdfium.readNumbers(
    1,
    "float",
    ([p]) => p !== undefined && lib.FPDFTextObj_GetFontSize(object, p),
  )?.[0];
  const color = pdfium.readNumbers(
    4,
    "i32",
    ([r, g, b, a]) =>
      r !== undefined &&
      g !== undefined &&
      b !== undefined &&
      a !== undefined &&
      lib.FPDFPageObj_GetFillColor(object, r, g, b, a),
  );
  if (!matrix || !size || !color) return undefined;
  const [a = 0, b = 0, c = 0, d = 0, e = 0, f = 0] = matrix;
  // PDFium exposes axis-aligned advance boxes. Inverting their enclosing
  // rectangles is faithful only for axis-aligned or quarter-turned text.
  const nearZero = (value: number) => Math.abs(value) < 1e-6;
  if (!(nearZero(b) && nearZero(c)) && !(nearZero(a) && nearZero(d)))
    return undefined;
  const determinant = a * d - b * c;
  if (Math.abs(determinant) < 1e-9) return undefined;
  const local = (x: number, y: number) => ({
    x: (d * (x - e) - c * (y - f)) / determinant,
    y: (-b * (x - e) + a * (y - f)) / determinant,
  });
  let left = Infinity;
  let right = -Infinity;
  let baseline: number | undefined;
  const count = lib.FPDFText_CountChars(textPage);
  for (let index = 0; index < count; index += 1) {
    if (
      lib.FPDFText_GetTextObject(textPage, index) !== object ||
      lib.FPDFText_IsGenerated(textPage, index) === 1
    )
      continue;
    const bounds = pdfium.readNumbers(
      4,
      "float",
      ([p]) =>
        p !== undefined && lib.FPDFText_GetLooseCharBox(textPage, index, p),
    );
    if (!bounds) continue;
    const [l = 0, t = 0, r = 0, bottom = 0] = bounds;
    for (const [x, y] of [
      [l, t],
      [l, bottom],
      [r, t],
      [r, bottom],
    ]) {
      if (x === undefined || y === undefined) continue;
      const point = local(x, y);
      left = Math.min(left, point.x);
      right = Math.max(right, point.x);
    }
    if (baseline === undefined) {
      const origin = pdfium.readNumbers(
        2,
        "double",
        ([x, y]) =>
          x !== undefined &&
          y !== undefined &&
          lib.FPDFText_GetCharOrigin(textPage, index, x, y),
      );
      if (origin?.[0] !== undefined && origin[1] !== undefined)
        baseline = local(origin[0], origin[1]).y;
    }
  }
  if (baseline === undefined || !Number.isFinite(left) || right <= left)
    return undefined;
  return {
    matrix,
    color,
    rect: {
      x: left,
      y: baseline - size * (UNDERLINE_OFFSET_EM + UNDERLINE_THICKNESS_EM / 2),
      width: right - left,
      height: size * UNDERLINE_THICKNESS_EM,
    },
  };
}

export function createUnderline(
  pdfium: Pdfium,
  object: number,
  textPage: number,
): number | undefined {
  const geometry = underlineGeometry(pdfium, object, textPage);
  if (!geometry) return undefined;
  const { lib } = pdfium;
  const { rect, matrix, color } = geometry;
  const path = lib.FPDFPageObj_CreateNewRect(
    rect.x,
    rect.y,
    rect.width,
    rect.height,
  );
  if (!path) return undefined;
  const [a = 0, b = 0, c = 0, d = 0, e = 0, f = 0] = matrix;
  const [r = 0, g = 0, blue = 0, alpha = 255] = color;
  lib.FPDFPageObj_Transform(path, a, b, c, d, e, f);
  lib.FPDFPageObj_SetFillColor(path, r, g, blue, alpha);
  lib.FPDFPath_SetDrawMode(path, 1, false);
  return path;
}

/** Validate native path points as well as appearance before trusting a saved ownership mark. */
export function isUnderlinePath(
  pdfium: Pdfium,
  text: number,
  path: number,
  textPage: number,
): boolean {
  const expected = underlineGeometry(pdfium, text, textPage);
  if (!expected) return false;
  const { lib } = pdfium;
  const mode = pdfium.readNumbers(
    2,
    "i32",
    ([fill, stroke]) =>
      fill !== undefined &&
      stroke !== undefined &&
      lib.FPDFPath_GetDrawMode(path, fill, stroke),
  );
  const color = pdfium.readNumbers(
    4,
    "i32",
    ([r, g, b, a]) =>
      r !== undefined &&
      g !== undefined &&
      b !== undefined &&
      a !== undefined &&
      lib.FPDFPageObj_GetFillColor(path, r, g, b, a),
  );
  const matrix = pdfium.readNumbers(
    6,
    "float",
    ([p]) => p !== undefined && lib.FPDFPageObj_GetMatrix(path, p),
  );
  if (
    mode?.[0] !== 1 ||
    mode[1] !== 0 ||
    !color ||
    !matrix ||
    color.some((value, index) => value !== expected.color[index]) ||
    lib.FPDFPath_CountSegments(path) !== 5
  )
    return false;
  const transform = (values: readonly number[], x: number, y: number) => ({
    x: (values[0] ?? 0) * x + (values[2] ?? 0) * y + (values[4] ?? 0),
    y: (values[1] ?? 0) * x + (values[3] ?? 0) * y + (values[5] ?? 0),
  });
  const { rect } = expected;
  const corners = [
    [rect.x, rect.y],
    [rect.x + rect.width, rect.y],
    [rect.x + rect.width, rect.y + rect.height],
    [rect.x, rect.y + rect.height],
  ];
  const visited = new Set<number>();
  let firstCorner: number | undefined;
  for (let index = 0; index < 5; index += 1) {
    const segment = lib.FPDFPath_GetPathSegment(path, index);
    if (lib.FPDFPathSegment_GetType(segment) === 1) return false; // Bezier curve
    const point = pdfium.readNumbers(
      2,
      "float",
      ([x, y]) =>
        x !== undefined &&
        y !== undefined &&
        lib.FPDFPathSegment_GetPoint(segment, x, y),
    );
    if (point?.[0] === undefined || point[1] === undefined) return false;
    const actual = transform(matrix, point[0], point[1]);
    const cornerIndex = corners.findIndex(([x = 0, y = 0]) => {
      const corner = transform(expected.matrix, x, y);
      return (
        Math.abs(actual.x - corner.x) < NATIVE_PRECISION &&
        Math.abs(actual.y - corner.y) < NATIVE_PRECISION
      );
    });
    if (cornerIndex < 0) return false;
    if (index === 0) firstCorner = cornerIndex;
    if (index === 4) return visited.size === 4 && cornerIndex === firstCorner;
    if (visited.has(cornerIndex)) return false;
    visited.add(cornerIndex);
  }
  return false;
}
