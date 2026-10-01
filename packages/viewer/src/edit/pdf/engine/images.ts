import { ViewerError } from "../../../errors.js";
import { decodeBinary } from "../../operations.js";
import type { EditOperation } from "../../types.js";
import type { InsertImageOperation, PdfOperation } from "../types.js";
import { OBJECT_IMAGE } from "./elements.js";
import { pageToUser } from "./geometry.js";
import type { OperationHandler } from "./operations.js";
import { validateRect } from "./text-box.js";

/*
 * Raster images. JPEG data is handed to PDFium as it is, so the file keeps
 * the original stream; PNG data is decoded by the host (createImageBitmap in
 * the worker) into pixels PDFium stores losslessly, alpha included.
 */

/** FPDFBitmap_CreateEx pixel format with alpha. */
const BITMAP_BGRA = 4;

export interface DecodedImage {
  readonly width: number;
  readonly height: number;
  /** Row-major RGBA, 4 bytes per pixel. */
  readonly rgba: Uint8Array;
}

export type ImageDecoder = (
  bytes: Uint8Array,
  mimeType: "image/png" | "image/jpeg",
) => Promise<DecodedImage>;

/**
 * Decoded pixels for the PNG operations of a batch, keyed by content, so the
 * synchronous apply never decodes. Only the last few images are kept.
 */
export class ImageCache {
  readonly #decoded = new Map<string, DecodedImage>();

  /** Decodes every PNG the batch inserts; must run before validate or apply. */
  async prepare(
    operations: readonly EditOperation[],
    decode: ImageDecoder,
  ): Promise<void> {
    for (const raw of operations) {
      const operation = raw as PdfOperation;
      if (operation.op !== "insertImage" || operation.mimeType !== "image/png")
        continue;
      const bytes = decodeBinary(operation.data);
      if (!isPng(bytes)) continue;
      const key = contentKey(bytes);
      if (this.#decoded.has(key)) continue;
      this.#decoded.set(key, await decode(bytes, "image/png"));
      while (this.#decoded.size > 8)
        this.#decoded.delete(this.#decoded.keys().next().value!);
    }
  }

  get(bytes: Uint8Array): DecodedImage | undefined {
    return this.#decoded.get(contentKey(bytes));
  }
}

export const insertImage: OperationHandler<InsertImageOperation> = {
  validate(operation, context, issue) {
    if (operation.pageIndex >= context.pageCount) {
      issue("/pageIndex", "unknown-target", `No page ${operation.pageIndex}`);
      return;
    }
    validateRect(operation.rect, context.geometry(operation.pageIndex), issue);
    const bytes = decodeBinary(operation.data);
    if (bytes.byteLength > context.limits.maxInputBytes)
      throw new ViewerError(
        "resource-limit",
        "Image data exceeds maxInputBytes",
        {
          details: {
            bytes: bytes.byteLength,
            limit: context.limits.maxInputBytes,
          },
        },
      );
    const size =
      operation.mimeType === "image/png" ? pngSize(bytes) : jpegSize(bytes);
    if (!size) {
      issue(
        "/data",
        "invalid-data",
        `The data is not a readable ${operation.mimeType} image`,
      );
      return;
    }
    const pixels = size.width * size.height;
    if (
      !Number.isSafeInteger(pixels) ||
      pixels > context.limits.maxDecodedPixels
    )
      throw new ViewerError(
        "resource-limit",
        "Decoded image exceeds maxDecodedPixels",
        { details: { pixels, limit: context.limits.maxDecodedPixels } },
      );
    if (operation.mimeType === "image/png" && !context.images.get(bytes))
      issue("/data", "invalid-data", "The PNG could not be decoded");
  },

  apply(operation, context) {
    const { pdfium } = context;
    const { lib } = pdfium;
    const bytes = decodeBinary(operation.data);
    const id = context.newId(operation.pageIndex);
    const geometry = context.geometry(operation.pageIndex);
    const image = lib.FPDFPageObj_NewImageObj(context.document);
    if (operation.mimeType === "image/jpeg") {
      const loaded = pdfium.withFileAccess(bytes, (fileAccess) =>
        lib.FPDFImageObj_LoadJpegFileInline(0, 0, image, fileAccess),
      );
      if (!loaded)
        throw new ViewerError("edit-failed", "PDFium could not read the JPEG", {
          details: { stage: "apply" },
        });
    } else {
      const decoded = context.images.get(bytes)!;
      setBitmap(context.pdfium, image, decoded);
    }
    // The unit square maps onto the page-space rectangle, upright on screen:
    // its bottom-left and the two edges leaving it, in user space.
    const { rect } = operation;
    const origin = pageToUser(geometry, rect.x, rect.y + rect.height);
    const right = pageToUser(
      geometry,
      rect.x + rect.width,
      rect.y + rect.height,
    );
    const up = pageToUser(geometry, rect.x, rect.y);
    lib.FPDFImageObj_SetMatrix(
      image,
      right.x - origin.x,
      right.y - origin.y,
      up.x - origin.x,
      up.y - origin.y,
      origin.x,
      origin.y,
    );
    context.withPage(operation.pageIndex, (page) => {
      lib.FPDFPage_InsertObject(page, image);
    });
    context.appendObjects(operation.pageIndex, [{ id, type: OBJECT_IMAGE }]);
    return {
      createdIds: [id],
      changedPages: [operation.pageIndex],
      warnings: [],
    };
  },
};

function setBitmap(
  pdfium: import("./pdfium.js").Pdfium,
  image: number,
  decoded: DecodedImage,
): void {
  const { lib } = pdfium;
  const stride = decoded.width * 4;
  const bgra = new Uint8Array(decoded.rgba.byteLength);
  for (let offset = 0; offset < bgra.byteLength; offset += 4) {
    bgra[offset] = decoded.rgba[offset + 2]!;
    bgra[offset + 1] = decoded.rgba[offset + 1]!;
    bgra[offset + 2] = decoded.rgba[offset]!;
    bgra[offset + 3] = decoded.rgba[offset + 3]!;
  }
  const buffer = pdfium.writeBytes(bgra);
  try {
    const bitmap = lib.FPDFBitmap_CreateEx(
      decoded.width,
      decoded.height,
      BITMAP_BGRA,
      buffer,
      stride,
    );
    if (!bitmap)
      throw new ViewerError(
        "edit-failed",
        "PDFium could not create the bitmap",
        {
          details: { stage: "apply" },
        },
      );
    try {
      if (!lib.FPDFImageObj_SetBitmap(0, 0, image, bitmap))
        throw new ViewerError(
          "edit-failed",
          "PDFium could not store the image",
          {
            details: { stage: "apply" },
          },
        );
    } finally {
      lib.FPDFBitmap_Destroy(bitmap);
    }
  } finally {
    pdfium.free(buffer);
  }
}

export function isPng(bytes: Uint8Array): boolean {
  return (
    bytes.length > 24 &&
    bytes[0] === 0x89 &&
    bytes[1] === 0x50 &&
    bytes[2] === 0x4e &&
    bytes[3] === 0x47
  );
}

/** Width and height from the IHDR chunk. */
export function pngSize(
  bytes: Uint8Array,
): { width: number; height: number } | undefined {
  if (!isPng(bytes)) return undefined;
  const view = new DataView(bytes.buffer, bytes.byteOffset, bytes.byteLength);
  const width = view.getUint32(16);
  const height = view.getUint32(20);
  return width > 0 && height > 0 ? { width, height } : undefined;
}

/** Width and height from the first SOF marker of a baseline or progressive JPEG. */
export function jpegSize(
  bytes: Uint8Array,
): { width: number; height: number } | undefined {
  if (bytes.length < 4 || bytes[0] !== 0xff || bytes[1] !== 0xd8)
    return undefined;
  const view = new DataView(bytes.buffer, bytes.byteOffset, bytes.byteLength);
  let offset = 2;
  while (offset + 9 < bytes.length) {
    if (bytes[offset] !== 0xff) return undefined;
    const marker = bytes[offset + 1]!;
    if (
      marker === 0xd8 ||
      (marker >= 0xd0 && marker <= 0xd7) ||
      marker === 0x01
    ) {
      offset += 2;
      continue;
    }
    const length = view.getUint16(offset + 2);
    if (
      marker >= 0xc0 &&
      marker <= 0xcf &&
      marker !== 0xc4 &&
      marker !== 0xc8 &&
      marker !== 0xcc
    ) {
      const height = view.getUint16(offset + 5);
      const width = view.getUint16(offset + 7);
      return width > 0 && height > 0 ? { width, height } : undefined;
    }
    if (marker === 0xda) return undefined;
    offset += 2 + length;
  }
  return undefined;
}

/** Cheap content key: length plus an FNV-1a hash of the bytes. */
function contentKey(bytes: Uint8Array): string {
  let hash = 0x811c9dc5;
  for (const byte of bytes) hash = Math.imul(hash ^ byte, 0x01000193) >>> 0;
  return `${bytes.byteLength}:${hash.toString(16)}`;
}
