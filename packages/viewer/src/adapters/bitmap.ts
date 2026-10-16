import { ViewerError } from "../errors.js";

export async function drawEncodedImage(
  data: Uint8Array,
  mimeType: string,
  target: HTMLCanvasElement | OffscreenCanvas,
  options: {
    readonly dpr: number;
    readonly cssWidth?: number;
    readonly cssHeight?: number;
    readonly scale?: number;
  },
): Promise<{ width: number; height: number }> {
  const blob = new Blob([data.slice()], { type: mimeType });
  const decoded = await decodeImage(blob);
  try {
    const naturalWidth = decoded.width;
    const naturalHeight = decoded.height;
    const dpr = Math.max(1, options.dpr);
    const scale = options.scale ?? 1 / dpr;
    const cssWidth = options.cssWidth ?? naturalWidth * scale;
    const cssHeight = options.cssHeight ?? naturalHeight * scale;
    target.width = Math.max(1, Math.ceil(cssWidth * dpr));
    target.height = Math.max(1, Math.ceil(cssHeight * dpr));
    if ("style" in target) {
      target.style.width = `${cssWidth}px`;
      target.style.height = `${cssHeight}px`;
    }
    const context = target.getContext("2d");
    if (!context)
      throw new ViewerError(
        "render-failed",
        "Canvas 2D context is unavailable",
      );
    context.setTransform(dpr, 0, 0, dpr, 0, 0);
    context.clearRect(0, 0, cssWidth, cssHeight);
    context.drawImage(decoded.source, 0, 0, cssWidth, cssHeight);
    return { width: naturalWidth, height: naturalHeight };
  } finally {
    decoded.close();
  }
}

/**
 * Decode an encoded image only to read its natural size, in CSS pixels at
 * zoom 1. Uses the same decoder as {@link drawEncodedImage}, so EXIF
 * orientation is applied and the result matches what a render will draw.
 */
export async function measureEncodedImage(
  data: Uint8Array,
  mimeType: string,
): Promise<{ width: number; height: number }> {
  const decoded = await decodeImage(
    new Blob([data.slice()], { type: mimeType }),
  );
  try {
    return { width: decoded.width, height: decoded.height };
  } finally {
    decoded.close();
  }
}

async function decodeImage(blob: Blob): Promise<{
  readonly source: CanvasImageSource;
  readonly width: number;
  readonly height: number;
  readonly close: () => void;
}> {
  const isSvg = blob.type === "image/svg+xml";
  if (!isSvg && typeof createImageBitmap === "function") {
    const bitmap = await createImageBitmap(blob, {
      imageOrientation: "from-image",
    });
    return {
      source: bitmap,
      width: bitmap.width,
      height: bitmap.height,
      close: () => bitmap.close(),
    };
  }
  if (typeof document === "undefined")
    throw new ViewerError(
      "render-failed",
      "No browser image decoder is available",
    );
  const objectUrl = isSvg ? null : URL.createObjectURL(blob);
  try {
    const image = new Image();
    image.decoding = "async";
    image.src =
      objectUrl ??
      `data:image/svg+xml,${encodeURIComponent(await blob.text())}`;
    await image.decode();
    return {
      source: image,
      width: image.naturalWidth,
      height: image.naturalHeight,
      close: () => {
        if (objectUrl !== null) URL.revokeObjectURL(objectUrl);
      },
    };
  } catch (error) {
    if (objectUrl !== null) URL.revokeObjectURL(objectUrl);
    throw new ViewerError("invalid-file", "Browser image decode failed", {
      cause: error,
    });
  }
}
