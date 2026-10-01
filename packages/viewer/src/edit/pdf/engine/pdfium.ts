import { ViewerError } from "../../../errors.js";

/**
 * The PDFium entry points the PDF edit engine calls, typed from
 * `@embedpdf/pdfium`'s `functions` table. Pointers and handles are plain
 * numbers in the WASM address space.
 */
export interface PdfiumFunctions {
  PDFiumExt_Init(): void;
  PDFiumExt_OpenFileWriter(): number;
  PDFiumExt_GetFileWriterSize(writer: number): number;
  PDFiumExt_GetFileWriterData(
    writer: number,
    buffer: number,
    size: number,
  ): number;
  PDFiumExt_CloseFileWriter(writer: number): void;
  FPDF_GetLastError(): number;
  FPDF_CreateNewDocument(): number;
  FPDF_LoadMemDocument(data: number, size: number, password: string): number;
  FPDF_CloseDocument(document: number): void;
  FPDF_SaveAsCopy(document: number, writer: number, flags: number): boolean;
  FPDF_GetPageCount(document: number): number;
  FPDF_GetSignatureCount(document: number): number;
  FPDF_LoadPage(document: number, pageIndex: number): number;
  FPDF_ClosePage(page: number): void;
  FPDF_GetPageWidthF(page: number): number;
  FPDF_GetPageHeightF(page: number): number;
  /** Media box ∩ crop box in user space, as an FS_RECTF {left, top, right, bottom}. */
  FPDF_GetPageBoundingBox(page: number, rect: number): boolean;
  FPDF_PageToDevice(
    page: number,
    startX: number,
    startY: number,
    sizeX: number,
    sizeY: number,
    rotate: number,
    pageX: number,
    pageY: number,
    deviceX: number,
    deviceY: number,
  ): boolean;
  FPDFPage_SetCropBox(
    page: number,
    left: number,
    bottom: number,
    right: number,
    top: number,
  ): void;
  FPDF_MovePages(
    document: number,
    pageIndices: number,
    count: number,
    destination: number,
  ): boolean;
  FPDFPage_New(
    document: number,
    pageIndex: number,
    width: number,
    height: number,
  ): number;
  FPDFPage_Delete(document: number, pageIndex: number): void;
  FPDFPage_GetRotation(page: number): number;
  FPDFPage_SetRotation(page: number, rotation: number): void;
  FPDFPage_CountObjects(page: number): number;
  FPDFPage_GetObject(page: number, index: number): number;
  FPDFPage_InsertObject(page: number, object: number): void;
  FPDFPage_InsertObjectAtIndex(
    page: number,
    object: number,
    index: number,
  ): boolean;
  FPDFPage_GenerateContent(page: number): boolean;
  FPDFPageObj_GetType(object: number): number;
  /** Four floats: left, bottom, right, top in user space. */
  FPDFPageObj_GetBounds(
    object: number,
    left: number,
    bottom: number,
    right: number,
    top: number,
  ): boolean;
  /** Six floats a, b, c, d, e, f. */
  FPDFPageObj_GetMatrix(object: number, matrix: number): boolean;
  FPDFPageObj_GetFillColor(
    object: number,
    r: number,
    g: number,
    b: number,
    a: number,
  ): boolean;
  FPDFPageObj_GetStrokeColor(
    object: number,
    r: number,
    g: number,
    b: number,
    a: number,
  ): boolean;
  FPDFPageObj_GetStrokeWidth(object: number, width: number): boolean;
  FPDFPageObj_SetFillColor(
    object: number,
    r: number,
    g: number,
    b: number,
    a: number,
  ): boolean;
  FPDFPageObj_SetStrokeColor(
    object: number,
    r: number,
    g: number,
    b: number,
    a: number,
  ): boolean;
  FPDFPageObj_SetStrokeWidth(object: number, width: number): boolean;
  FPDFPageObj_CreateNewRect(
    x: number,
    y: number,
    width: number,
    height: number,
  ): number;
  FPDFPath_GetDrawMode(path: number, fillMode: number, stroke: number): boolean;
  FPDFPageObj_CreateNewPath(x: number, y: number): number;
  FPDFPath_MoveTo(path: number, x: number, y: number): boolean;
  FPDFPath_LineTo(path: number, x: number, y: number): boolean;
  FPDFPath_BezierTo(
    path: number,
    x1: number,
    y1: number,
    x2: number,
    y2: number,
    x3: number,
    y3: number,
  ): boolean;
  FPDFPath_Close(path: number): boolean;
  FPDFPageObj_SetLineJoin(object: number, join: number): boolean;
  FPDFPath_SetDrawMode(
    path: number,
    fillMode: number,
    stroke: boolean,
  ): boolean;
  FPDFTextObj_GetText(
    textObject: number,
    textPage: number,
    buffer: number,
    bytes: number,
  ): number;
  FPDFTextObj_GetFontSize(textObject: number, size: number): boolean;
  FPDFTextObj_GetFont(textObject: number): number;
  FPDFFont_GetFamilyName(font: number, buffer: number, bytes: number): number;
  FPDFFont_GetBaseFontName(font: number, buffer: number, bytes: number): number;
  FPDFFont_GetFlags(font: number): number;
  FPDFFont_GetIsEmbedded(font: number): number;
  FPDFFont_GetAscent(font: number, fontSize: number, ascent: number): boolean;
  FPDFFont_GetDescent(font: number, fontSize: number, descent: number): boolean;
  FPDFPageObj_Destroy(object: number): void;
  FPDFPage_RemoveObject(page: number, object: number): boolean;
  FPDFFont_GetWeight(font: number): number;
  FPDFText_GetCharBox(
    textPage: number,
    index: number,
    left: number,
    right: number,
    bottom: number,
    top: number,
  ): boolean;
  FPDFText_GetTextObject(textPage: number, index: number): number;
  FPDFText_GetUnicode(textPage: number, index: number): number;
  FPDFText_GetCharIndexFromTextIndex(
    textPage: number,
    textIndex: number,
  ): number;
  FPDFPageObjMark_CountParams(mark: number): number;
  FPDFPageObj_Transform(
    object: number,
    a: number,
    b: number,
    c: number,
    d: number,
    e: number,
    f: number,
  ): void;
  FPDFPageObj_CreateTextObj(
    document: number,
    font: number,
    fontSize: number,
  ): number;
  FPDFPageObj_NewImageObj(document: number): number;
  FPDFPageObj_AddMark(object: number, name: string): number;
  FPDFPageObj_CountMarks(object: number): number;
  FPDFPageObj_GetMark(object: number, index: number): number;
  FPDFPageObjMark_GetName(
    mark: number,
    buffer: number,
    bufferBytes: number,
    outBytes: number,
  ): boolean;
  FPDFPageObjMark_SetStringParam(
    document: number,
    object: number,
    mark: number,
    key: string,
    value: string,
  ): boolean;
  FPDFPageObjMark_GetParamStringValue(
    mark: number,
    key: string,
    buffer: number,
    bufferBytes: number,
    outBytes: number,
  ): boolean;
  FPDFImageObj_LoadJpegFileInline(
    pages: number,
    count: number,
    image: number,
    fileAccess: number,
  ): boolean;
  FPDFImageObj_SetBitmap(
    pages: number,
    count: number,
    image: number,
    bitmap: number,
  ): boolean;
  FPDFBitmap_CreateEx(
    width: number,
    height: number,
    format: number,
    buffer: number,
    stride: number,
  ): number;
  FPDFBitmap_Destroy(bitmap: number): void;
  FPDFImageObj_GetImagePixelSize(
    image: number,
    width: number,
    height: number,
  ): boolean;
  FPDFImageObj_SetMatrix(
    image: number,
    a: number,
    b: number,
    c: number,
    d: number,
    e: number,
    f: number,
  ): boolean;
  FPDFText_LoadStandardFont(document: number, name: string): number;
  FPDFText_LoadFont(
    document: number,
    data: number,
    size: number,
    fontType: number,
    cid: boolean,
  ): number;
  FPDFFont_Close(font: number): void;
  /** Copies the embedded font program; reports the size needed through `outBytes`. */
  FPDFFont_GetFontData(
    font: number,
    buffer: number,
    bytes: number,
    outBytes: number,
  ): boolean;
  FPDFText_SetText(textObject: number, text: number): boolean;
  FPDFText_LoadPage(page: number): number;
  FPDFText_ClosePage(textPage: number): void;
  FPDFText_CountChars(textPage: number): number;
  FPDFText_GetText(
    textPage: number,
    start: number,
    count: number,
    buffer: number,
  ): number;
}

/** The Emscripten runtime helpers `@embedpdf/pdfium` re-exports. */
export interface PdfiumRuntime {
  readonly HEAPU8: Uint8Array;
  readonly wasmExports: {
    malloc(size: number): number;
    free(pointer: number): void;
  };
  UTF16ToString(pointer: number, maxBytesToRead?: number): string;
  stringToUTF16(
    text: string,
    pointer: number,
    maxBytesToWrite?: number,
  ): number;
  addFunction(
    callback: (...args: number[]) => number,
    signature: string,
  ): number;
  removeFunction(pointer: number): void;
  getValue(pointer: number, type: "i32" | "float" | "double"): number;
  setValue(
    pointer: number,
    value: number,
    type: "i32" | "float" | "double",
  ): void;
}

export type PdfiumLibrary = PdfiumFunctions & {
  readonly pdfium: PdfiumRuntime;
};

interface PdfiumPackage {
  init(overrides: { readonly wasmBinary: ArrayBuffer }): Promise<PdfiumLibrary>;
}

/** `FPDF_SaveAsCopy` flags. */
export const SAVE_INCREMENTAL = 1;
export const SAVE_FULL = 2;

/** `FPDF_GetLastError` value for a missing or wrong password. */
const PASSWORD_ERROR = 4;

/** Size of `FPDF_FILEACCESS` in wasm32: file length, callback, parameter. */
const FILE_ACCESS_BYTES = 12;

/**
 * Owns one PDFium WASM instance and the memory traffic across its boundary.
 * PDFium keeps global heap state, so an instance must be used by one caller at
 * a time; the edit worker serializes calls.
 */
export class Pdfium {
  readonly lib: PdfiumLibrary;
  readonly #runtime: PdfiumRuntime;

  private constructor(lib: PdfiumLibrary) {
    this.lib = lib;
    this.#runtime = lib.pdfium;
  }

  /**
   * Instantiates PDFium from WASM bytes. The bytes are always supplied by the
   * caller: the package's default would fetch the module from a public CDN.
   */
  static async load(wasmBinary: ArrayBuffer | Uint8Array): Promise<Pdfium> {
    const { init } =
      (await import("@embedpdf/pdfium")) as unknown as PdfiumPackage;
    const lib = await init({ wasmBinary: exactArrayBuffer(wasmBinary) });
    lib.PDFiumExt_Init();
    return new Pdfium(lib);
  }

  malloc(size: number): number {
    const pointer = this.#runtime.wasmExports.malloc(Math.max(1, size));
    if (!pointer)
      throw new ViewerError("resource-limit", "PDFium is out of memory", {
        details: { bytes: size },
      });
    return pointer;
  }

  free(pointer: number): void {
    if (pointer) this.#runtime.wasmExports.free(pointer);
  }

  /** Copies bytes into the WASM heap; the caller frees the pointer. */
  writeBytes(bytes: Uint8Array): number {
    const pointer = this.malloc(bytes.byteLength);
    this.#runtime.HEAPU8.set(bytes, pointer);
    return pointer;
  }

  readBytes(pointer: number, length: number): Uint8Array {
    // The heap view is replaced when memory grows, so it is read on each call.
    return this.#runtime.HEAPU8.slice(pointer, pointer + length);
  }

  /** Writes a NUL-terminated UTF-16LE string; the caller frees the pointer. */
  writeWideString(text: string): number {
    const bytes = (text.length + 1) * 2;
    const pointer = this.malloc(bytes);
    this.#runtime.stringToUTF16(text, pointer, bytes);
    return pointer;
  }

  /** Reads a NUL-terminated UTF-16LE string of at most `bytes` bytes. */
  readWideStringAt(pointer: number, bytes: number): string {
    return this.#runtime.UTF16ToString(pointer, bytes);
  }

  /**
   * Reads a UTF-16LE string from an API that returns the byte size it needs
   * when called with an empty buffer.
   */
  readWideString(read: (buffer: number, bytes: number) => number): string {
    const bytes = read(0, 0);
    if (bytes <= 2) return "";
    const pointer = this.malloc(bytes);
    try {
      read(pointer, bytes);
      return this.#runtime.UTF16ToString(pointer, bytes);
    } finally {
      this.free(pointer);
    }
  }

  /**
   * Reads a UTF-16LE string from an API that reports the byte size it needs
   * through an `out_buflen` pointer.
   */
  readWideStringOut(
    read: (buffer: number, bytes: number, outBytes: number) => boolean,
  ): string {
    const outBytes = this.malloc(4);
    try {
      if (!read(0, 0, outBytes)) return "";
      const bytes = this.#runtime.getValue(outBytes, "i32");
      if (bytes <= 2) return "";
      const pointer = this.malloc(bytes);
      try {
        if (!read(pointer, bytes, outBytes)) return "";
        return this.#runtime.UTF16ToString(pointer, bytes);
      } finally {
        this.free(pointer);
      }
    } finally {
      this.free(outBytes);
    }
  }

  /** Reads a NUL-terminated UTF-8 string from an API that returns the byte size it needs. */
  readUtf8String(read: (buffer: number, bytes: number) => number): string {
    const bytes = read(0, 0);
    if (bytes <= 1) return "";
    const pointer = this.malloc(bytes);
    try {
      read(pointer, bytes);
      const raw = this.readBytes(pointer, bytes);
      const end = raw.indexOf(0);
      return new TextDecoder().decode(end < 0 ? raw : raw.subarray(0, end));
    } finally {
      this.free(pointer);
    }
  }

  /**
   * Calls `read` with a scratch buffer of `count` numbers of `type` and
   * returns what was written. Used for the many PDFium out-parameters.
   */
  readNumbers(
    count: number,
    type: "i32" | "float" | "double",
    read: (pointers: number[]) => boolean,
  ): number[] | undefined {
    const size = type === "double" ? 8 : 4;
    const base = this.malloc(count * size);
    try {
      const pointers = Array.from({ length: count }, (_, i) => base + i * size);
      if (!read(pointers)) return undefined;
      return pointers.map((pointer) => this.#runtime.getValue(pointer, type));
    } finally {
      this.free(base);
    }
  }

  /** Writes 32-bit integers into a fresh array; the caller frees it. */
  writeInt32Array(values: readonly number[]): number {
    const pointer = this.malloc(values.length * 4);
    values.forEach((value, index) =>
      this.#runtime.setValue(pointer + index * 4, value, "i32"),
    );
    return pointer;
  }

  /**
   * Runs `use` with an `FPDF_FILEACCESS` that serves `bytes` to PDFium through
   * a WASM table callback. Only for APIs that read the data before returning,
   * such as the inline JPEG loader.
   */
  withFileAccess<T>(bytes: Uint8Array, use: (fileAccess: number) => T): T {
    const runtime = this.#runtime;
    const callback = runtime.addFunction(
      (_parameter, position, buffer, size) => {
        if (position < 0 || size < 0 || position + size > bytes.byteLength)
          return 0;
        runtime.HEAPU8.set(bytes.subarray(position, position + size), buffer);
        return 1;
      },
      "iiiii",
    );
    const fileAccess = this.malloc(FILE_ACCESS_BYTES);
    try {
      runtime.setValue(fileAccess, bytes.byteLength, "i32");
      runtime.setValue(fileAccess + 4, callback, "i32");
      runtime.setValue(fileAccess + 8, 0, "i32");
      return use(fileAccess);
    } finally {
      this.free(fileAccess);
      runtime.removeFunction(callback);
    }
  }

  /** Opens a document from bytes; PDFium reads them lazily, so they stay in the heap. */
  openDocument(bytes: Uint8Array): PdfiumDocument {
    const data = this.writeBytes(bytes);
    const handle = this.lib.FPDF_LoadMemDocument(data, bytes.byteLength, "");
    if (!handle) {
      const error = this.lib.FPDF_GetLastError();
      this.free(data);
      throw error === PASSWORD_ERROR
        ? new ViewerError(
            "encrypted-document",
            "Password-protected PDF documents are not supported",
          )
        : new ViewerError("invalid-file", "PDFium could not open the PDF", {
            details: { pdfiumError: error },
          });
    }
    return new PdfiumDocument(this, handle, data);
  }

  createDocument(): PdfiumDocument {
    const handle = this.lib.FPDF_CreateNewDocument();
    if (!handle)
      throw new ViewerError("internal", "PDFium could not create a document");
    return new PdfiumDocument(this, handle, 0);
  }
}

export class PdfiumDocument {
  readonly #pdfium: Pdfium;
  readonly #data: number;
  #handle: number;

  constructor(pdfium: Pdfium, handle: number, data: number) {
    this.#pdfium = pdfium;
    this.#handle = handle;
    this.#data = data;
  }

  get handle(): number {
    if (!this.#handle)
      throw new ViewerError("lifecycle-error", "The PDF document is closed");
    return this.#handle;
  }

  /** Serializes the document; `incremental` appends to the bytes it was opened from. */
  save(mode: "incremental" | "full"): Uint8Array {
    const { lib } = this.#pdfium;
    const writer = lib.PDFiumExt_OpenFileWriter();
    if (!writer)
      throw new ViewerError("internal", "PDFium could not open a file writer");
    try {
      const flags = mode === "incremental" ? SAVE_INCREMENTAL : SAVE_FULL;
      if (!lib.FPDF_SaveAsCopy(this.handle, writer, flags))
        throw new ViewerError("internal", "PDFium could not save the PDF", {
          details: { mode, pdfiumError: lib.FPDF_GetLastError() },
        });
      const size = lib.PDFiumExt_GetFileWriterSize(writer);
      const buffer = this.#pdfium.malloc(size);
      try {
        lib.PDFiumExt_GetFileWriterData(writer, buffer, size);
        return this.#pdfium.readBytes(buffer, size);
      } finally {
        this.#pdfium.free(buffer);
      }
    } finally {
      lib.PDFiumExt_CloseFileWriter(writer);
    }
  }

  close(): void {
    if (!this.#handle) return;
    this.#pdfium.lib.FPDF_CloseDocument(this.#handle);
    this.#handle = 0;
    this.#pdfium.free(this.#data);
  }
}

function exactArrayBuffer(bytes: ArrayBuffer | Uint8Array): ArrayBuffer {
  if (bytes instanceof ArrayBuffer) return bytes;
  return bytes.buffer instanceof ArrayBuffer &&
    bytes.byteOffset === 0 &&
    bytes.byteLength === bytes.buffer.byteLength
    ? bytes.buffer
    : bytes.slice().buffer;
}
