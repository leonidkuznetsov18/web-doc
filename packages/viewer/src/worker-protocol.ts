import type {
  DocumentFormat,
  DocumentInfo,
  ResourceLimits,
  TextRun,
  ViewerErrorData,
  ViewerProgress,
  ViewerWarning,
} from "./contracts.js";
import type { EditableFormat } from "./edit/types.js";

export type DocumentWorkerOperation =
  | "init"
  | "open"
  | "get-info"
  | "render"
  | "get-text-map"
  | "close"
  | "destroy";

/** Operations of an edit engine worker; payloads mirror the engine interface. */
export type EditWorkerOperation =
  | "edit-init"
  | "edit-open"
  | "edit-validate"
  | "edit-apply"
  | "edit-materialize"
  | "edit-restore"
  | "edit-put-asset"
  | "edit-elements"
  | "edit-element"
  | "edit-elements-at"
  | "edit-find-text"
  | "edit-text-layout"
  | "edit-text-paragraph"
  | "edit-text-font"
  | "edit-position-at"
  | "edit-range-rects"
  | "edit-render-without"
  | "edit-page-layout"
  | "edit-pptx-slides"
  | "edit-pptx-layouts"
  | "edit-docx-revisions"
  | "edit-docx-text-style"
  | "edit-docx-text-colors"
  | "edit-pptx-text-style"
  | "edit-dispose";

export type WorkerOperation = DocumentWorkerOperation | EditWorkerOperation;

export interface WorkerRequest {
  readonly kind: "request";
  readonly id: number;
  readonly operation: WorkerOperation;
  readonly payload?: unknown;
}

export interface WorkerCancel {
  readonly kind: "cancel";
  readonly id: number;
}

export interface WorkerSuccess {
  readonly kind: "success";
  readonly id: number;
  readonly result?: unknown;
}

export interface WorkerFailure {
  readonly kind: "failure";
  readonly id: number;
  readonly error: ViewerErrorData;
}

export interface WorkerProgressMessage {
  readonly kind: "progress";
  readonly id: number;
  readonly progress: ViewerProgress;
}

export interface WorkerWarningMessage {
  readonly kind: "warning";
  readonly id: number;
  readonly warning: ViewerWarning;
}

export type WorkerInboundMessage = WorkerRequest | WorkerCancel;
export type WorkerOutboundMessage =
  WorkerSuccess | WorkerFailure | WorkerProgressMessage | WorkerWarningMessage;

export interface WorkerOpenPayload {
  readonly data: ArrayBuffer;
  readonly format: DocumentFormat;
  readonly limits: ResourceLimits;
  readonly fileName?: string;
  readonly contentType?: string;
}

export interface WorkerRenderPayload {
  readonly pageIndex: number;
  readonly zoom: number;
  readonly devicePixelRatio: number;
}

export interface EditWorkerInitPayload {
  /** Where the worker fetches the engine's WebAssembly from. */
  readonly wasmUrl: string;
  /** TrueType font for text the standard PDF fonts cannot encode. */
  readonly fallbackFontUrl?: string;
}

/** A host-registered font as the edit worker receives it. */
export interface EditWorkerFont {
  readonly family: string;
  readonly weight: number;
  readonly style: "normal" | "italic" | "oblique";
  /** Font bytes, or an absolute URL the worker fetches on first use. */
  readonly source: ArrayBuffer | string;
}

export interface EditWorkerOpenPayload {
  readonly data: ArrayBuffer;
  readonly limits: ResourceLimits;
  /** The session format; a worker that serves several formats picks its engine by it. */
  readonly format?: EditableFormat;
  readonly fileName?: string;
  readonly fonts?: readonly EditWorkerFont[];
}

/** A rendered page as the worker returns it; `data` is transferred, not copied. */
export interface EditWorkerBitmap {
  readonly pageIndex: number;
  readonly scale: number;
  readonly width: number;
  readonly height: number;
  readonly data: ArrayBuffer;
}

export interface EditWorkerOpenResult {
  readonly pageCount: number;
}

export type WorkerOperationResult =
  | DocumentInfo
  | readonly TextRun[]
  | ImageBitmap
  | ArrayBuffer
  | EditWorkerOpenResult
  | undefined;

export function transferablesFor(value: unknown): Transferable[] {
  if (value instanceof ArrayBuffer) return [value];
  if (typeof ImageBitmap !== "undefined" && value instanceof ImageBitmap)
    return [value];
  if (ArrayBuffer.isView(value)) return [value.buffer];
  if (value && typeof value === "object") {
    const data = (value as { data?: unknown }).data;
    if (data instanceof ArrayBuffer) return [data];
  }
  return [];
}
