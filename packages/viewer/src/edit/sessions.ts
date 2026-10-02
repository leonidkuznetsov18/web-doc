import type { DocxEditSession } from "./docx/types.js";
import type { PdfEditSession } from "./pdf/types.js";
import type { PptxEditSession } from "./pptx/types.js";

/**
 * The session of the loaded document, discriminated by `format`. Format
 * modules add their member as they ship.
 */
export type EditSession = PdfEditSession | PptxEditSession | DocxEditSession;
