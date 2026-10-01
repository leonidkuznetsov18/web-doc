import type { PdfEditSession } from "./pdf/types.js";

/**
 * The session of the loaded document, discriminated by `format`. Format
 * modules add their member as they ship.
 */
export type EditSession = PdfEditSession;
