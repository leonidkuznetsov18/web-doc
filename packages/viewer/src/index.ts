export * from "./contracts.js";
export * from "./edit/types.js";
export * from "./edit/sessions.js";
export * from "./edit/pdf/types.js";
export * from "./edit/pptx/types.js";
export * from "./edit/docx/types.js";
export * from "./client.js";
export * from "./detect.js";
export * from "./errors.js";
export * from "./format.js";
export * from "./limits.js";
export * from "./interaction.js";
export * from "./fuzzy-search.js";
export * from "./fuzzy-worker-protocol.js";
export * from "./fuzzy-worker-client.js";
export * from "./search-reveal.js";
export * from "./adapters/docx-images.js";
export * from "./render-scheduler.js";
export * from "./i18n.js";
export * from "./font-manifest.js";
export * from "./fonts.js";
export * from "./ui.js";
export * from "./ui-styles.js";
export * from "./registry.js";
export * from "./viewer.js";
export * from "./worker-client.js";
export * from "./worker-adapter.js";
export * from "./worker-endpoint.js";
export * from "./worker-protocol.js";
export * from "./adapters/office.js";
export * from "./adapters/pdf.js";
export * from "./adapters/image.js";
export * from "./adapters/csv.js";
export * from "./adapters/csv-parser.js";
export * from "./adapters/svg.js";

import type { ViewerClientOptions, ViewerOptions } from "./contracts.js";
import { ViewerClient } from "./client.js";

export function createViewer(
  options: ViewerOptions = {},
  clientOptions: ViewerClientOptions = {},
) {
  return ViewerClient.create(clientOptions).createViewer(options);
}
