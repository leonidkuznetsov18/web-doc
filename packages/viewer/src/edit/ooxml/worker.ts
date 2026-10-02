import { ViewerError } from "../../errors.js";
import type { WorkerLike } from "../../worker-client.js";
import type { EditEngineContext } from "../engine.js";
import { packageRelativeUrl, resolveAssetUrl } from "../worker-engine.js";

/*
 * The OOXML edit worker serves one package for one session, as a PPTX or a
 * DOCX engine; both providers start it the same way.
 */

export interface OoxmlEditProviderOptions {
  readonly workerUrl?: string | URL;
  /** Test hook: supplies the worker instead of the packaged script. */
  readonly createWorker?: () => WorkerLike;
}

export function createOoxmlEditWorker(
  options: OoxmlEditProviderOptions,
  context: EditEngineContext,
): WorkerLike {
  if (options.createWorker) return options.createWorker();
  if (typeof Worker === "undefined")
    throw new ViewerError(
      "edit-unsupported",
      `${context.format.toUpperCase()} editing requires a browser module Worker`,
      { details: { format: context.format, reason: "no-worker" } },
    );
  const url = resolveAssetUrl(
    options.workerUrl,
    context.assetBaseUrl,
    context.assetBaseUrl
      ? new URL("workers/ooxml-edit-worker.js", context.assetBaseUrl)
      : packageRelativeUrl(
          "../../../workers/ooxml-edit-worker.js",
          import.meta.url,
        ),
  );
  return new Worker(url, { type: "module", name: "web-doc-ooxml-edit" });
}
