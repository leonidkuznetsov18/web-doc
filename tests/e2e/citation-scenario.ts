/** Shared browser regression, also runnable in an existing signed-in browser. */
export async function citationBoundaryScenario(
  mode: "worker" | "main" | "unavailable",
) {
  const { ViewerClient } = await import("/main.js");
  const passage =
    "The shift from manual review to automated checks is the difference between incremental improvement and reliable delivery.";
  const texts = ["Earlier paragraph.", passage, "The next chapter"];
  const warnings: string[] = [];
  let offset = 0;
  const runs = texts.map((text, index) => {
    const logicalStart = offset;
    offset += text.length;
    return {
      text,
      x: 24,
      y: 24 + index * 40,
      width: text.length * 8,
      height: 20,
      font: "14px sans-serif",
      fontSize: 14,
      direction: "ltr" as const,
      textLayer: "docx" as const,
      coordinateWidth: 1100,
      coordinateHeight: 180,
      logicalStart,
      logicalEnd: offset,
    };
  });
  const host = document.createElement("div");
  Object.assign(host.style, { width: "1100px", height: "240px" });
  document.body.append(host);
  const client = ViewerClient.create({
    assetBaseUrl: new URL(
      mode === "unavailable" ? "/missing-worker/" : "/",
      location.href,
    ),
    logger: { warn: (message: string) => warnings.push(message) },
    adapters: [
      {
        id: "citation-text-runs",
        formats: ["pdf"] as const,
        open: async () => ({}),
        getInfo: async () =>
          ({
            format: "pdf",
            unit: "page",
            pageCount: 1,
            pageSizes: [{ width: 1100, height: 180 }],
          }) as const,
        render: async (_handle: unknown, canvas: HTMLCanvasElement) => {
          canvas.width = 1100;
          canvas.height = 180;
          const context = canvas.getContext("2d");
          if (!context) throw new Error("Canvas unavailable");
          context.font = "14px sans-serif";
          context.textBaseline = "top";
          for (const run of runs) context.fillText(run.text, run.x, run.y);
        },
        getTextMap: async () => runs,
        close: () => {},
      },
    ],
  });
  const viewer = client.createViewer({ container: host, initialZoom: 1 });
  try {
    await viewer.load(new TextEncoder().encode("%PDF-1.7\ncitation fixture"), {
      fileName: "fixture.pdf",
    });
    const result = await viewer.search(
      passage.replace("and reliable", "and\nreliable"),
      { fuzzy: { worker: mode !== "main" } },
    );
    await new Promise<void>((resolve, reject) => {
      const timeout = setTimeout(() => {
        observer.disconnect();
        reject(new Error("Highlight did not render"));
      }, 5000);
      const check = () => {
        if (
          host.querySelectorAll('[data-zrimo-layer="highlight"] > div').length >
          0
        ) {
          clearTimeout(timeout);
          observer.disconnect();
          resolve();
        }
      };
      const observer = new MutationObserver(check);
      observer.observe(host, { childList: true, subtree: true });
      check();
    });
    const highlights = host.querySelectorAll(
      '[data-zrimo-layer="highlight"] > div',
    );
    if (result.matches.length !== 1)
      throw new Error(`Expected one match, got ${result.matches.length}`);
    const match = result.matches[0];
    if (
      match?.start !== texts[0]!.length ||
      match.end !== texts[0]!.length + passage.length ||
      match.text !== passage
    )
      throw new Error(`Incorrect passage: ${JSON.stringify(result)}`);
    if (highlights.length !== 1)
      throw new Error(`Expected one highlighted run, got ${highlights.length}`);
    if ((mode === "unavailable") !== warnings.length > 0)
      throw new Error(`Unexpected worker fallback: ${warnings}`);
    return {
      mode,
      strategy: result.strategy,
      start: match.start,
      end: match.end,
      highlightCount: highlights.length,
      warnings,
    };
  } finally {
    await viewer.destroy();
    await client.destroy();
    host.remove();
  }
}
