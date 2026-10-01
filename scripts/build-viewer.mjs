import { execFileSync } from "node:child_process";
import { cp, mkdir, rm, writeFile } from "node:fs/promises";
import { resolve } from "node:path";
import { pathToFileURL } from "node:url";
import { build } from "esbuild";

const root = resolve(import.meta.dirname, "..");
const packageRoot = resolve(root, "packages/viewer");
const dist = resolve(packageRoot, "dist");

await rm(dist, { recursive: true, force: true });

execFileSync(
  process.execPath,
  [
    resolve(root, "node_modules/typescript/bin/tsc"),
    "-p",
    "tsconfig.build.json",
  ],
  { cwd: packageRoot, stdio: "inherit" },
);

const { viewerCss } = await import(
  `${pathToFileURL(resolve(dist, "ui-styles.js")).href}?build=${Date.now()}`
);
await writeFile(resolve(dist, "styles.css"), `${viewerCss.trim()}\n`);

await mkdir(resolve(dist, "assets"), { recursive: true });
await mkdir(resolve(dist, "fonts"), { recursive: true });
await mkdir(resolve(dist, "workers"), { recursive: true });
for (const [entry, outfile] of [
  ["src/csv-worker.ts", "workers/csv-worker.js"],
  // Bundled so the worker realm gets Fuse.js without a bare-specifier import.
  ["src/fuzzy-search-worker.ts", "workers/fuzzy-search-worker.js"],
  // Bundles the PDFium JavaScript glue; the WASM itself is fetched at runtime.
  ["src/pdf-edit-worker.ts", "workers/pdf-edit-worker.js"],
])
  await build({
    entryPoints: [resolve(packageRoot, entry)],
    bundle: true,
    format: "esm",
    platform: "browser",
    target: ["es2022"],
    outfile: resolve(dist, outfile),
  });
await cp(
  resolve(root, ".cache/wasm-bindgen/legacy"),
  resolve(dist, "assets/legacy"),
  { recursive: true },
);
for (const asset of ["image"])
  await cp(
    resolve(root, ".cache/wasm-bindgen", asset),
    resolve(dist, "assets", asset),
    { recursive: true },
  );
for (const worker of ["legacy-converter-worker.js", "image-worker.js"])
  await cp(resolve(dist, worker), resolve(dist, "workers", worker));

const pdfJsRoot = resolve(root, "node_modules/pdfjs-dist");
const pdfJsAssets = resolve(dist, "assets/pdfjs");
await mkdir(pdfJsAssets, { recursive: true });
for (const directory of ["cmaps", "standard_fonts", "wasm", "iccs"])
  await cp(resolve(pdfJsRoot, directory), resolve(pdfJsAssets, directory), {
    recursive: true,
  });
await cp(
  // Keep the worker on the same compatibility build as the lazy-loaded
  // PDF.js runtime. Mixing the modern worker with the legacy runtime still
  // leaves unsupported intrinsics inside the worker realm.
  resolve(pdfJsRoot, "legacy/build/pdf.worker.min.mjs"),
  resolve(dist, "workers/pdf.worker.min.mjs"),
);

await mkdir(resolve(dist, "assets/pdfium"), { recursive: true });
await cp(
  resolve(root, "node_modules/@embedpdf/pdfium/dist/pdfium.wasm"),
  resolve(dist, "assets/pdfium/pdfium.wasm"),
);

await cp(resolve(packageRoot, "fonts"), resolve(dist, "fonts"), {
  recursive: true,
});
for (const [source, destination] of [
  ["LICENSE", "LICENSE-MIT"],
  ["LICENSE-APACHE", "LICENSE-APACHE"],
  ["THIRD_PARTY_NOTICES.md", "THIRD_PARTY_NOTICES.md"],
])
  await cp(resolve(root, source), resolve(packageRoot, destination));

console.log("Viewer TypeScript, workers, WASM, CSS, and font assets built");
