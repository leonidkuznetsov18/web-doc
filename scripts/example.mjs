import { context } from "esbuild";
import { copyFile, cp, mkdir } from "node:fs/promises";
import { dirname, resolve } from "node:path";
import { createRequire } from "node:module";

const mode = process.argv[2] ?? "build";
const root = process.cwd();
const outdir = resolve(root, "dist");
const viewerRequire = createRequire(
  new URL("../packages/viewer/package.json", import.meta.url),
);
const officeDist = dirname(viewerRequire.resolve("@silurus/ooxml/docx"));

await mkdir(outdir, { recursive: true });
await copyFile(resolve(root, "index.html"), resolve(outdir, "index.html"));
await copyFile(
  resolve(root, "../../packages/viewer/dist/styles.css"),
  resolve(outdir, "viewer.css"),
);
// The engine on its own as well, so browser tests can ask it for geometry.
await cp(officeDist, resolve(outdir, "vendor/ooxml"), {
  recursive: true,
});
for (const wasm of [
  "docx_parser_bg.wasm",
  "xlsx_parser_bg.wasm",
  "pptx_parser_bg.wasm",
])
  await copyFile(resolve(officeDist, wasm), resolve(outdir, wasm));
try {
  await cp(
    resolve(root, "../../packages/viewer/dist/workers"),
    resolve(outdir, "workers"),
    { recursive: true },
  );
} catch (error) {
  if (
    !(error instanceof Error) ||
    !("code" in error) ||
    error.code !== "ENOENT"
  )
    throw error;
}
try {
  await cp(
    resolve(root, "../../packages/viewer/dist/fonts"),
    resolve(outdir, "fonts"),
    { recursive: true },
  );
} catch (error) {
  if (
    !(error instanceof Error) ||
    !("code" in error) ||
    error.code !== "ENOENT"
  )
    throw error;
}
try {
  await cp(
    resolve(root, "../../packages/viewer/dist/assets"),
    resolve(outdir, "assets"),
    { recursive: true },
  );
} catch (error) {
  if (
    !(error instanceof Error) ||
    !("code" in error) ||
    error.code !== "ENOENT"
  )
    throw error;
}
try {
  await cp(resolve(root, "../../.cache/corpus"), resolve(outdir, "corpus"), {
    recursive: true,
  });
} catch (error) {
  if (
    !(error instanceof Error) ||
    !("code" in error) ||
    error.code !== "ENOENT"
  )
    throw error;
}
try {
  await cp(
    resolve(root, "../../.cache/wasm-bindgen"),
    resolve(outdir, "wasm"),
    {
      recursive: true,
    },
  );
} catch (error) {
  if (
    !(error instanceof Error) ||
    !("code" in error) ||
    error.code !== "ENOENT"
  )
    throw error;
}

const buildContext = await context({
  entryPoints: [resolve(root, "src/main.ts")],
  bundle: true,
  format: "esm",
  outfile: resolve(outdir, "main.js"),
  sourcemap: true,
  target: ["es2022"],
});

if (mode === "serve") {
  await buildContext.watch();
  const server = await buildContext.serve({
    host: process.env.HOST ?? "127.0.0.1",
    port: Number(process.env.PORT ?? 4173),
    servedir: outdir,
  });
  console.log(`Example server: http://${server.host}:${server.port}`);
} else {
  await buildContext.rebuild();
  await buildContext.dispose();
}
