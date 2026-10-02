import assert from "node:assert/strict";
import { mkdtemp, rm, writeFile } from "node:fs/promises";
import { tmpdir } from "node:os";
import { resolve } from "node:path";
import { pathToFileURL } from "node:url";
import test from "node:test";

import { build } from "esbuild";

/*
 * A consumer that imports the package by name and bundles it with esbuild
 * without code splitting inlines the lazily loaded edit engines as lazily
 * initialised modules. A value the entry only re-exports with `export *`
 * from such a module then comes out undefined (web-doc 0.7.0 shipped
 * `WorkerRpcClient` that way). This bundles a consumer that names every
 * value export and checks that each one is defined.
 */

const root = resolve(import.meta.dirname, "..");
const entries = {
  "web-doc": resolve(root, "packages/viewer/src/index.ts"),
  "web-doc/headless": resolve(root, "packages/viewer/src/headless.ts"),
};

/** The smallest DOM the package touches while its modules evaluate. */
function installDomShim() {
  const element = () => ({
    style: {},
    append() {},
    appendChild() {},
    setAttribute() {},
    addEventListener() {},
    removeEventListener() {},
    querySelector: () => null,
  });
  globalThis.window ??= globalThis;
  globalThis.self ??= globalThis;
  globalThis.document ??= {
    createElement: element,
    createElementNS: element,
    head: element(),
    body: element(),
    querySelector: () => null,
    addEventListener() {},
    removeEventListener() {},
  };
  globalThis.navigator ??= { language: "en" };
}

async function bundle(entryPoint, outfile) {
  await build({
    entryPoints: [entryPoint],
    bundle: true,
    format: "esm",
    platform: "browser",
    target: ["es2022"],
    outfile,
    logLevel: "silent",
    loader: { ".wasm": "file" },
  });
  return import(pathToFileURL(outfile).href);
}

/** Classes and constants are assigned inside esbuild's lazy wrappers; function declarations are hoisted out, so only the former can come out undefined. */
function canBeLazy(value) {
  return (
    typeof value !== "function" ||
    /^class[\s{]/.test(Function.prototype.toString.call(value))
  );
}

for (const [specifier, entry] of Object.entries(entries))
  test(`every class and constant of ${specifier} stays defined in a consumer bundle without code splitting`, async () => {
    const directory = await mkdtemp(resolve(tmpdir(), "web-doc-consumer-"));
    try {
      installDomShim();
      // The names a consumer can import: what the package itself exports.
      const library = await bundle(entry, resolve(directory, "library.mjs"));
      const names = Object.keys(library)
        .filter((name) => canBeLazy(library[name]))
        .sort();
      assert.ok(names.includes("ViewerClient"));
      if (specifier === "web-doc") assert.ok(names.includes("WorkerRpcClient"));
      // One consumer per name: a consumer that imports a single value keeps
      // nothing else alive, so no other module initialises the lazy one.
      const consumers = await Promise.all(
        names.map(async (name) => {
          const file = resolve(directory, `consumer-${name}.mjs`);
          await writeFile(
            file,
            `import { ${name} } from ${JSON.stringify(entry)};\nexport { ${name} };\n`,
          );
          return file;
        }),
      );
      // `.mjs`, so Node reads the bundles as ES modules wherever the
      // temporary directory is (a `.js` file there would be CommonJS).
      await build({
        entryPoints: consumers,
        bundle: true,
        format: "esm",
        platform: "browser",
        target: ["es2022"],
        outdir: resolve(directory, "out"),
        outExtension: { ".js": ".mjs" },
        logLevel: "silent",
        loader: { ".wasm": "file" },
      });
      const missing = [];
      for (const name of names) {
        const bundled = await import(
          pathToFileURL(resolve(directory, "out", `consumer-${name}.mjs`)).href
        );
        if (bundled[name] === undefined) missing.push(name);
      }
      assert.deepEqual(
        missing,
        [],
        `undefined in a consumer bundle: ${missing.join(", ")}`,
      );
    } finally {
      await rm(directory, { recursive: true, force: true });
    }
  });
