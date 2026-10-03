import { execFileSync, spawnSync } from "node:child_process";
import { readFile } from "node:fs/promises";
import { resolve } from "node:path";
import { pathToFileURL } from "node:url";

// The checks that keep `main` releasable. GitHub Actions runs only the
// release (release.yml); everything else runs here, from the git hooks in
// .githooks (installed by `npm ci` through the `prepare` script):
//
//   pre-commit   node scripts/verify.mjs --staged   Prettier and rustfmt over the staged files
//   pre-push     node scripts/verify.mjs --push     the checks the pushed commits need
//   npm run verify                                  every check, as CI ran them
//
// WEB_DOC_VERIFY=full|quick overrides what a push runs: every check, or only
// `npm run check`.

const root = resolve(import.meta.dirname, "..");

const RUNTIME = [
  /^packages\//,
  /^crates\//,
  /^scripts\//,
  /^Cargo\.(toml|lock)$/,
  /^package(-lock)?\.json$/,
  /^rust-toolchain\.toml$/,
  /^tsconfig\.base\.json$/,
];
const BROWSER = [
  /^tests\//,
  /^examples\//,
  /^playwright(\.[a-z]+)?\.config\.ts$/,
];
const PAGES = [/^docs\//, /^examples\//];
const RUST = [/^crates\//, /^fuzz\//, /^Cargo\.(toml|lock)$/];
const DEPENDENCIES = [
  /^package(-lock)?\.json$/,
  /^Cargo\.(toml|lock)$/,
  /^deny\.toml$/,
];

const touches = (files, patterns) =>
  files.some((file) => patterns.some((pattern) => pattern.test(file)));

/**
 * Every check, in the order CI ran them, with the changes that make it
 * relevant. `npm run check` (formatting, types, unit and regression tests,
 * clippy, licenses, the repository audit) always runs.
 */
export const CHECKS = [
  { script: "check", needs: () => true },
  { script: "test:pack", needs: (files) => touches(files, RUNTIME) },
  { script: "fuzz:js", needs: (files) => touches(files, RUNTIME) },
  { script: "report:size", needs: (files) => touches(files, RUNTIME) },
  {
    script: "audit:vulnerabilities",
    needs: (files) => touches(files, DEPENDENCIES),
  },
  { script: "report:sbom", needs: (files) => touches(files, DEPENDENCIES) },
  {
    script: "test:e2e",
    needs: (files) => touches(files, [...RUNTIME, ...BROWSER]),
  },
  {
    script: "test:e2e:matrix",
    needs: (files) => touches(files, [...RUNTIME, ...BROWSER]),
  },
  {
    script: "test:e2e:performance",
    needs: (files) => touches(files, [...RUNTIME, ...BROWSER]),
  },
  {
    script: "pages:build",
    needs: (files) => touches(files, [...RUNTIME, ...PAGES]),
  },
  {
    script: "test:pages",
    needs: (files) => touches(files, [...RUNTIME, ...PAGES]),
  },
  {
    script: "fuzz:rust",
    env: { FUZZ_SECONDS: "2" },
    needs: (files) => touches(files, RUST),
  },
];

/** The checks a change to `files` needs; `mode` is "full", "quick" or unset. */
export function plan(files, mode) {
  if (mode === "full") return CHECKS;
  if (mode === "quick") return CHECKS.slice(0, 1);
  return CHECKS.filter((check) => check.needs(files));
}

function git(args) {
  return execFileSync("git", args, { cwd: root, encoding: "utf8" }).trim();
}

const ZERO = /^0+$/;

/** Files the pushed commits change, from the refs git hands pre-push on stdin. */
async function pushedFiles() {
  let input = "";
  for await (const chunk of process.stdin) input += chunk;
  const files = new Set();
  for (const line of input.split("\n").filter(Boolean)) {
    const [, localSha, , remoteSha] = line.split(" ");
    if (ZERO.test(localSha)) continue; // a deleted branch pushes no code
    let base = ZERO.test(remoteSha) ? "" : remoteSha;
    try {
      if (base) git(["cat-file", "-e", `${base}^{commit}`]);
    } catch {
      base = ""; // the remote moved past what this clone knows
    }
    if (!base) {
      try {
        base = git(["merge-base", localSha, "origin/main"]);
      } catch {
        return undefined; // nothing to compare with: run everything
      }
    }
    for (const file of git(["diff", "--name-only", base, localSha]).split("\n"))
      if (file) files.add(file);
  }
  return [...files];
}

/** Runs the checks in order and stops at the first failure; resolves its exit status. */
function runAll(checks) {
  for (const check of checks) {
    console.log(`\n▶ npm run ${check.script}`);
    const result = spawnSync("npm", ["run", check.script], {
      cwd: root,
      stdio: "inherit",
      env: { ...process.env, ...check.env },
    });
    if (result.status !== 0) {
      console.error(
        `\n✖ npm run ${check.script} failed. Fix it and push again; WEB_DOC_VERIFY=quick git push runs only \`npm run check\`.`,
      );
      return result.status ?? 1;
    }
  }
  return 0;
}

/** Tracked reports under artifacts/ that differ from HEAD. */
function changedReports() {
  return git(["diff", "--name-only", "--", "artifacts"])
    .split("\n")
    .filter(Boolean);
}

/**
 * The paths `npm run format:check` formats; the hook checks no more, so a
 * generated file outside them (CHANGELOG.md) commits as it is written.
 */
async function formatScope() {
  const manifest = JSON.parse(
    await readFile(resolve(root, "package.json"), "utf8"),
  );
  const command = manifest.scripts["format:check"].split("&&")[0];
  return command
    .replace(/^\s*prettier --check\s+/, "")
    .trim()
    .split(/\s+/);
}

/** Prettier over the staged content of staged files, rustfmt when Rust is staged. */
async function verifyStaged() {
  const staged = git([
    "diff",
    "--cached",
    "--name-only",
    "--diff-filter=ACMR",
    "-z",
  ])
    .split("\0")
    .filter(Boolean);
  const prettier = await import("prettier");
  const scope = await formatScope();
  const unformatted = [];
  for (const file of staged.filter((file) =>
    scope.some((path) => file === path || file.startsWith(`${path}/`)),
  )) {
    const path = resolve(root, file);
    const info = await prettier.getFileInfo(path, {
      ignorePath: resolve(root, ".prettierignore"),
    });
    if (info.ignored || !info.inferredParser) continue;
    const source = git(["show", `:${file}`]);
    const options = { ...(await prettier.resolveConfig(path)), filepath: path };
    if (!(await prettier.check(`${source}\n`, options))) unformatted.push(file);
  }
  if (unformatted.length > 0) {
    console.error(
      `✖ Not formatted with Prettier:\n  ${unformatted.join("\n  ")}\nRun npx prettier --write on them and stage the result.`,
    );
    process.exit(1);
  }
  if (staged.some((file) => file.endsWith(".rs"))) {
    const result = spawnSync("cargo", ["fmt", "--all", "--check"], {
      cwd: root,
      stdio: "inherit",
    });
    if (result.status !== 0) {
      console.error("✖ Rust is not formatted: run cargo fmt --all.");
      process.exit(result.status ?? 1);
    }
  }
}

async function main(argv) {
  // Actions only release, and the release builds and pack-tests the package
  // itself; its own commit and push must not wait on the developer checks.
  if (process.env.CI && !argv.includes("--all")) return;
  const mode = process.env.WEB_DOC_VERIFY;
  if (argv.includes("--staged")) return verifyStaged();
  if (argv.includes("--all")) process.exit(runAll(plan([], "full")));
  if (argv.includes("--push")) {
    const files = await pushedFiles();
    const checks = plan(files ?? [], files ? mode : "full");
    console.log(
      `Pre-push checks: ${checks.map((check) => check.script).join(", ")}`,
    );
    // The browser, packaging and size checks rewrite the tracked reports in
    // artifacts/; a push leaves the ones it found clean as they were, so the
    // working tree it leaves is the one it started from. `npm run verify`
    // keeps the fresh reports.
    const dirty = new Set(changedReports());
    const status = runAll(checks);
    const rewritten = changedReports().filter((path) => !dirty.has(path));
    if (rewritten.length > 0) git(["checkout", "--", ...rewritten]);
    process.exit(status);
  }
  throw new Error("Usage: node scripts/verify.mjs --staged | --push | --all");
}

if (import.meta.url === pathToFileURL(process.argv[1] ?? "").href) {
  await main(process.argv.slice(2));
}
