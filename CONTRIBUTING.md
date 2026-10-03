# Contributing to Zrimo

Thank you for helping improve Zrimo. Bug reports should include the browser, integration mode, document format and the smallest reproducible input that you are allowed to share. Never commit private documents or customer data.

## Development setup

Install Node.js 22.13+ or 24+, npm 11, and Rust 1.94.1 with the `wasm32-unknown-unknown` target.

```bash
npm ci
npm run build
npm run check
```

## Checks

GitHub Actions run no checks; the only workflow that runs on its own is the release. Every check runs on your machine, from the git hooks that `npm ci` installs (`core.hooksPath` points at `.githooks`):

- `commit-msg` checks the message with commitlint.
- `pre-commit` checks the staged files with Prettier, and with rustfmt when Rust is staged.
- `pre-push` runs what the pushed commits need (`scripts/verify.mjs`). `npm run check` always runs. Changes to the viewer, the Rust crates, the scripts or the dependencies add packaging, JS fuzzing, the size budget, the Pages build and all three browser suites; browser specs and examples add the browser suites; docs add the Pages build and its browser check; lockfiles add the vulnerability audit and the SBOM; Rust and fuzz targets add a short cargo-fuzz run.

`WEB_DOC_VERIFY=full git push` runs every check, `WEB_DOC_VERIFY=quick git push` only `npm run check`, and `npm run verify` runs every check without pushing. Nothing else checks a change, so do not push with `--no-verify`.

The full set needs the Playwright browsers (`npx playwright install chromium firefox webkit`), `cargo install cargo-audit --locked` for the vulnerability audit, and a nightly toolchain with `cargo install cargo-fuzz --locked` for the Rust fuzz run. The performance suite keeps the DOCX edit latency budget on every matrix browser and runs with one worker, after the parallel functional suites.

Public qualification fixtures are downloaded into ignored `.cache/corpus/` with pinned hashes. User-provided regression files belong outside the repository or in ignored `.tmp/`; do not add them to tests, docs, screenshots or release artifacts.

## Changes

- Add focused unit and browser coverage for behavior changes.
- Keep parsing bounded and fail closed on malformed or unsupported input.
- Update public API and integration documentation with the code.
- Run `npm run audit:repository` and `npm run test:pack` before submitting packaging changes.

## Commits and releases

Every commit that lands on `main` must follow [Conventional Commits](https://www.conventionalcommits.org/), because the changelog and the version number are generated from the history:

```
<type>(<scope>): <subject>

<body>

BREAKING CHANGE: <what changed for integrators>
```

- `type` is one of `feat`, `fix`, `perf`, `revert`, `refactor`, `docs`, `test`, `build`, `ci`, `chore`.
- `scope` is required: `viewer`, `core`, `docs`, `examples`, `fuzz`, `ci`, `release`, `deps`, `repo` or `brand`. Extend the list in `commitlint.config.mjs` when a new area appears.
- The subject is lower-case, imperative, at most 72 characters including the prefix, without a trailing period. Body lines wrap at 100 characters.
- `feat` produces a minor release, `fix`/`perf`/`revert` a patch release, and a `BREAKING CHANGE:` footer (or `feat!:`) a major release. Only these types appear in `CHANGELOG.md`; everything else is release-neutral and hidden from it.
- Pull request titles follow the same format: they become the commit subject on a squash merge.

The `commit-msg` hook checks every message locally; no workflow repeats it, so keep the hooks installed.

Releases are automatic. On every push to `main`, the `Release` workflow runs [semantic-release](https://semantic-release.gitbook.io/): it derives the next version from the commits since the last tag, updates `CHANGELOG.md`, `packages/viewer/package.json`, `package-lock.json` and `release-status.json` in a `chore(release): x.y.z [skip ci]` commit, tags it `vx.y.z`, and publishes a GitHub release with the generated notes plus the verified `zrimo-viewer-x.y.z.tgz`, its `SHA256SUMS` and the package content report. No release is made when the commits since the last tag are all release-neutral. Consumers pin the tarball asset URL of a release as the `web-doc` dependency.

## License

Unless explicitly stated otherwise, contributions intentionally submitted for inclusion in Zrimo are licensed under the same `MIT OR Apache-2.0` terms as the project, without additional restrictions.
