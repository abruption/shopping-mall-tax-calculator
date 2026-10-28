# Security baseline validation (2026-10-28)

Scope: issue #5. This is a draft change, not a release or an assertion that the
application has no vulnerabilities.

## Resolved baseline

- SheetJS CE 0.20.3 from the official `cdn.sheetjs.com` tarball; SHA-512 integrity
  is recorded in the authoritative `pnpm-lock.yaml`. Covers the fixed floors for
  CVE-2023-30533 and CVE-2024-22363.
- Electron 44.10.0, verified from both downloaded and Linux unpacked binaries:
  Node 24.22.1, Chromium 152.0.7977.169. macOS 13+ is now required.
- All resolved brace-expansion instances are 1.1.21, 2.1.7, or 5.0.12. A scoped
  1.x override preserves legacy minimatch compatibility; no cross-major override.
- electron-builder 26.18.1 replaces its old git-hosted node-gyp dependency chain
  with registry dependencies. No supply-chain policy was disabled. pnpm added a
  version-specific release-age exception for the explicitly selected Electron
  release. Dependency install scripts were initially skipped for review; only
  Electron's official installer was explicitly run to obtain its checked binary.
- `pnpm-lock.yaml` is the single lockfile, matching the existing pnpm scripts.
  SWC CLI is explicitly pinned to 0.8.1: its official downloader chain uses
  patched file-type 21.3.4+, avoiding a development-only decompression advisory
  introduced by the 0.7.x dependency refresh. Both build and watch CLI paths were
  tested. Other dependency changes remain within their existing major ranges.

## Checks

Passed on Linux / Node 24.22.1:

- Frozen lockfile installation (`--ignore-scripts`).
- TypeScript strict typecheck (`tsc --noEmit`). Existing invalid global declaration,
  unknown catch values, implicit callback type and nullable filter types were
  repaired without weakening compiler settings.
- Jest: 5 suites, 36 tests, including real XLS and XLSX import, header detection,
  calculation, XLSX/CSV export, and literal markup-shaped filename rendering.
  Two old mocked expectations were aligned with the existing null-on-invalid-date
  behavior and required explicit column options, not a behavior change.
- SWC build and execution of the built renderer as a browser script in jsdom;
  markup-shaped filenames remain text, no image/event-handler node is created.
- `electron-builder --dir --linux --config.electronDist=node_modules/electron/dist`.
  The downloaded distribution is reused because the container's home cache is
  read-only. No installer or release artifact was published.
- Packaged executable `ELECTRON_RUN_AS_NODE=1 ... -p process.versions` confirms the
  runtime versions above, not merely the build host version.

Not passed / not run:

- GUI smoke launch: headless Electron terminated in this restricted container;
  D-Bus socket access was denied, display/Xvfb unavailable, and font cache paths
  were not writable. No sandbox disabling or platform-permission bypass used.
- Native macOS/Windows installer and interactive file-dialog smoke tests require
  those platforms. Verify before distributing a release; Ready status alone does not certify these platforms.
- No repository lint command existed. A new read-only CI workflow repeats the
  typecheck, tests, build, Linux packaging/runtime verification and separate audits.

## Separate dependency audit results

`pnpm audit --json`: **60 findings** (41 high, 12 moderate, 7 low).
`pnpm audit --prod --json`: **18 findings** (12 high, 4 moderate, 2 low).
These are advisory counts, not 60/18 distinct vulnerable packages or demonstrated
exploits. Both audit commands return nonzero. No findings were hidden or ignored.

Production paths still needing a separately tested migration:

- `i18next-node-fs-backend@2.1.3` pins `json5@2.0.0` and `js-yaml@3.13.1`.
  The backend is deprecated; upstream recommends `i18next-fs-backend`. A blind
  override of exact dependency pins or backend/API migration is not included.
- `electron-store -> conf -> ajv[/ajv-formats] -> fast-uri@3.0.6` remains flagged
  for URI parsing/normalization advisories. Application exploitability was not
  demonstrated; do not interpret a library alert as proof of SSRF here.

Development audit additionally reports tmp, glob, webpack, minimatch,
serialize-javascript, picomatch, lodash, @xmldom/xmldom, browserslist,
js-yaml and deepmerge. These require upstream dependency-path upgrades and/or
compatibility work. The named brace-expansion findings are resolved; a general
major-version audit fix was deliberately not forced. CDN SheetJS advisories and
Electron runtime advisories must be checked upstream because package audit alone
is not complete coverage.

Comparison against the unchanged baseline lockfile using the same audit service:
full 167 (5 critical / 90 high / 54 moderate / 18 low) -> 60; production 21
(0 critical / 14 high / 5 moderate / 2 low) -> 18. No new advisory IDs remain
relative to baseline. This reduction is not a clean audit or proof of reachability.
