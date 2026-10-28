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
- Jest: 7 suites, 64 tests, including real XLS and XLSX import, header detection,
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

`pnpm audit --json`: **1 low finding**.
`pnpm audit --prod --json`: **0 findings**.
These are advisory counts, not 1/0 distinct vulnerable packages or demonstrated
exploits. The full audit returns nonzero; the production audit passes. No findings were hidden or ignored.

Unused production dependencies removed after source, dynamic-loading, HTML and build-configuration checks:

- `i18next-node-fs-backend` and `i18next` were never imported or initialized. The
  Korean interface and locale JSON resources remain unchanged. Their unused
  json5/js-yaml dependency path is no longer shipped. README no longer claims an
  implemented language switch.
- `electron-store` was never imported or used; its conf/ajv/fast-uri production
  path is no longer shipped. Packaged app.asar metadata confirms SheetJS is the
  only production dependency. No settings storage feature was removed.
- CI now requires the production audit to pass; the full audit rejects every advisory except the exact documented unpatched
  low Jest development path, retained visibly in its report.

The development/build follow-up removed unused swc-loader and patched remaining
compatible dependency lines. Full audit is now 1 low (deepmerge through Jest,
upstream has no patched release); production remains 0. See
[dependency paths, patched versions, reachability and residual risk](development-dependency-audit.md).
No findings were suppressed. CI reports every full-audit finding and validates only the exact current
unpatched Jest development risk; any new advisory/version/path/severity or
published fix requires renewed investigation. Production audit must be clean.
Plain full audit remains nonzero.

Historical comparison using the same audit service: original repository baseline
167 full / 21 production -> initial PR 59 / 0 -> current 1 / 0. This is not a
clean full audit or comprehensive coverage of CDN SheetJS/Chromium/Node.
