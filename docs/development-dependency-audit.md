# Development/build audit follow-up (2026-10-28)

This follows the first security baseline in PR #6, at `7c469f0`.
Full audit improved from **59 (40 high, 12 moderate, 7 low) to 1 low**;
production audit remains **0**. No advisory is suppressed or excluded. Plain full audit still exits nonzero. CI now validates the exact visible residual: only this GHSA, deepmerge 4.3.1,
low severity, unpatched, development-only and known Jest dependency paths. Any
new advisory (including low), changed version/path/severity, or published fix
fails the gate for renewed investigation. The raw report is retained; no pnpm
ignore/exclude option is used. Production audit must be completely clean. A green CI run does not mean a clean full audit.

## Changes and dependency paths

| Dependency path / reason | Patched resolution |
| --- | --- |
| Unused swc-loader -> webpack -> terser-webpack-plugin -> serialize-javascript | Removed swc-loader. No webpack config, import, or build command used it; SWC CLI/core remain. |
| electron-builder -> flatpak-bundler -> tmp-promise -> tmp | 0.2.7 (also fixes the type-confusion bypass in 0.2.6) |
| Jest -> babel-plugin-istanbul -> test-exclude -> glob | 10.5.0; the advisory's fixed 10.x release, not an incompatible glob 12 override |
| Jest/coverage config -> js-yaml 3; builder-util -> js-yaml 4 | 3.15.2 / 4.3.2 |
| Electron ASAR/dir-compare and filelist -> minimatch | 3.1.4 / 5.1.8 |
| Jest haste-map -> anymatch -> picomatch | 2.3.2 |
| Electron flatpak-bundler -> lodash | 4.18.1 |
| macOS signing/plist -> @xmldom/xmldom | 0.8.15 |
| Electron builder schema validation -> ajv -> fast-uri | 3.1.9 |
| Jest/Babel compilation targets -> browserslist | 4.28.7 |

The bounded `pnpm-workspace.yaml` overrides set security floors only on affected
release lines. Inspection of 48 installed dependency edges confirmed the selected
versions satisfy their consumers' original semver ranges. No force install, peer
bypass, cross-major override, audit-ignore list, or registry substitution is used.
The lockfile and added tests prevent reintroducing these vulnerable versions.

Build-only does not mean harmless: packaging reads filesystem paths, archives,
YAML/JSON build configuration and macOS plist data. The glob advisory affects its
shell-invoking CLI; this repository does not call `glob -c`, but the affected
installed version was patched anyway. Linux packaging, test and watch flows were
checked; native Windows/macOS installer smoke remains a platform verification
limit, not a claimed pass. Their build configuration is unchanged.

## Remaining low finding: deepmerge

- **CVE-2026-93753 / GHSA-gcqx-w48m-w2h4**, deepmerge 4.3.1.
- Installed through `jest -> @jest/core -> jest-config -> deepmerge` (and the
  equivalent jest-cli/ts-jest paths); development-only, absent from app.asar's
  production dependencies.
- The [reviewed advisory](https://github.com/advisories/GHSA-gcqx-w48m-w2h4) lists
  **no patched version**. Registry checks at this run returned deepmerge 4.3.1 and
  Jest 30.5.2 as latest. Replacing the test framework or forcing another merge
  implementation is not a safe compatibility patch.
- Source inspection of jest-config shows `deepmerge.all` in its `mergeConfig`
  helper and `deepmerge` in preset/global configuration merging. This project
  supplies a checked-in Jest config and ts-jest preset, no user spreadsheet data
  or application IPC data to either path. There is no application runtime path
  established. Do not execute untrusted Jest configs or presets: they are already
  executable JavaScript with the developer/CI process's privileges.
- This is a documented residual risk, **not a fixed or waived advisory**. Update
  Jest/deepmerge when upstream publishes a compatible fix and rerun full audit.

References: [glob fixed 10.5.0](https://github.com/advisories/GHSA-5j98-mcp5-4vw2),
[tmp fixed 0.2.7](https://github.com/advisories/GHSA-7c78-jf6q-g5cm).

## Verification

A fresh node_modules installation with the frozen lockfile and scripts disabled
succeeded, followed by typecheck, 7 Jest suites / 64 tests, SWC build/watch, Linux
unpacked packaging and actual packaged runtime version verification. Added tests
protect patched lockfile floors and audit fail-closed behavior, including new low
findings, runtime paths, changed versions and an upstream published fix.

The narrow residual policy is not permanent acceptance: an upstream patched
version reported by the advisory immediately invalidates it. Re-evaluate on Jest
or deepmerge upgrades, new affected paths, severity changes, or new audit findings.
