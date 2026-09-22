# Changelog

This project uses [Semantic Versioning](https://semver.org/) (`MAJOR.MINOR.PATCH`) for the platform
and per-feature where features version independently. Every user-visible behavior change requires a
CHANGELOG entry here, not just internal refactors (§43).

## [Unreleased]

### Added — developer tooling & CI
- **`build/new-feature.ps1`** — scaffolds a new feature module from
  `build/templates/feat_template.bas` (safe-run skeleton pre-wired), reserves an
  unused `Err.Raise` base, and with `-Register` inserts the `FR_Register` line
  into `modFeatureRegistry.FR_RegisterFeatures()` automatically. Prints the
  characterization-test row and CHANGELOG stub to add.
- **`build/validate.ps1`** — added checks: `Attribute VB_Name` present and
  matching the file; balanced `Sub`/`Function`/`Property` blocks;
  leading-underscore identifiers (uncompilable — the historical bug); and
  **feature-registry drift** (every `src\feat_*.bas` is registered, and every
  registered module exists).
- **`.github/workflows/validate.yml`** — CI runs `build/validate.ps1` on push/PR.
- **`.gitattributes`** — pins VBA sources (`.bas`/`.cls`/`.frm`) to CRLF; docs and
  CI files to LF.
- **`docs/WORKFLOW.md`** — the end-to-end macro development loop.

### Added — per-macro pipeline (dimension-tools first)
- **`build/deploy.ps1` now takes `-Macro <id>`** and resolves the target from the
  registry: source dir, `forms` handling, and the CorelDRAW project name. It does
  both directions -- push (`-Apply`, default) and pull (`-Pull [-Apply]`) -- with a
  dry-run plan by default. `.frm` import is controlled by each macro's
  `importForms` flag. Unknown ids exit `6`.
- **`build/package-macro.ps1`** -- packages a registered macro (modules, forms,
  README) into `build\_out\macros\<id>` with a `MANIFEST.txt`. Output lives under
  the gitignored `build\_out` tree, so it no longer writes into the source folders.
  Fixes the upstream manifest defect (`System.Object[]`).
- **`build/validate.ps1 -Macro <id>`** -- lints one macro's `src\` (structure +
  `Attribute VB_Name`); scoped so a broken legacy macro cannot block work on
  another. Host document modules (`ThisMacroStorage`, ...) are exempt from
  `Option Explicit`.
- Registry: added a per-macro `importForms` flag.
- Surfaced (not yet fixed) by the new lint: `vittixImposition` has three modules
  with **no `Attribute VB_Name`** and `mdlDebug.bas` mis-declares its own name.

### Added — macro sources migrated into this repo
- Migrated the four macro projects from the standalone `vittixcdrMacro` checkout
  into `macros/` (`VittixSelectSame`, `vittixBleed`, `vittixdimension`,
  `vittixImposition`), together with their per-macro sync scripts, the shared
  docs/installer/scripts, and the two demo `.cdr` files. 61 files, every one
  hash-verified against the source; the migration included the previously
  **uncommitted/untracked** work (notably the whole `vittixImposition` macro).
  Generated output (`build/`, `release/`, `dist/`) and VCS/editor folders were
  excluded.
- `macros/registry.json` now points at the in-repo copies (`macroRoot: "macros"`).
- `.gitattributes`: `.cdr` treated as binary.

### Added — multi-macro support
- **`macros/registry.json`** — declarative registry of the CorelDRAW macro projects
  this framework manages (id, Global Macro Project name, folder, entry points,
  stability). First entries: `select-same`, `bleed`, `dimension-tools`, `imposition`.
- **`build/macros.ps1`** — inventory + verification of the registered macros
  (`-Json`, `-Verify`). Resolves the macro root from `..\vittixcdrMacro` by default,
  overridable with `-MacroRoot` / `VITTIX_MACRO_ROOT`.
- **`docs/MACROS.md`** — the macro model, why the macros stay separate VBA projects
  (module-name collisions across macros), and the first inventory's findings.
- `build/validate.ps1` check 9: macro-registry drift vs the macro checkout. Warns
  (does not fail) when the macro root is absent, so CI stays green off-machine.

### Added — deployment to CorelDRAW
- **`build/deploy.ps1`** — imports the repo's VBA modules into a CorelDRAW `.gms`
  project. Dry-run by default (`-Apply` to write); validates first; never quits a
  CorelDRAW instance it did not start. `-List` enumerates loaded GMS projects and
  their macros. Grounded in the CorelDRAW 2021 automation model, verified from the
  installed type library: `Application` exposes `VBE`, `GMSManager`
  (`Projects.Load/Unload`, `RunMacro`).
- **`tools/dev-import/modDevImport.bas`** — in-app bootstrap importer: import it into
  the project once, then run `DevImport_RunAll` inside CorelDRAW to pull every staged
  module in. Deliberately outside the staged set so it never re-imports itself.
- **`docs/DEPLOYMENT.md`** — the two deployment paths, prerequisites, saving semantics.
- `build/validate.ps1` now also lints `tools/dev-import` (structural checks only).

### Fixed
- **`new-feature.ps1` corrupted non-ASCII files.** It read source with PowerShell's
  default encoding and rewrote UTF-8, so characters in a file it touched were
  double-encoded. This mangled `modFeatureRegistry.bas`'s header (`—` → `â€“`,
  `§` → `Â§`) during testing. Reads are now explicit UTF-8. Separately, non-ASCII
  characters inside `.ps1` string literals broke parsing under Windows PowerShell
  5.1; the build scripts are now ASCII-only, and the corrupted registry header was
  restored. Both shells parse all four scripts cleanly.
- **Project failed to compile at all** — every private helper across the codebase (`modLogger`,
  `modVersion`, `modErrorHandler`, `modFeatureRegistry`, `feat_RoundedCorners`, `feat_ExportText`) was
  named with a leading underscore (`_openFile`, `_Split`, `_getCount`, `_setCount`, `_isDisabled`,
  `_disable`, `_index`, `_SelectionRound`, `_WriteDocument`) per the (incorrect) convention in
  `CONVENTIONS.md`. VBA identifiers must start with a letter, so every one of these declarations —
  and every call site — was a syntax error. Since VBA compiles the whole project as a unit, this blocked
  *every* macro, including unrelated ones like `modTestRunner.RunAll`. Renamed all of them to a `p_`
  prefix and fixed `CONVENTIONS.md`/`ADDING_A_MACRO.md` so the mistake isn't reintroduced.
- `feat_ExportText.Main` called `modExportImportServices.SVE_BuildExportPath` and
  `_WriteDocumentText`, neither of which existed (`modExportImportServices` was never built; the real
  private sub was named `_WriteDocument`, and did nothing but write a hardcoded placeholder line,
  ignoring the document entirely). Added `modExportImportServices` (`SVE_BuildExportPath`) and
  rewrote the write path (now `p_WriteDocumentText`) to actually enumerate shapes across all pages
  and export each text shape's `Text.Story`, with real file-IO failures propagating to the caller
  instead of being swallowed.
- `docs/TESTING.md` claimed the runner "calls each `test_*` procedure". VBA has no
  reflection, so no such auto-discovery exists; the assertions are listed
  explicitly in `modTestRunner.RunAll`. Corrected the doc to match reality.
- `modFeatureRegistry.FR_RegisterFeatures` was not idempotent — calling it more than once per session
  (e.g. re-running `modBootstrap.Main` while testing) duplicated every registered feature. It now
  resets its table before re-registering.
- `modVersionDetect.VD_ParseCorelVersion` stripped all non-digit characters and concatenated what was
  left into one number (`"2021.0.0"` → 202100, `"18.0.0.448"` → 1800448), so version ordering only
  held by coincidence of digit-string length. It now splits on `.` and weights each segment
  independently, matching `modVersion.VER_Compare`'s semantics instead of a separate, less reliable
  scheme.
- Added regression tests to `modTestRunner.RunAll` for the version-parsing and registry-idempotency
  fixes above, so they can't silently regress again.

### Added — greenfield bootstrap
- Scaffolded platform per the target architecture: Core Framework, Infrastructure, UI, Compatibility,
  Security, Development/Build tooling.
- Foundation modules:
  - Feature registry (`modFeatureRegistry`) with registration, availability gating, and session
    circuit breaker.
  - Safe-run pattern (inlined per feature — see `docs/CONVENTIONS.md`) for uniform error handling and
    state restoration on every exit path.
  - Pure-logic test harness (`modTestRunner`) + `pl_*` modules.
  - Static validation entry point (`build/validate.ps1`).
- Placeholder UserForm `frmAbout.frm`.

### Deprecation policy
- No features shipped yet; once the first real feature lands, this section documents the
  deprecate → warn → remove timeline (§43).

## Format

- `Added` / `Changed` / `Deprecated` / `Removed` / `Fixed` / `Security`.
- Reference feature IDs from the registry and commit messages (`feat(plate):`, `fix(module):`) per §47.