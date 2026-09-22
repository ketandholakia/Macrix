# Changelog

This project uses [Semantic Versioning](https://semver.org/) (`MAJOR.MINOR.PATCH`) for the platform
and per-feature where features version independently. Every user-visible behavior change requires a
CHANGELOG entry here, not just internal refactors (§43).

## [Unreleased]

### Fixed — macrixImposition "Sub or Function not defined"
- Self-inflicted regression: I made `RunImposition` **Private**, but **`Private` in VBA is
  module-scoped**, so the form's `cmdRun_Click` could no longer call it → *"Compile error:
  Sub or Function not defined"*. `Option Private Module` is the *project*-scoped equivalent:
  it hides a module from the macro list without breaking in-project calls.
- Restructured to match dimension-tools: new **`modMain.bas`** holds the single public entry
  `ShowImpositionForm`; `modImposition` is now `Option Private Module` with a public-but-hidden
  `RunImposition`. Macro list = 1 → `modMain.ShowImpositionForm` (registry updated).
- Lesson recorded in `docs/DEPLOYMENT.md`: a `.gms` is flushed by CorelDRAW **~40–60 s after**
  the instance exits, and *every* instance rewrites the `.gms` it loaded on exit — so verify
  only after the write settles, or a stale instance can clobber it.

### Fixed — macrixImposition runtime error 424
- `ShowImpositionForm` died with **"Run-time error '424': Object required"**. The runtime
  form builder did `Set newFormComp = vbProj.VBComponents.Add(3)` and then, unguarded,
  `newFormComp.Name = FORM_NAME` — when `Add` returns Nothing, that next line raises 424.
- The dialog is now a **real design-time form**. `macros/macrixImposition/scripts/Build-Form.ps1`
  builds `frmImpositionSettings` (26 controls + the 190-line code-behind, **extracted from
  `mdlFormBuilder.bas`** so the two stay in step) and `ShowImpositionForm` now just calls
  `frmImpositionSettings.Show`. No runtime VBE access, no form generation.
- `BuildImpositionFormWithCode` still exists but now guards the `Add` and reports clearly
  instead of raising 424.

### Added — macrixImposition built
- **Fixed the compile blockers:** `modImposition.bas`, `modSettings.bas` and
  `ImpositionMacro.bas` had **no `Attribute VB_Name`**, and `mdlDebug.bas` declared
  `VB_Name = "mdlDebugLog"` — the VBE would have created mis-named components.
- **Resolved "Ambiguous name detected":** `ImpositionMacro.bas` was a standalone copy
  defining the same four procedures as `modImposition.bas`; it moved to
  `macros/macrixImposition/legacy/` (with a note) and `src/` now holds only the modular set.
- **Macro-list hygiene** (same as dimension-tools): `RunImposition`, `BuildPageOrder`,
  `PlacePageInCell`, `DrawCropMarks` are now Private and `Option Private Module` was added
  to `modSettings`/`mdlFormBuilder`/`mdlDebug`, leaving one entry:
  **`modImposition.ShowImpositionForm`**.
- **Hardened the form builder:** `Application.VBE.ActiveVBProject` is usually Nothing at
  runtime, so the new `TargetProject()` finds the project that owns `mdlFormBuilder`
  instead — otherwise the runtime form build failed silently.
- Deployed into `macrixImposition.gms`; verified 5 components and a 1-entry macro list.
  The settings dialog is built at run time (the checked-in `.frm` is VB6 format).

### Changed — Vittix → Macrix rebrand finished for the live projects
- CorelDRAW projects renamed in the GMS folder: `VittixDimensionTools.gms` →
  **`MacrixTools.gms`**, `VittixSelectSame.gms` → `MacrixSelectSame.gms`,
  `vittixBleed.gms` → `macrixBleed.gms`, `vittixImposition.gms` → `macrixImposition.gms`.
  The name stored *inside* each `.gms` was updated as well — the project name lives in the
  file, not merely in the filename.
- `MacrixTools`: entry point is **`modMain.MacrixDimension`** (macro list = 1 entry), the
  dialog caption is set at design time and re-asserted in `UserForm_Initialize`, and
  settings persist under the **`Macrix`** key.
- Added `build/rebrand-live.ps1` (guarded: refuses to run while CorelDRAW is open).
- Note: `vittix.gms` (the original framework scaffold) was left untouched.

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

### Changed — dimension-tools: one entry only in CorelDRAW's macro list
- Added `Option Private Module` to `modDimension`, `modLabel`, `modLayers`,
  `modSettings`, `modUnits` and `mdlFormBuilder` (`modGeometry`/`modUtils` already had it).
  Their public procedures still work everywhere inside the project but no longer appear in
  CorelDRAW's macro list.
- `modMain.VittixDimensionTools_Create` is now **Private** (it was a second listing), and
  the dialog's OK button calls `modDimension.CreateDimensionsForSelection` directly.
- Verified against the project itself: `GMSProject.Macros.Count = 1` → `modMain.VittixDimension`.

### Fixed/Added — dimension-tools: colour actually applies, plus a swatch picker
- **Fixed: a hand-typed colour had no effect.** `ApplyTextFormatting` ran
  `story.Font = "Arial"` (an unsupported member on a TextRange) and shared a single error
  handler, so that failure jumped past the colour step. Formatting is now three
  independent, individually-guarded steps (`ApplyTextColor` first, then `ApplyStoryFont`,
  then `ApplyStorySize`), so one unsupported member can no longer skip the rest.
- **Added a colour picker:** a 2×12 palette of clickable swatches. Each swatch carries its
  hex in `Tag`; clicking one writes it into the Text Color field, which previews live. The
  24 click handlers are generated by `Build-Form.ps1`.
- No colour-dialog API exists on the CorelDRAW `Application` object (checked the type
  library), so the picker is built into the form rather than calling an OS dialog.
- Form rebuilt: **54 controls**, 330×540 pt; stray `UserForm1` artefacts from failed
  designer builds removed. `Build-Form.ps1` also coerces MSForms geometry to `Single`
  (passing a Double throws "Specified cast is not valid").

### Added — dimension-tools: caption colour + OK default button
- **Text Color** setting: a `#RRGGBB` (also accepts `R,G,B` and `#RGB`) field, with a live
  colour swatch beside it that previews as you type. The colour is applied to the label's
  text fill in `modLabel.ApplyTextColor` and persisted with the other settings.
- New helpers `ColorFromHex` / `ColorToHex` in `modUtils`.
- **OK is now the default button** (Enter) and **Cancel answers Esc**, set on the controls
  and re-asserted in `UserForm_Initialize`.
- Form rebuilt: **30 controls**, 330×486 pt, re-exported to `forms/frmDimension.frm`.

### Changed — dimension-tools: Font Size dropped, Text Width % now works
- **Removed the Font Size setting and control.** Labels are sized solely by
  `Text Width %` — a percentage of the selected object's width. `VDT_Settings.fontSize`,
  the `txtFontSize`/`lblFontSize` controls, and the persist/apply/read paths are gone.
- **Root cause of "neither works":** the code set `Text.Story.FontSize`, `.FontName` and
  `.TextRange.FontSize`. None of those members exist — a TextRange exposes `Font` and
  `Size` — and `On Error Resume Next` swallowed the failure, so nothing happened at all.
  Verified empirically against CorelDRAW 2021: setting `Story.Size` changes the shape's
  width exactly as expected, while `Text.FontProperties.Size` is not settable.
- `AppliedTextFormatting` now uses `Story.Font` / `Story.Size`; `ScaleTextToWidth` computes
  `newSize = currentSize * (targetMM / currentWidthMM)` from `Shape.SizeWidth` (the
  document unit is already millimetres during label creation), clamped to 0.5–2000 pt,
  and aborts with a logged error instead of silently doing nothing.
- Form rebuilt without the Font Size row (**27 controls**, 330×456 pt) and re-exported;
  a stray `UserForm1` left by a failed form build was removed.
- `Build-Form.ps1`: documents that removing and re-adding a form in one session can fail
  with a path/file access error (re-run succeeds then).

### Changed — dimension-tools label placement
- **New `Center` position**: `Label_PositionShape` now has a `center` case that puts the
  label at the object's centre, and the form's Position list is
  `Auto, Center, Above, Below, Left, Right`.
- **`Left`/`Right` now read vertically**: the label is rotated 90 degrees and its centre
  placed on that edge (`Shape.Rotate 90` applied before the move).
- Updated `modLabel.bas`, `mdlFormBuilder.bas` (injected form code) and
  `macros/vittixdimension/scripts/Build-Form.ps1`; re-exported `forms/frmDimension.frm`.

### Fixed — dimension-tools compile error (VBA 461)
- The deployed `frmDimension` was a **control-less** form (the repo `.frm` has no inline
  control definitions; they live in the binary `.frx`), while `modDimension` and the
  form's own code referenced `cmbUnit`, `cmbDecimals`, `cmbPosition`, `txtFontSize`,
  `txtTextWidthPercent`, `txtGap`, `txtPadding`, `txtCornerRadius`, `chk*`, `txtTemplate`
  and `LoadFormState`. Every one is *"Method or data member not found"*, and VBA
  compiles the whole project as a unit, so nothing ran.
- **Rebuilt `frmDimension` as a real design-time form** (29 controls + code-behind, via the
  VBIDE Designer API) and simplified `modDimension.ShowDimensionForm` to
  `frmDimension.Show`, removing the runtime VBE dependency (the old runtime builder could
  never run, because the project did not compile).
- Added **`macros/vittixdimension/scripts/Build-Form.ps1`** to reproduce the form.
- `deploy.ps1` gained **`-RemoveComponents <names>`** (clear stale components) and
  **`-RunMacro <Module.Macro>`** (`GMSManager.RunMacro(ModuleName, MacroName, Parameters)`).
- Verified in the `.gms`: components `ThisMacroStorage, frmDimension, mdlFormBuilder` + the
  9 modules; `frmDimension` has 29 controls and `LoadFormState`; `modDimension` uses the
  direct form path.
- Follow-up (reported as "form is clipped"): the rebuilt form kept the default **240x180 pt**
  size while its controls span 300x441, and its `Caption` was still the default
  `UserForm1`. The form is now **330x486 pt** with `Caption = "Vittix Dimension"`. The form
  property setters take **string** values (an Int32 is rejected), and `Width`/`Height` are
  outer (points) while control `Left`/`Top` are inner coordinates.
- Follow-up (reported as "Compile error: Invalid optional parameter type" after pressing OK):
  `modLabel.CreateArtisticTextShape` declared `Optional settings As VDT_Settings`, and
  **VBA forbids a user-defined Type as an `Optional` parameter**. Reordered to
  `settings As VDT_Settings, Optional objectWidthMM As Double = 0` (a required parameter
  may not follow an optional one) and updated its single caller.
- `build/validate.ps1` gained a lint rule (6b) for exactly this: an `Optional` parameter
  typed as a user-defined Type. Verified against a synthetic bad module.
- **Duplicate-component bug in `deploy.ps1`**: the push loop captured a COM object across
  the `VBComponents` enumerator, which could go invalid, so `Remove` silently failed and
  VBE appended `1`/`2` (creating `modMain1`, `modUnits1`, `modUtils1`). The loop now removes
  by index and also clears any suffixed duplicates. Verified: 11 components, 0 duplicates.

### Added — binary .gms build (dimension-tools)
- **`build/deploy.ps1 -OutDir <dir>`** copies the rebuilt `.gms` out as a build
  artifact after a push.
- **Fixed:** `deploy.ps1` now calls `Application.InitializeVBA()` before touching
  `Application.VBE`. A freshly launched automation instance exposes **no** VBA project
  model until then (`VBE` was `$null`), so deploying against a new instance would have
  failed with "VBE not accessible".
- `VBProject.SaveAs` is not supported for `.gms`; CorelDRAW persists the project when
  the instance closes, and `deploy.ps1` now quits-then-verifies instead of reporting a
  false failure. Re-runs with unchanged sources are byte-identical.
- **Verified on CorelDRAW 2021 (v23.5.0.506):** rebuilt `VittixDimensionTools.gms`
  (115730 → 112146 bytes, magic `GMS\x01`) and confirmed the artifact re-loads with
  all 9 modules plus the preserved `frmDimension`/`UserForm1`.
- `docs/DEPLOYMENT.md`: new "Building a binary `.gms`" section.

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