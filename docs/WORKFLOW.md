# Macro development workflow

The day-to-day loop for adding and maintaining macros in this framework. Pairs
with `docs/ADDING_A_MACRO.md` (rules) and `docs/RELEASE_CHECKLIST.md` (shipping).

## 0. One-time setup
- `git init` (or clone), then `powershell -File build/validate.ps1` should be clean.
- Locate the CorelDRAW GMS folder: `build\sync_gms_modules.bat` auto-detects it,
  or set the `VITTIX_GMS_DIR` environment variable to override.

## 1. Scaffold a macro
```powershell
powershell -File build/new-feature.ps1 -Name my-tool -Display "My Tool" -Stability Beta -Register
```
Creates `src\feat_MyTool.bas` from `build\templates\feat_template.bas` (already
wired with the safe-run pattern) and, with `-Register`, inserts the
`FR_Register` line into `modFeatureRegistry.FR_RegisterFeatures()`. Without
`-Register` it prints the exact line to paste. The script also prints the
characterization-test row and CHANGELOG stub to add.

## 2. Implement
- Keep `Main` thin; put the real work in the private `p_DoWork`.
- Reach the object model through `mod*Services` where one covers it; otherwise
  use `CallByName` / `modBinding` directly and extract a service once a second
  feature needs the same access.
- Release COM references inside loops (`modComLifecycle.CM_Release`); never cache
  a `Shape`/`Selection` across an undo/delete/ungroup.
- The template already restores `Optimization`/`EventsEnabled` on both exit paths.

## 3. Validate (Tier 0 — no CorelDRAW needed)
```powershell
powershell -File build/validate.ps1
```
Checks: `Option Explicit`; `Attribute VB_Name` matches the file; balanced
`Sub`/`Function`/`Property` blocks; duplicate public procedures; leading-underscore
(uncompilable) identifiers; hard-coded absolute paths; and **registry drift** —
every `src\feat_*.bas` is registered, and every registered module exists.

## 4. Test
- Pure logic → add assertions to `modTestRunner.RunAll`; put helper logic in a
  `pl_*` module (zero CorelDRAW references).
- Corel-dependent → add a row to `tests/manual/CHARACTERIZATION.md` and run it on
  a real CorelDRAW install. This is the manual release gate (§41).

## 5. Document
- Add a CHANGELOG entry under `[Unreleased]` for any user-visible change.
- Keep the module header comment (`Depends on:`, CorelDRAW-version assumptions) accurate.

## 6. Package & deploy
```powershell
powershell -File build/package.ps1        # versioned zip into build\_out

# push modules into the live CorelDRAW VBA project (.gms):
powershell -File build/deploy.ps1 -List   # inspect loaded projects + macros
powershell -File build/deploy.ps1         # dry run (validates + prints plan)
powershell -File build/deploy.ps1 -Apply  # import into the target project
```
CorelDRAW's `.gms` is a **binary** project, so source can't be linked directly — it
must be imported (via `Application.VBE` from `deploy.ps1`, or via the in-app bootstrap
`modDevImport.DevImport_RunAll`). Full mechanism, prerequisites, and the recommended
bootstrap path: **`docs/DEPLOYMENT.md`**.

## 7. Release
Follow `docs/RELEASE_CHECKLIST.md` (functional gate, safety/install, docs, review,
version bump).

## What the framework gives each macro
- A registry entry with an **independent version and stability** (Stable / Beta /
  Experimental / Deprecated) and availability gating.
- A per-feature **session circuit breaker** — auto-disables after N failures in a
  session (default 3) instead of failing repeatedly.
- **Uniform error reporting and logging** through `modErrorHandler` / `modLogger`.
- **Local diagnostics only** — no telemetry (`CFG_TelemetryAllowed` is False by design).
