# Vittix CorelDRAW VBA Development System

A small, structured software platform for developing, testing, packaging, installing, maintaining, and
extending CorelDRAW VBA automation — not a grab-bag of unrelated `.bas` macros.

## Status

| Item | Status |
| --- | --- |
| Architecture | Scaffolded (v1) — see `docs/ARCHITECTURE.md` |
| Core services | Stubs + safe-run pattern in place |
| Pure-logic modules | Run via `modTestRunner` harness |
| Corel-dependent code | **Not validated** — requires a live CorelDRAW for characterization tests |
| CI | Static checks only (`build/validate.ps1`); manual checks for COM behavior |

## Repository layout

```
src/
  modBootstrap.bas                Entry points & feature wiring
  modFeatureRegistry.bas          Feature/module registration (§12)
  modConfig.bas · modLogger.bas   Settings · diagnostics (§8, §9)
  modErrorHandler.bas             Safety + failure circuit breaker (§10, §46)
  modVersion.bas · modEnvironment.bas
  modPathManager.bas              Path handling
  modComLifecycle.bas             COM reference release discipline (§37)
  modBinding.bas · modReferenceRegistry.bas · modTrust.bas   (§38, §42)
  modAppServices.bas … core framework services
  modCorelCompatibility.bas · modVersionDetect.bas
  modDialogHelper.bas · modUiUtilities.bas · frmAbout.frm
  feat_*.bas                      End-user features
  pl_*.bas                        Pure-logic, Corel-free modules (§40)
  modTestRunner.bas               Lightweight test harness
build/
    validate.ps1   Static checks (naming, structure, registry drift, paths)
    new-feature.ps1        Scaffolds a new registered feature module
    deploy.ps1     Push/pull src <-> a CorelDRAW .gms (per-macro via -Macro; dry-run default)
    macros.ps1     Inventory/verify the macros registered in macros\registry.json
    package-macro.ps1   Packages one registered macro into build\_out\macros\<id>
    templates/feat_template.bas   Safe-run feature skeleton
    package.ps1    Assembles a versioned source package for import/distribution
    sync_gms_modules.bat   Copies .bas/.frm modules into the CorelDRAW GMS folder
    installer/INNO_SETUP_TEMPLATE.iss
tools/dev-import/modDevImport.bas   In-app bootstrap importer (runs inside CorelDRAW)
macros/registry.json              Declarative list of the managed CorelDRAW macro projects
macros/<macro>/{src,forms,scripts}  Macros themselves (migrated; one folder per .gms project)
.github/workflows/validate.yml   CI: runs build/validate.ps1 on push/PR
docs/DEPLOYMENT.md               How source reaches a CorelDRAW .gms project
docs/MACROS.md                   The macros this framework manages + their layout
docs/
    ARCHITECTURE.md  CONVENTIONS.md  ADDING_A_MACRO.md  TESTING.md  RELEASE_CHECKLIST.md
```

## Development workflow

The end-to-end macro lifecycle (scaffold → implement → validate → test → package
→ deploy → release) is in `docs/WORKFLOW.md`. The short version:

```powershell
# 1. Scaffold a new macro (creates src/feat_MyTool.bas and registers it)
powershell -File build/new-feature.ps1 -Name my-tool -Display "My Tool" -Stability Beta -Register

# 2. Static gate — must report 0 errors
powershell -File build/validate.ps1
```

`build/validate.ps1` also fails on **registry drift**: a `src/feat_*.bas` that is
not registered, or a registered module that doesn't exist.

## Quick start

See `docs/ADDING_A_MACRO.md` for how to add a new feature with minimal repetition.

## Build & validate

```powershell
powershell -File build/validate.ps1        # cross-platform static lint (safe for CI)
powershell -File build/package.ps1         # assemble versioned package into build/_out
```

### Deploy modules to CorelDRAW

```batch
build\sync_gms_modules.bat                  # auto-detect GMS folder and copy .bas/.frm files
build\sync_gms_modules.bat "C:\path\to\GMS" # copy to an explicit GMS folder path
```

The script auto-detects the CorelDRAW GMS folder from `%APPDATA%\Corel` (or
`%LOCALAPPDATA%\Corel` as fallback), or you can override via the
`VITTIX_GMS_DIR` environment variable. See the script header for exit codes.