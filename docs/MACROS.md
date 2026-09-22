# Macros managed by this framework

This repository is the **development framework**. The macros themselves live in one
or more **macro checkouts** with a per-macro layout:

```
<macroRoot>/<macro>/
    src/      *.bas modules, *.cls classes   (git-tracked source of truth)
    forms/    *.frm / *.frx UserForms
    scripts/  Sync-To-GMS.ps1, Sync-From-GMS.ps1
```

`macros/registry.json` is the declarative list of what this framework manages. It is
the single place to look up a macro's **id**, its **CorelDRAW Global Macro Project
name**, its folder, and its **entry points**.

## Why the macros stay as separate projects

CorelDRAW stores one VBA project per `.gms`. The four macros here are four separate
projects, and their module names **collide** across macros:

| Module | Present in |
| --- | --- |
| `mdlFormBuilder.bas` | Select Same, Dimension Tools, Imposition (three *divergent* copies) |
| `modSettings.bas` | Dimension Tools, Imposition |
| `modUnits.bas` / `mdlUnits.bas` | Dimension Tools, Select Same |

They therefore cannot be flattened into this repo's single `Macrix` project without a
module-namespace plan (prefixing every module). The framework treats them as **N
projects**, not as features of one project.

## Registered macros

| id | CorelDRAW project | folder | entry point(s) |
| --- | --- | --- | --- |
| `select-same` | `MacrixSelectSame` | `MacrixSelectSame` | `mdlMain.SelectSimilarObjects` |
| `bleed` | `macrixBleed` | `macrixBleed` | `Bleeds.Start` |
| `dimension-tools` | `MacrixTools` | `macrixTools` | `modMain.MacrixDimension`, `modMain.MacrixTools_Create` |
| `imposition` | `macrixImposition` | `macrixImposition` | `modImposition.RunImposition` |

## Macro root

The macro sources **live in this repository** under `macros\<dir>` — migrated
from the original checkout (`github.com/ketandholakia/Vittix-CDR-Macro`, MIT),
including its uncommitted work. `macros/README.md`, `macros/LICENSE`,
`macros/docs/`, `macros/installer/` and the root `macros/scripts/` are carried over
from that checkout for provenance and tooling.

The macro root defaults to `macros` and can be overridden with
`build/macros.ps1 -MacroRoot <path>` or the `MACRIX_MACRO_ROOT` environment
variable (useful to point at a separate checkout).

When the macro root is absent, `validate.ps1` prints an informational line (not a
warning), so static CI stays green off-machine.

## Commands

```powershell
# inventory / verification
powershell -File build/macros.ps1                        # modules/forms/entry points per macro
powershell -File build/macros.ps1 -Verify                # exit 1 if any declared file/proc is missing
powershell -File build/macros.ps1 -Json                  # machine-readable dump

# validation (scoped per macro so one legacy macro cannot block another)
powershell -File build/validate.ps1                          # framework src + registry drift
powershell -File build/validate.ps1 -Macro dimension-tools   # + lint that macro's src\

# package a macro
powershell -File build/package-macro.ps1 -Macro dimension-tools

# deploy / sync a macro against its CorelDRAW project
powershell -File build/deploy.ps1 -Macro dimension-tools              # dry-run push plan
powershell -File build/deploy.ps1 -Macro dimension-tools -Apply       # push source -> project
powershell -File build/deploy.ps1 -Macro dimension-tools -Pull        # dry-run pull plan
powershell -File build/deploy.ps1 -Macro dimension-tools -Pull -Apply # pull project -> source
powershell -File build/deploy.ps1 -Macro dimension-tools -Stage       # stage for modDevImport
```

Each macro declares `importForms` in the registry: `true` imports its `.frm`, `false`
skips it because the UserForm is generated at runtime by `mdlFormBuilder` (the `.frm`
is VB6-format and CorelDRAW's VBE cannot load it reliably).

## Known issues in the migrated macro sources

- **`mdlFormBuilder.bas` exists in three macros at three different revisions**
  (13.4 KB / 17.5 KB / 18.9 KB) — a shared module copy-pasted and diverged.
  Candidate for extraction into one shared module.
- **`macrixTools`'s `frmDimension` must not be imported from the repo `.frm`.** Its
  controls live in the binary `.frx`, and importing the `.frm` into CorelDRAW's VBE
  produces a control-less form — which makes the whole project fail to compile with
  "Method or data member not found". The form lives in the `.gms`; rebuild it with
  `macros/macrixTools/scripts/Build-Form.ps1` and keep `importForms: false`.
- **`macrixImposition` — RESOLVED.** It had three modules with no `Attribute VB_Name`
  (`ImpositionMacro.bas`, `modImposition.bas`, `modSettings.bas`) and `mdlDebug.bas`
  declaring `VB_Name = "mdlDebugLog"`, plus `ImpositionMacro.bas` defining the same four
  procedures as `modImposition.bas` (an "Ambiguous name detected" compile error). The
  attributes are fixed, the standalone copy now lives in `macrixImposition/legacy/`, and
  the macro deploys with a single macro-list entry: `modImposition.ShowImpositionForm`.
  Its `frmImpositionSettings` is **built at runtime** by `mdlFormBuilder`, because the
  checked-in `.frm` is VB6 format (`Begin VB.Form`) that the VBE cannot import.
- **Source of truth moved here.** The original `macrixcdrMacro` checkout is now
  marked archived **locally** (`ARCHIVED.md` + a README banner, commit `0048ff4`);
  its uncommitted work was deliberately preserved there. Its public remote
  `ketandholakia/Vittix-CDR-Macro` (visibility: public) is **not yet archived**.
- Per-macro `Sync-To-GMS.ps1` / `Sync-From-GMS.ps1` are near-duplicates differing
  only by the hard-coded project name.
- `Sync-*-GMS.ps1` call `GetActiveObject("CorelDRAW.Application")`, which fails for a
  normally-launched CorelDRAW (it is not in the Running Object Table), so they fall
  back to launching a new instance. See `docs/DEPLOYMENT.md`.
