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

They therefore cannot be flattened into this repo's single `Vittix` project without a
module-namespace plan (prefixing every module). The framework treats them as **N
projects**, not as features of one project.

## Registered macros

| id | CorelDRAW project | folder | entry point(s) |
| --- | --- | --- | --- |
| `select-same` | `VittixSelectSame` | `VittixSelectSame` | `mdlMain.SelectSimilarObjects` |
| `bleed` | `vittixBleed` | `vittixBleed` | `Bleeds.Start` |
| `dimension-tools` | `VittixDimensionTools` | `vittixdimension` | `modMain.VittixDimension`, `modMain.VittixDimensionTools_Create` |
| `imposition` | `vittixImposition` | `vittixImposition` | `modImposition.RunImposition` |

## Macro root

Sources are found under a **macro root**. Default: the sibling checkout
`..\vittixcdrMacro`. Override with:

- `build/macros.ps1 -MacroRoot <path>`, or
- the `VITTIX_MACRO_ROOT` environment variable.

When the macro root is absent (e.g. on a CI runner), `validate.ps1` **warns** rather
than failing, so static CI stays green off-machine.

## Commands

```powershell
powershell -File build/macros.ps1            # inventory: modules/forms/entry points per macro
powershell -File build/macros.ps1 -Verify    # exit 1 if any declared file is missing
powershell -File build/macros.ps1 -Json      # machine-readable dump
powershell -File build/validate.ps1          # includes macro-registry drift checks
```

## Known issues in the macro checkout (as of the first inventory)

- **`mdlFormBuilder.bas` exists in three macros at three different revisions**
  (13.4 KB / 17.5 KB / 18.9 KB) — a shared module copy-pasted and diverged.
  Candidate for extraction into one shared module.
- **`vittixImposition` has `RunImposition` (and `BuildPageOrder`, `PlacePageInCell`,
  `DrawCropMarks`) defined in both `ImpositionMacro.bas` and `modImposition.bas`** —
  likely one module is a superseded copy of the other.
- **Uncommitted / untracked work**: the whole `vittixImposition` macro and
  `vittixdimension/src/mdlFormBuilder.bas` are untracked in that repo.
- Per-macro `Sync-To-GMS.ps1` / `Sync-From-GMS.ps1` are near-duplicates differing
  only by the hard-coded project name.
- `Sync-*-GMS.ps1` call `GetActiveObject("CorelDRAW.Application")`, which fails for a
  normally-launched CorelDRAW (it is not in the Running Object Table), so they fall
  back to launching a new instance. See `docs/DEPLOYMENT.md`.
