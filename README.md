# Macrix CorelDRAW Macro Development

Development repository for a set of **CorelDRAW VBA macros**. Each macro is its own
CorelDRAW VBA project (its own `.gms`); this repository holds the macro sources, the
registry that tracks them, and the tooling to lint, package and sync them into
CorelDRAW.

## Macros

| Macro | CorelDRAW project | Sources | Entry point(s) | Forms |
| --- | --- | --- | --- | --- |
| Macrix Select Same | `MacrixSelectSame` | `macros/MacrixSelectSame/` | `mdlMain.SelectSimilarObjects` | imported |
| Macrix Bleed | `macrixBleed` | `macros/macrixBleed/` | `Bleeds.Start` | imported |
| Macrix Dimension Tools | `MacrixTools` | `macros/macrixTools/` | `modMain.MacrixDimension`, `modMain.MacrixTools_Create` | built at runtime |
| Macrix Imposition | `macrixImposition` | `macros/macrixImposition/` | `modImposition.RunImposition` | built at runtime |

The macros are four separate VBA projects whose module names collide, so they stay
as four projects rather than being flattened into one. Per-macro details and known
issues: [`docs/MACROS.md`](docs/MACROS.md).

## Everyday commands

```powershell
# inventory / verify the registry against the sources
powershell -File build/macros.ps1 -Verify

# lint one macro (scoped, so one macro cannot block another)
powershell -File build/validate.ps1 -Macro dimension-tools

# package one macro
powershell -File build/package-macro.ps1 -Macro dimension-tools

# sync with CorelDRAW (dry run by default; -Apply to write)
powershell -File build/deploy.ps1 -Macro dimension-tools               # plan a push
powershell -File build/deploy.ps1 -Macro dimension-tools -Apply        # source -> .gms
powershell -File build/deploy.ps1 -Macro dimension-tools -Pull -Apply  # .gms -> source
```

Full lifecycle: [`docs/WORKFLOW.md`](docs/WORKFLOW.md). How source reaches a `.gms`:
[`docs/DEPLOYMENT.md`](docs/DEPLOYMENT.md).

## Layout

```
macros/                          the macros (one folder per CorelDRAW VBA project)
  registry.json                  id, CorelDRAW project name, folder, entry points, importForms
  <macro>/{src,forms,scripts}
build/                           tooling
  macros.ps1                     inventory / -Verify / -Json
  validate.ps1                   static checks (+ -Macro <id> to lint one macro)
  package-macro.ps1              package one macro into build\_out\macros\<id>
  deploy.ps1                     push/pull/stage a macro's code to/from CorelDRAW
  new-feature.ps1                scaffold a module for the shared scaffold project (src/)
  package.ps1                    package the shared scaffold project (src/)
  templates/feat_template.bas    safe-run module skeleton
tools/dev-import/modDevImport.bas  in-CorelDRAW bootstrap importer
src/                             OPTIONAL shared scaffold (see below)
docs/                            see docs/README.md
.github/workflows/validate.yml   CI: static validation
```

## Optional shared scaffold (`src/`)

`src/` is a single-project VBA **scaffold**: safe-run entry pattern, feature
registry, session circuit breaker, COM-lifecycle helpers, a pure-logic test harness,
and static validation. **No shipped macro depends on it yet** — it is kept as
reference material and a starting point for new macro code.

Its documentation: `docs/ARCHITECTURE.md`, `docs/CONVENTIONS.md`,
`docs/ADDING_A_MACRO.md`, `docs/TESTING.md`, `docs/RELEASE_CHECKLIST.md`.

```powershell
powershell -File build/validate.ps1        # static lint (safe for CI, no CorelDRAW needed)
powershell -File build/package.ps1         # assemble a versioned package of src/
```

## Provenance

The macros were migrated from `github.com/ketandholakia/Vittix-CDR-Macro` (that
checkout has been archived). Migration details and mapping: [`docs/MACROS.md`](docs/MACROS.md).
