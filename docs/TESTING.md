# Testing strategy

Reality check first (§40, §41): full automation of CorelDRAW COM behavior on a standard (non-Windows,
non-CorelDRAW) CI runner is not possible. We use two tiers.

## Tier 1 — Pure-logic unit tests (fully automated, CI-safe)

Modules named `pl_*` contain **zero** CorelDRAW object-model references. They are exercised by the
`modTestRunner.RunAll` harness, which holds explicit `T_Check` assertions and
records pass/fail to the Immediate window and the diagnostics log. VBA has no
reflection, so there is no auto-discovery of test procedures — assertions are
listed explicitly inside `RunAll`.

Run from CorelDRAW VBA IDE:
```vb
modTestRunner.RunAll
```
Output: a console/report and an optional log file written by `modLogger`.

## Tier 2 — Characterization / integration tests (manual)

Core-dependent behaviors cannot be fully automated. Each one lives in
`tests/manual/CHARACTERIZATION.md` as a checklist with:
- precondition document state,
- exact steps,
- expected result,
- pass/fail tick.

These run on a real CorelDRAW install in our manual release gate (§41).

## Static CI checks (Tier 0)

`build/validate.ps1` runs in CI on **any** runner (PowerShell 5+ / pwsh) and requires no CorelDRAW:
- `Option Explicit` presence,
- module-name convention (regex),
- duplicate public procedure detection across modules,
- namespace-prefix lint for public members,
- forbidden hard-coded `C:\`-style absolute paths flagged,
- reference registry consistency (every `Tools → References`-style dependency documented).

## What we do NOT claim
- No automated down-test of CorelDRAW object-model behavior.
- No claim that Corel-dependent modules are "verified" — they are `Validation: Characterization
  (manual)` until run on real installs.
```