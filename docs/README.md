# Docs index

## Macros (primary)

- [MACROS.md](MACROS.md) — the managed macros: registry, layout, entry points, known issues.
- [WORKFLOW.md](WORKFLOW.md) — the end-to-end macro lifecycle (scaffold → lint → package → deploy → release).
- [DEPLOYMENT.md](DEPLOYMENT.md) — how source gets into a CorelDRAW `.gms`, and the COM prerequisites.

## Shared scaffold (`src/`) — optional

These describe the single-project scaffold under `src/`. No shipped macro depends on
it yet; it is reference material for new macro code.

- [ARCHITECTURE.md](ARCHITECTURE.md) — scaffold module map and design rules.
- [CONVENTIONS.md](CONVENTIONS.md) — naming and the safe-run entry pattern.
- [ADDING_A_MACRO.md](ADDING_A_MACRO.md) — adding a registered feature to the scaffold.
- [TESTING.md](TESTING.md) — pure-logic tests (Tier 1) and manual characterization (Tier 2).
- [RELEASE_CHECKLIST.md](RELEASE_CHECKLIST.md) — release gate for the scaffold.
