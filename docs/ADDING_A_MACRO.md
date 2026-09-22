# How to add a new macro (feature)

> **Scope:** this describes adding a **feature to the shared scaffold project**
> (`src/`), not to the shipped macros under `macros/`. To work on a shipped macro,
> see [`MACROS.md`](MACROS.md).

The fast path. Full detail in §26 of the master prompt.

1. **Decide the feature id** — a kebab-case slug, e.g. `rounded-corners`. Keep this stable; it is your
   identifier in the registry, CHANGELOG, and commit messages.

2. **Add the feature module** — `src/feat_<YourName>.bas`:
   - One public entry sub, conventionally named `Main`, callable from CorelDRAW.
   - It must follow the **safe-run pattern** in `docs/CONVENTIONS.md` (`On Error GoTo Cleanup` + a
     single cleanup label that restores app state and reports/counts any failure via
     `modErrorHandler` on every exit path). Copy `feat_RoundedCorners.Main` as a starting point.
   - Do the real work in a private sub/function prefixed `p_` — **not** a leading underscore; VBA
     identifiers must start with a letter (see Naming in `docs/CONVENTIONS.md`). Keep the entry point
     itself thin.

3. **Register it** — in `modFeatureRegistry.FR_RegisterFeatures()` add one line:
   ```vb
   FR_Register "rounded-corners", "Rounded Corners", "feat_RoundedCorners", "Main", _
       fb_Available, fb_Stable, "0.1.0", "1.0.0"
   ```
   The registry handles gating + availability. `FR_RegisterFeatures` resets its table before
   re-registering everything, so it's safe to call more than once per session.

4. **Keep the UI wiring out** — do not hard-code a menu. If the feature needs a menu/dock, add a
   registration record with a named UI cause and let `modUiUtilities` bind it.

5. **Wire services** — if it transforms selection, call `modSelectionServices`/`modShapeServices`
   instead of reaching into the object model directly. This is what makes reuse possible. If no
   service module covers what you need yet (e.g. text or color access), it's fine to call
   `CallByName`/`modBinding` directly from the feature for now — extract a new `modXServices` module
   once a second feature needs the same access, rather than speculatively before there's a second
   caller.

6. **Test**
   - Pure logic → add assertions to `modTestRunner.RunAll` (see the existing `T_Check` calls).
   - Corel-dependent → add a row to `tests/manual/CHARACTERIZATION.md` and run the checklist on a
     real install.

7. **Release doc** — bump `modVersion` PATCH (or MINOR for a new feature), add a CHANGELOG entry, and
   commit with `feat(rounded-corners): …` (§47).

## Anti-patterns to avoid
- ✗ Global mutable state accessed directly from a generator feature.
- ✗ Caching a `Shape`/`Selection` across an undo/delete/ungroup (§37).
- ✗ Changing `Optimization`/`EventsEnabled` without a Finally-equivalent restore (§39).
- ✗ Hard-coded file paths. Use `modPathManager`.
- ✗ A private helper named with a leading underscore (`_helper`) — won't compile. Use `p_helper`.
- ✗ Swallowing a failure with `On Error Resume Next` around anything whose success the feature
  actually depends on (vs. a late-bound member-existence probe) — it defeats the circuit breaker and
  reports false successes to the user. See "Error handling inside service helpers" in
  `docs/CONVENTIONS.md`.
