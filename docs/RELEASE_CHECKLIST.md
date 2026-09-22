# Release checklist

Gate for shipping a version. Each item maps to a § or a build script. **Nothing is considered done
until the check is actually run.**

## Functional
- [ ] `build/validate.ps1` passes with 0 errors (≥ 0 warnings audited).
- [ ] `build/package.ps1` produced a versioned, signed-if-configured package in `build/_out`.
- [ ] Pure-logic suite green: `modTestRunner.RunAll` reports all pass.
- [ ] Corel-dependent characterization suite (`tests/manual/CHARACTERIZATION.md`) fully ticked on the
      target release install(s) — §41 (manual-only, no automation claim).

## Safety / install (§19, §45)
- [ ] Installer records crypto hash of previous installed files and can roll back.
- [ ] Uninstaller deletes only files owned by this app (ownership manifest).
- [ ] First-run trust/origin copy verified on a clean machine (§38) — user trusts the macro project.

## Docs & policy
- [ ] CHANGELOG entry for every user-visible change (§43).
- [ ] Feature-registry availability/stability flags reflect real status (§12).
- [ ] Ref registry documents every reference actually used (§42).

## Code review (§47)
- [ ] Commit messages follow `type(feature-id):` convention.
- [ ] Any change to a shared core service names every feature that consumes it.

## Version
- [ ] `modVersion.CurrentVersion` bumped per SemVer; matching package version string.
```