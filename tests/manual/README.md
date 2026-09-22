# Manual Characterization Tests

Tier 2 of the testing strategy (see `docs/TESTING.md`). CorelDRAW-dependent behavior
cannot be fully automated on a CI runner (§41), so every Core-dependent module
listed in `CHARACTERIZATION.md` must be ticked on a **real** CorelDRAW install before
a release is certified.

## How to run
1. Open the matching document state described in each case.
2. Run the macro from CorelDRAW VBA.
3. Record PASS / FAIL and any notes.
4. Link the row to the build/version that was tested.

## Rules
- Do NOT mark a row PASS unless you ran it on the target install and observed the result.
- A FAIL must be captured (message + stack) and linked to a tracking issue before release.
- These tests are the ONLY evidence for "Corel-dependent modules are verified."