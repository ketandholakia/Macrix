# Characterization checklist

Targets: CorelDRAW X8, 2019, 2020, 2021 (adjust per §48 answer).
Status legend: [ ] pending · [x] passed · [!] failed → issue link.

## modAppServices
- [ ] `SV_App_ActiveDocument(Application)` returns a doc when one is open; Nothing when none.
- [ ] `SV_App_SetOptimization(Application, False)` then `True` restores redraw behavior.
- [ ] `SV_App_SetEventsEnabled(Application, False)` suppresses events and `True` re-enables.
- [ ] `BeginCommandGroup`/`EndCommandGroup` pair groups an undo step.

## modDocumentServices
- [ ] `SV_Doc_IsDirty` reflects unsaved edits on a test doc.
- [ ] `SV_Doc_Save(doc)` writes the file (no prompt).

## modSelectionServices
- [ ] `SV_Sel_Count(ActiveSelection)` equals the visually selected shape count.
- [ ] Selection count survives after `SV_App_SetOptimization False` toggle.

## modShapeServices / modLayerServices / modPageServices
- [ ] `SVS_SetSize` changes selected shape dimensions on the active page.
- [ ] `SV_Layer_Active(ActiveDocument)` `.Name` reads back the active layer.
- [ ] `SV_Pg_ActivePage(ActiveDocument).SizeWidth` matches the page setup.

Note: `modTextServices`/`modColorServices` don't exist yet (see ARCHITECTURE.md) — text access is
currently inlined in `feat_ExportText` via `CallByName`; characterize it under that feature below
rather than as a separate service module until one is extracted.

## export-import
- [ ] `SVE_BuildExportPath` (`modExportImportServices`) produces a sanitized, timestamped path and
      creates the user data dir if missing.

## feat_RoundedCorners
- [ ] On a shape selection, corners round (CornerRadius member present on this build).
- [ ] On an empty selection, shows a benign message and does not error.
- [ ] After 3 deliberate failures, the feature shows the session-disabled message (§46).

## feat_ExportText
- [ ] On a document with artistic/paragraph text shapes across multiple pages, the exported .txt
      contains each shape's text (one per line), via `Shape.Text.Story`.
- [ ] On a document with no text shapes, exports an empty (but valid, non-erroring) file.
- [ ] A genuine file-IO failure (e.g. read-only target dir) reports an error and counts toward the
      circuit breaker rather than silently claiming success — this path was previously untested and
      is a deliberate deviation from the "swallow via Resume Next" pattern used elsewhere; verify it
      actually surfaces to the user.

## Trust / install
- [ ] On a clean machine, first run shows the unsigned-macro trust message (§38).
- [ ] Reinstalling over an existing version preserves the prior files as a backup (§45).