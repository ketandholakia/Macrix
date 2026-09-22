Attribute VB_Name = "modExportImportServices"
Option Explicit
''
' modExportImportServices — export/import path helpers.
' Path construction is pure (no CorelDRAW object model) so it is unit-testable;
' the actual file I/O for any given export lives with the feature that owns it
' (see feat_ExportText), consistent with modPathManager owning "where things
' live on disk" and features owning "what gets written there".
''
' Depends on: modPathManager.
' CorelDRAW dependency: none.

' Build a sanitized, timestamped export path under `dir` so re-running an
' export in the same session never silently overwrites the previous run.
' Ensures `dir` exists. Pure apart from that directory-creation side effect.
Public Function SVE_BuildExportPath(ByVal dir As String, ByVal baseName As String, ByVal ext As String) As String
    modPathManager.PF_EnsureDir dir
    Dim stamp As String
    stamp = Format$(Now, "yyyymmdd_hhnnss")
    Dim safeName As String
    safeName = modPathManager.PF_SanitizeFilename(baseName & "_" & stamp)
    If Left$(ext, 1) <> "." Then ext = "." & ext
    SVE_BuildExportPath = modPathManager.PF_Join(dir, safeName & ext)
End Function
