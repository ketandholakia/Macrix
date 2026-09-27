Attribute VB_Name = "modMain"
Option Explicit
''
' modMain -- the single public entry point for this macro.
'
' Only this module is listed in CorelDRAW's macro list. Everything else lives in
' Option Private Module modules, which stay callable everywhere inside the project
' (the form included) but do not clutter the macro list.
''
' Depends on: modSettings (InitDefaultSettings), modImposition, frmImpositionSettings.
'             modTileFillSettings (InitDefaultTileSettings), modTileFill,
'             frmTileFillSettings.

Public Sub ShowImpositionForm()
    ' Load once per session: InitDefaultSettings resets to built-in defaults
    ' and overlays last-used registry values. Guarded so reopening the form
    ' never wipes the in-memory last-used values.
    If Not g_SettingsInitialized Then InitDefaultSettings
    ' frmImpositionSettings is a real design-time form in this project (built by
    ' scripts\Build-Form.ps1), so it is shown directly: no runtime form generation
    ' and no VBE access required. Its "Run" button calls modImposition.RunImposition.
    frmImpositionSettings.Show
End Sub

Public Sub ShowTileFillForm()
    ' Same pattern as ShowImpositionForm: InitDefaultTileSettings resets to
    ' built-in defaults and overlays last-used registry values, guarded so
    ' reopening the form never wipes in-memory last-used values.
    If Not g_TileSettingsInitialized Then InitDefaultTileSettings
    ' frmTileFillSettings is built the same way frmImpositionSettings is (see
    ' mdlTileFillFormBuilder.BuildTileFillFormWithCode) -- run that once from
    ' the Immediate Window (or via Build-Form.ps1, if that script is extended
    ' to target it) before this will find the form. Its "Tile Fill Selection"
    ' button calls modTileFill.RunTileFill.
    frmTileFillSettings.Show
End Sub
