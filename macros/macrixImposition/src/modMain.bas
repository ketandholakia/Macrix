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

Public Sub ShowImpositionForm()
    InitDefaultSettings
    ' frmImpositionSettings is a real design-time form in this project (built by
    ' scripts\Build-Form.ps1), so it is shown directly: no runtime form generation
    ' and no VBE access required. Its "Run" button calls modImposition.RunImposition.
    frmImpositionSettings.Show
End Sub
