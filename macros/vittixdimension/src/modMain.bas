Attribute VB_Name = "modMain"
Option Explicit
Rem modMain.bas - Entry points and high-level orchestration for Vittix Dimension Tools

Public Sub VittixDimension()
    On Error GoTo ErrHandler
    LoadSettings

    If ActiveDocument Is Nothing Then
        MsgBox "Open a document before running Vittix Dimension Tools.", vbExclamation, "Vittix"
        Exit Sub
    End If

    ShowDimensionForm
    Exit Sub
ErrHandler:
    LogError "VittixDimensionTools_Main", Err.Number, Err.Description
End Sub

' Only VittixDimension is meant to appear in CorelDRAW's macro list. The OK
' button on the dialog calls CreateDimensionsForSelection (modDimension) directly,
' so this stays Private and is therefore not listed.
Private Sub VittixDimensionTools_Create(Optional hideMe As Boolean = True)
    On Error GoTo ErrHandler
    If ActiveDocument Is Nothing Then
        MsgBox "Open a document before creating dimensions.", vbExclamation, "Vittix"
        Exit Sub
    End If

    CreateDimensionsForSelection gSettings
    SaveSettings
    Exit Sub
ErrHandler:
    LogError "VittixDimensionTools_Create", Err.Number, Err.Description
End Sub
