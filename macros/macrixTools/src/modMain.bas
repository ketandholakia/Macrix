Attribute VB_Name = "modMain"
Option Explicit
Rem modMain.bas - Entry points and high-level orchestration for Macrix Dimension Tools

Public Sub MacrixDimension()
    On Error GoTo ErrHandler
    LoadSettings

    If ActiveDocument Is Nothing Then
        MsgBox "Open a document before running Macrix Dimension Tools.", vbExclamation, "Macrix"
        Exit Sub
    End If

    ShowDimensionForm
    Exit Sub
ErrHandler:
    LogError "MacrixTools_Main", Err.Number, Err.Description
End Sub

' Only MacrixDimension is meant to appear in CorelDRAW's macro list. The OK
' button on the dialog calls CreateDimensionsForSelection (modDimension) directly,
' so this stays Private and is therefore not listed.
Private Sub MacrixTools_Create(Optional hideMe As Boolean = True)
    On Error GoTo ErrHandler
    If ActiveDocument Is Nothing Then
        MsgBox "Open a document before creating dimensions.", vbExclamation, "Macrix"
        Exit Sub
    End If

    CreateDimensionsForSelection gSettings
    SaveSettings
    Exit Sub
ErrHandler:
    LogError "MacrixTools_Create", Err.Number, Err.Description
End Sub
