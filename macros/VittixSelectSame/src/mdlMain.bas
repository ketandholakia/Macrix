Attribute VB_Name = "mdlMain"
Option Explicit

' ===========================================================================
' mdlMain
'
' Run SelectSimilarObjects (this is the macro you assign to a shortcut key
' or toolbar button - see README.md).
' ===========================================================================

Public Sub SelectSimilarObjects()

    If Application.Documents.Count = 0 Then
        MsgBox "Please open a document first.", vbExclamation, "Select Similar Objects"
        Exit Sub
    End If

    Dim refShape As Shape
    Set refShape = GetReferenceShapeForMacro()

    If refShape Is Nothing Then
        MsgBox "Please select a single object first, then run Select Similar Objects.", _
            vbExclamation, "Select Similar Objects"
        Exit Sub
    End If

    Set frmSelectSimilar.ReferenceShape = refShape
    frmSelectSimilar.Show
End Sub

Public Function GetReferenceShapeForMacro() As Shape
    On Error Resume Next
    If Not ActiveDocument.ActiveShape Is Nothing Then
        Set GetReferenceShapeForMacro = ActiveDocument.ActiveShape
    ElseIf ActiveDocument.ActiveSelectionRange.Count >= 1 Then
        Set GetReferenceShapeForMacro = ActiveDocument.ActiveSelectionRange(1)
    End If
End Function


