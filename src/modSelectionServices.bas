Attribute VB_Name = "modSelectionServices"
Option Explicit
''
' modSelectionServices — Selection handling. Late-bound (§15 Selection Safety).
'
' Selection safety rules (§15, §37):
'   - Always re-acquire the active selection; never cache a Selection/ShapeRange
'     across an operation that invalidates it (undo/delete/ungroup).
'   - Release references every iteration.
' Depends on: modComLifecycle. CorelDRAW: yes (characterization).

Public Function SV_Sel_Active(ByVal app As Object) As Object
    Set SV_Sel_Active = modAppServices.SV_App_ActiveSelection(app)
End Function

' Count of shapes in the active selection (re-acquired each call).
Public Function SV_Sel_Count(ByVal sel As Object) As Long
    On Error Resume Next
    SV_Sel_Count = CLng(CallByName(sel, "Count", VbGet))
    On Error GoTo 0
End Function

Public Function SV_Sel_IsEmpty(ByVal sel As Object) As Boolean
    SV_Sel_IsEmpty = (SV_Sel_Count(sel) = 0)
End Function