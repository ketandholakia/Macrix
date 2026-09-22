Attribute VB_Name = "modShapeServices"
Option Explicit
''
' modShapeServices — Shape / ShapeRange operations. Late-bound.
' Safety (§37): never cache a Shape/ShapeRange across an invalidating operation
' (undo / delete / ungroup) — re-acquire. Release refs each iteration.
' Depends on: modComLifecycle, modLogger. CorelDRAW: yes (characterization).

Public Function SVS_ShapeCount(ByVal shapeRange As Object) As Long
    On Error Resume Next
    SVS_ShapeCount = CLng(CallByName(shapeRange, "Count", VbGet))
    On Error GoTo 0
End Function

Public Sub SVS_Delete(ByVal shapeRange As Object)
    On Error Resume Next
    CallByName shapeRange, "Delete", VbMethod
    On Error GoTo 0
End Sub

Public Sub SVS_SetSize(ByVal shapeRange As Object, ByVal width As Double, ByVal height As Double)
    On Error Resume Next
    CallByName shapeRange, "SizeWidth", VbLet, width
    CallByName shapeRange, "SizeHeight", VbLet, height
    On Error GoTo 0
End Sub