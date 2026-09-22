Attribute VB_Name = "modComLifecycle"
Option Explicit
''
' modComLifecycle: COM object lifecycle discipline (§37).
'
' VBA does not GC COM references deterministically. Long-running or loop-heavy
' macros against CorelDRAW's object model can leak or leave the app degraded if
' references are not released. These helpers encode the rules:
'   - release ShapeRange / Selection / Layer / Page references inside loops,
'   - never cache a shape across an invalidating operation (undo/delete/ungroup),
'   - hoist object acquisition out of tight loops where the source is stable,
'   - pair any state change with a restore in a cleanup handler.
''
' Depends on: modLogger (optional reporting).
' CorelDRAW dependency: none (these act on generic Object types).

' Zero any Object var you no longer need. Call at the end of each loop iteration
' and in every cleanup label.
Public Sub CM_Release(ByRef target As Object)
    On Error Resume Next
    If Not target Is Nothing Then
        Set target = Nothing
    End If
End Sub

' Release a whole object array (Variant() of Objects) — handy for range cleanup.
Public Sub CM_ReleaseArray(ByRef items As Variant)
    Dim i As Long
    If IsArray(items) Then
        For i = LBound(items) To UBound(items)
            If IsObject(items(i)) Then Set items(i) = Nothing
        Next i
    End If
End Sub

' Convenience: returns an interface the caller can call functions on
' without storing a reference. For pointer-heavy Corel loops prefer re-acquiring
' an object each access instead of caching.
Public Function CM_Wrap(ByVal obj As Object) As Object
    Set CM_Wrap = obj
End Function

' Paired command-group / optimization restore pattern can be delegated here.
Public Function CM_CaptureOptimization(ByVal app As Object, ByVal propName As String) As Boolean
    On Error Resume Next
    CM_CaptureOptimization = CBool(CallByName(app, propName, VbGet))
    On Error GoTo 0
End Function

Public Sub CM_RestoreOptimization(ByVal app As Object, ByVal state As Boolean)
    On Error Resume Next
    CallByName app, "Optimization", VbLet, state
    On Error GoTo 0
End Sub

' Hoist/loop helper: create a fresh instance of a late-bound COM server.
' Pure-ish; requires runtime COM.
Public Function CM_CreateInstance(ByVal progId As String) As Object
    On Error Resume Next
    Set CM_CreateInstance = CreateObject(progId)
    On Error GoTo 0
End Function