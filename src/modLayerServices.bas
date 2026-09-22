Attribute VB_Name = "modLayerServices"
Option Explicit
''
' modLayerServices — Layer handling (active layer, create/rename, visibility).
' Late-bound. Depends on: modComLifecycle. CorelDRAW: yes (characterization).
' References are re-acquired per call; never cached across invalidation (§37).

Public Function SV_Layer_Active(ByVal doc As Object) As Object
    On Error Resume Next
    Set SV_Layer_Active = CallByName(doc, "ActiveLayer", VbGet)
    On Error GoTo 0
End Function

Public Function SV_Layer_Name(ByVal layer As Object) As String
    On Error Resume Next
    SV_Layer_Name = CStr(CallByName(layer, "Name", VbGet))
    On Error GoTo 0
End Function

Public Sub SV_Layer_SetVisible(ByVal layer As Object, ByVal visible As Boolean)
    On Error Resume Next
    CallByName layer, "Visible", VbLet, visible
    On Error GoTo 0
End Sub