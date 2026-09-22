Attribute VB_Name = "@@MODULE@@"
Option Explicit
''
' @@MODULE@@ — @@DISPLAY@@.
'
' Entry point: Main (invoked from CorelDRAW; registered in modFeatureRegistry as
' "@@FEATURE_ID@@").
'
' Scaffolded by build\new-feature.ps1. Follows the mandated safe-run pattern
' (docs/CONVENTIONS.md):
'   • thin public entry point ("Main"),
'   • single-exit cleanup label restoring Optimization/EventsEnabled on BOTH paths,
'   • circuit-breaker pre-check + failure recording (modErrorHandler),
'   • object references released on every exit (modComLifecycle).
'
' Depends on: modAppServices, modSelectionServices, modComLifecycle,
' modUiUtilities, modErrorHandler.
' CorelDRAW: yes — mark for manual characterization
' (tests/manual/CHARACTERIZATION.md) until verified on the target install.

Private Const F_BreakerId As String = "@@FEATURE_ID@@"

Public Sub Main()
    Dim app As Object, doc As Object, sel As Object
    On Error GoTo Cleanup
    If modErrorHandler.EH_IsDisabled(F_BreakerId) Then
        Err.Raise @@ERR_BASE@@, "@@MODULE@@.Main", "feature session-disabled"
    End If
    Set app = modAppServices.SV_App_Host()
    If app Is Nothing Then Err.Raise @@ERR_BASE@@ + 1, "@@MODULE@@.Main", "host unavailable"
    Set doc = modAppServices.SV_App_ActiveDocument(app)
    Set sel = modAppServices.SV_App_ActiveSelection(app)
    If modSelectionServices.SV_Sel_IsEmpty(sel) Then
        modUiUtilities.UU_MsgInfo "Please select one or more shapes first."
        GoTo Cleanup
    End If

    modAppServices.SV_App_SetOptimization app, False
    modAppServices.SV_App_SetEventsEnabled app, False
    p_DoWork sel
    modErrorHandler.EH_RecordSuccess F_BreakerId
    modUiUtilities.UU_MsgInfo "@@DISPLAY@@ complete."

Cleanup:
    modAppServices.SV_App_SetOptimization app, True
    modAppServices.SV_App_SetEventsEnabled app, True
    modComLifecycle.CM_Release app
    modComLifecycle.CM_Release doc
    modComLifecycle.CM_Release sel
    If Err.Number <> 0 Then
        modErrorHandler.EH_ReportError Err.Number, "@@MODULE@@.Main", Err.Description
        If modErrorHandler.EH_RecordFailure(F_BreakerId) Then
            modUiUtilities.UU_MsgWarn "This feature is disabled for this session; see the log."
        End If
    End If
End Sub

' The real work. Corel-touching; release each reference inside the loop (§37).
' Replace the body below with the actual operation. Prefer calling a mod*Services
' helper over reaching into the object model directly (see docs/ADDING_A_MACRO.md).
Private Sub p_DoWork(ByVal sel As Object)
    On Error Resume Next
    Dim i As Long, cnt As Long
    Dim range As Object, sh As Object
    cnt = modSelectionServices.SV_Sel_Count(sel)
    Set range = CallByName(sel, "Shapes", VbGet)
    For i = 1 To cnt
        Set sh = CallByName(range, "Item", VbMethod, i)
        ' TODO: perform the operation on `sh` here.
        modComLifecycle.CM_Release sh
    Next i
    modComLifecycle.CM_Release range
    On Error GoTo 0
End Sub
