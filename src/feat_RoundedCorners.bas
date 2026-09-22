Attribute VB_Name = "feat_RoundedCorners"
Option Explicit
''
' feat_RoundedCorners — sample feature: round the corners of selected shapes.
'
' Demonstrates the mandated entry pattern:
'   • thin public entry point ("Main") → registered in modFeatureRegistry,
'   • single-exit cleanup label restoring Optimization/EventsEnabled on BOTH paths,
'   • §46 circuit-breaker pre-check + failure recording,
'   • references released on every exit (§37).
'
' The shape-mutation line is Corel-dependent and version-sensitive; mark for manual
' characterization (§40) and validate on the target install before promoting.
'
' Depends on: service modules, modUiUtilities, modComLifecycle.
' CorelDRAW: yes (characterization).

Private Const F_BreakerId As String = "rounded-corners"

Public Sub Main()
    Dim app As Object, doc As Object, sel As Object
    On Error GoTo Cleanup
    If modErrorHandler.EH_IsDisabled(F_BreakerId) Then
        Err.Raise 1001, "feat_RoundedCorners.Main", "feature session-disabled"
    End If
    Set app = modAppServices.SV_App_Host()
    If app Is Nothing Then Err.Raise 1002, "feat_RoundedCorners.Main", "host unavailable"
    Set doc = modAppServices.SV_App_ActiveDocument(app)
    Set sel = modAppServices.SV_App_ActiveSelection(app)
    If modSelectionServices.SV_Sel_IsEmpty(sel) Then
        modUiUtilities.UU_MsgInfo "Please select one or more shapes first."
        GoTo Cleanup
    End If

    modAppServices.SV_App_SetOptimization app, False
    modAppServices.SV_App_SetEventsEnabled app, False
    p_SelectionRound sel
    modErrorHandler.EH_RecordSuccess F_BreakerId
    modUiUtilities.UU_MsgInfo "Rounded corners applied."

Cleanup:
    modAppServices.SV_App_SetOptimization app, True
    modAppServices.SV_App_SetEventsEnabled app, True
    modComLifecycle.CM_Release app
    modComLifecycle.CM_Release doc
    modComLifecycle.CM_Release sel
    If Err.Number <> 0 Then
        modErrorHandler.EH_ReportError Err.Number, "feat_RoundedCorners.Main", Err.Description
        If modErrorHandler.EH_RecordFailure(F_BreakerId) Then
            modUiUtilities.UU_MsgWarn "This feature is disabled for this session; see the log."
        End If
    End If
End Sub

' Apply a modest corner radius to each selected shape. No shape reference is
' retained beyond a single iteration (§37). The shape-enumeration and radius
' members are characterization pending — tune to the installed type library.
Private Sub p_SelectionRound(ByVal sel As Object)
    On Error Resume Next
    Dim i As Long, cnt As Long
    Dim range As Object, sh As Object
    cnt = modSelectionServices.SV_Sel_Count(sel)
    Set range = CallByName(sel, "Shapes", VbGet)   ' confirm member
    For i = 1 To cnt
        Set sh = CallByName(range, "Item", VbMethod, i)
        CallByName sh, "CornerRadius", VbLet, 0.125
        modComLifecycle.CM_Release sh
    Next i
    modComLifecycle.CM_Release range
    On Error GoTo 0
End Sub