Attribute VB_Name = "modErrorHandler"
Option Explicit
''
' modErrorHandler: uniform error handling + session failure circuit breaker (§10, §46).
'
' §46 circuit breaker: if a feature errors repeatedly within a session
' (CFG_CircuitBreakerThreshold failures), it is auto-disabled for the remainder
' of the session, logged clearly, and surfaced to the user pointing at the
' diagnostics — never permanently.
'
' The single-exit restore pattern (Finally-equivalent) is: the caller sets up
' `On Error GoTo Cleanup`, does work, and both the success path and the error path
' `GoTo Cleanup` where application state (Optimization/EventsEnabled/command groups)
' and object references are restored before exit. This module supplies the reporting
' and circuit-breaking pieces of that pattern; the restore itself stays in the
' feature/caller per §39 (pair suppression with restoration on every path).
''
' Depends on: modConfig, modLogger.
' Pure: EH_FormatMessage / EH_IsDisabled logic are unit-testable.

Private m_failures As New Collection   ' featureId -> Long counter
Private m_disabled As New Collection   ' featureId (one-shot) set

Public Function EH_FormatMessage(ByVal errNo As Long, ByVal src As String, ByVal desc As String) As String
    EH_FormatMessage = "[#" & CStr(errNo) & " in " & src & "] " & desc
End Function

Public Sub EH_ReportError(ByVal errNo As Long, ByVal src As String, ByVal desc As String)
    Dim msg As String
    msg = EH_FormatMessage(errNo, src, desc)
    modLogger.LOG_Error msg
    Debug.Print msg
End Sub

' Record a failure for a feature; returns True if it crosses the threshold and is
' now disabled for the session.
Public Function EH_RecordFailure(ByVal featureId As String) As Boolean
    Dim cnt As Long
    EH_RecordFailure = False
    cnt = p_getCount(featureId) + 1
    p_setCount featureId, cnt
    If cnt >= modConfig.CFG_CircuitBreakerThreshold Then
        If Not p_isDisabled(featureId) Then
            p_disable featureId
            modLogger.LOG_Warn "circuit breaker: feature '" & featureId & _
                "' disabled for this session after " & CStr(cnt) & " failures"
            EH_RecordFailure = True
        End If
    End If
End Function

' Call on every success so a feature that recovers does not accumulate failures.
Public Sub EH_RecordSuccess(ByVal featureId As String)
    p_setCount featureId, 0
End Sub

Public Function EH_IsDisabled(ByVal featureId As String) As Boolean
    EH_IsDisabled = p_isDisabled(featureId)
End Function

Public Sub EH_ResetAllSessions()
    On Error Resume Next
    Dim i As Long
    For i = m_failures.Count To 1 Step -1
        m_failures.Remove i
    Next i
    For i = m_disabled.Count To 1 Step -1
        m_disabled.Remove i
    Next i
    On Error GoTo 0
End Sub

' ---- private helpers ----

Private Function p_getCount(ByVal featureId As String) As Long
    On Error Resume Next
    p_getCount = CLng(m_failures(featureId))
    On Error GoTo 0
End Function

Private Sub p_setCount(ByVal featureId As String, ByVal cnt As Long)
    On Error Resume Next
    m_failures.Remove Key:=featureId
    On Error GoTo 0
    m_failures.Add cnt, Key:=featureId
End Sub

Private Function p_isDisabled(ByVal featureId As String) As Boolean
    On Error Resume Next
    Dim dummy As Variant
    dummy = m_disabled.Item(featureId)   ' throws if absent
    p_isDisabled = (Err.Number = 0)
    On Error GoTo 0
End Function

Private Sub p_disable(ByVal featureId As String)
    On Error Resume Next
    m_disabled.Add featureId, Key:=featureId
    On Error GoTo 0
End Sub