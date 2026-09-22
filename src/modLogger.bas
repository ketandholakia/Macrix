Attribute VB_Name = "modLogger"
Option Explicit
''
' modLogger: structured local-only diagnostics (§9, §43: no telemetry).
' Default: writes to a local log file under the user's app-data dir managed by
' modPathManager. Nothing is ever transmitted off-box (§44) unless a feature
' explicitly opts in and the user has consented — which this module forbids by
' design when CFG_TelemetryAllowed is False.
''
' Depends on: modConfig, modPathManager (no Core object model).

Private m_LogFileHandle As Integer

' Write a line at the configured level (numeric; quiet unless level >= threshold).
Public Sub LOG_Log(ByVal level As Long, ByVal message As String)
    If level < modConfig.CFG_LogLevel Then Exit Sub
    If m_LogFileHandle = 0 Then Call p_openFile
    Dim line As String
    line = Format$(Now, "yyyy-mm-dd hh:nn:ss") & " [" & LOG_LevelTag(level) & "] " & message
    If modConfig.CFG_LogToFile And m_LogFileHandle <> 0 Then
        Print #m_LogFileHandle, line
    End If
    Debug.Print line
End Sub

Public Sub LOG_Info(ByVal msg As String)
    LOG_Log modConfig.CFG_LevelInfo, msg
End Sub

Public Sub LOG_Warn(ByVal msg As String)
    LOG_Log modConfig.CFG_LevelWarn, msg
End Sub

Public Sub LOG_Error(ByVal msg As String)
    LOG_Log modConfig.CFG_LevelError, msg
End Sub

' Pure helper: level integer -> tag string (unit-testable).
Public Function LOG_LevelTag(ByVal level As Long) As String
    Select Case level
        Case 0: LOG_LevelTag = "DBG"
        Case 1: LOG_LevelTag = "INFO"
        Case 2: LOG_LevelTag = "WARN"
        Case 3: LOG_LevelTag = "ERR"
        Case Else: LOG_LevelTag = "?"
    End Select
End Function

Private Sub p_openFile()
    modPathManager.PF_EnsureDir modPathManager.PF_GetUserDataDir()
    m_LogFileHandle = FreeFile
    Open modPathManager.PF_Join( _
        modPathManager.PF_GetUserDataDir(), "macrix.log") For Append As #m_LogFileHandle
End Sub
