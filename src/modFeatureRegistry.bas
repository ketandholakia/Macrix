Attribute VB_Name = "modFeatureRegistry"
Option Explicit
''
' modFeatureRegistry â€” feature/module registration & gating (Â§12, Â§26, Â§43).
'
' Central register of every feature. Features are added in FR_RegisterFeatures()
' (one line each) and discovered/listed here rather than through hard-coded menus.
' The registry also drives availability gating, per-feature versioning (Â§43),
' deprecation status, and the Â§46 circuit-breaker pre-check. Actual invocation is
' the feature module's own public entry point, which itself wraps the safe-run
' pattern; FR_EntryPoint returns which "Module.entry" the caller should invoke.
''
' Depends on: modErrorHandler, modLogger, modVersion.
' CorelDRAW dependency: none for bookkeeping.

Public Enum FB_Avail
    fb_Available = 0
    fb_Disabled = 1
    fb_Deprecated = 2
End Enum

Public Enum FB_Stability
    fb_Stable = 0
    fb_Beta = 1
    fb_Experimental = 2
End Enum

Private Type FeatureRec
    Name As String
    Display As String
    Module As String
    Entry As String
    Avail As FB_Avail
    Stability As FB_Stability
    Version As String
    MinPlatform As String
End Type

Private m_feats() As FeatureRec
Private m_n As Long

' --- Registration ---

Public Sub FR_Register(ByVal name As String, ByVal display As String, _
                       ByVal moduleName As String, ByVal entry As String, _
                       ByVal avail As FB_Avail, ByVal stability As FB_Stability, _
                       ByVal version As String, Optional ByVal minPlatform As String = "0.0.0")
    If name = "" Then Exit Sub
    m_n = m_n + 1
    ReDim Preserve m_feats(1 To m_n)
    With m_feats(m_n)
        .Name = name
        .Display = display
        .Module = moduleName
        .Entry = entry
        .Avail = avail
        .Stability = stability
        .Version = version
        .MinPlatform = minPlatform
    End With
End Sub

' Register every shipped feature (single curated place per ADDING_A_MACRO).
' Idempotent: safe to call more than once in a session (modBootstrap.Main may
' run again e.g. during manual testing) â€” resets the table first so repeated
' calls never duplicate entries.
Public Sub FR_RegisterFeatures()
    m_n = 0
    Erase m_feats
    FR_Register "rounded-corners", "Rounded Corners", "feat_RoundedCorners", "Main", _
        fb_Available, fb_Stable, "0.1.0", "0.1.0"
    FR_Register "export-text", "Export Selection to Text", "feat_ExportText", "Main", _
        fb_Available, fb_Beta, "0.1.0", "0.1.0"
End Sub

' ---- Query ----

Public Function FR_IsAvailable(ByVal name As String) As Boolean
    Dim i As Long
    i = p_index(name)
    If i = 0 Then FR_IsAvailable = False: Exit Function
    If m_feats(i).Avail <> fb_Available Then Exit Function
    ' session circuit breaker
    If modErrorHandler.EH_IsDisabled(name) Then Exit Function
    FR_IsAvailable = True
End Function

Public Function FR_EntryPoint(ByVal name As String) As String
    Dim i As Long
    i = p_index(name)
    If i = 0 Then FR_EntryPoint = "": Exit Function
    FR_EntryPoint = m_feats(i).Module & "." & m_feats(i).Entry
End Function

Public Function FR_Count() As Long
    FR_Count = m_n
End Function

Public Function FR_FeatureName(ByVal idx As Long) As String
    If idx < 1 Or idx > m_n Then Exit Function
    FR_FeatureName = m_feats(idx).Name
End Function

Public Function FR_DisplayName(ByVal idx As Long) As String
    If idx < 1 Or idx > m_n Then Exit Function
    FR_DisplayName = m_feats(idx).Display
End Function

' Report string for About/logs.
Public Function FR_Describe() As String
    Dim i As Long, acc As String
    For i = 1 To m_n
        acc = acc & m_feats(i).Name & " [" & m_feats(i).Stability & "/" & m_feats(i).Version & "]; "
    Next i
    FR_Describe = acc
End Function

Private Function p_index(ByVal name As String) As Long
    Dim i As Long
    For i = 1 To m_n
        If StrComp(m_feats(i).Name, name, vbTextCompare) = 0 Then
            p_index = i
            Exit For
        End If
    Next i
End Function
