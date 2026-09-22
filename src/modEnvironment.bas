Attribute VB_Name = "modEnvironment"
Option Explicit
''
' modEnvironment: detect the runtime environment (§Environment Detection) and,
' in the now-thin layer, CorelDRAW version detection (see modVersionDetect for the
' version parsing logic; this module is only for non-Corel aspects).
''
' Depends on: modPathManager.
' CorelDRAW dependency: ENV_GetCorelInstallDir hits the registry app key (best-effort);
' all pure getters are testable.

Public Function ENV_IsWindows() As Boolean
    ENV_IsWindows = True   ' VBA/Windows is assumed by design; kept for completeness.
End Function

Public Function ENV_OsVersion() As String
    ENV_OsVersion = Environ$("OS")
End Function

Public Function ENV_Architecture() As String
    ' Environment prop may be absent in 32-bit; treat empty as x86 for safety.
    If Environ$("PROCESSOR_ARCHITEW6432") = "AMD64" Then
        ENV_Architecture = "AMD64"
    ElseIf Environ$("PROCESSOR_ARCHITECTURE") = "AMD64" Then
        ENV_Architecture = "AMD64"
    Else
        ENV_Architecture = "x86"
    End If
End Function

' Registry-based; hits the Windows registry via WScript. Corel detection should be
' treated as best-effort; users may install into non-default locations.
Public Function ENV_GetCorelInstallDir() As String
    Dim sh As Object
    Set sh = CreateObject("WScript.Shell")
    On Error Resume Next
    ENV_GetCorelInstallDir = sh.RegRead("HKLM\SOFTWARE\Corel\CorelDraw\CurrentVersion\Programs")
    On Error GoTo 0
    Set sh = Nothing
End Function

' free-ish: is a sane user data dir usable for logs/settings?
Public Function ENV_UserDataUsable() As Boolean
    Dim base As String
    base = modPathManager.PF_GetUserDataDir()
    ENV_UserDataUsable = (Len(base) > 0)
End Function