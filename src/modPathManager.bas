Attribute VB_Name = "modPathManager"
Option Explicit
''
' modPathManager: portable path/IO helpers (§Path Management).
' The only place allowed to know where things live on disk. Pure-logic helpers
' are unit-testable; the one caller that neither hard-codes nor fabricates is
' handled through Environ (Windows user-profile based).
''
' Depends on: nothing external.
' CorelDRAW dependency: none.

Public Const PF_AppFolder As String = "Vittix"

Public Function PF_GetUserDataDir() As String
    Dim base As String
    base = Environ$("APPDATA")
    If base = "" Then base = Environ$("USERPROFILE")   ' fallback if APPDATA unset
    PF_GetUserDataDir = PF_Join(base, PF_AppFolder)
End Function

Public Function PF_Join(ByVal a As String, ByVal b As String) As String
    a = Replace(a, "/", "\")
    While Right$(a, 1) = "\" And Len(a) > 1: a = Left$(a, Len(a) - 1): Wend
    PF_Join = a & "\" & b
End Function

Public Sub PF_EnsureDir(ByVal dirPath As String)
    If Dir$(dirPath, vbDirectory) = "" Then MkDir dirPath
End Sub

' Sanitize an arbitrary string into a Windows-safe file name. Pure (testable).
Public Function PF_SanitizeFilename(ByVal name As String) As String
    Dim i As Long, ch As String, res As String
    For i = 1 To Len(name)
        ch = Mid$(name, i, 1)
        If InStr("<>:""/\|?*", ch) > 0 Then
            res = res & "_"
        Else
            res = res & ch
        End If
    Next i
    res = Trim$(res)
    If res = "" Then res = "untitled"
    PF_SanitizeFilename = Left$(res, 200)
End Function