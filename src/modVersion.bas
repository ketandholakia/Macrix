Attribute VB_Name = "modVersion"
Option Explicit
''
' modVersion — Semantic Versioning for the platform and features (§43).
' The platform as a whole is MAJOR.MINOR.PATCH; individual features may version
' independently within the registry (each FR_Register carries its own version).
'
' Depends on: nothing (pure).
' CorelDRAW dependency: none.
''

Public Const VER_KitName As String = "Macrix CorelDRAW Macros"

' SemVer: MAJOR.MINOR.PATCH (no prerelease/build-meta for the toolkit itself).
Public Function VER_CurrentVersion() As String
    VER_CurrentVersion = "0.1.0"
End Function

Public Function VER_Major() As Long
    VER_Major = CLng(p_Split(VER_CurrentVersion, ".")(0))
End Function

Public Function VER_Minor() As Long
    VER_Minor = CLng(p_Split(VER_CurrentVersion, ".")(1))
End Function

Public Function VER_Patch() As Long
    VER_Patch = CLng(p_Split(VER_CurrentVersion, ".")(2))
End Function

' Compare two versions "a.b.c"; returns -1,0,1. Pure (unit-testable).
Public Function VER_Compare(ByVal a As String, ByVal b As String) As Long
    Dim pa As Variant, pb As Variant
    pa = p_Split(a, ".")
    pb = p_Split(b, ".")
    If CLng(pa(0)) < CLng(pb(0)) Then VER_Compare = -1: Exit Function
    If CLng(pa(0)) > CLng(pb(0)) Then VER_Compare = 1: Exit Function
    If CLng(pa(1)) < CLng(pb(1)) Then VER_Compare = -1: Exit Function
    If CLng(pa(1)) > CLng(pb(1)) Then VER_Compare = 1: Exit Function
    If CLng(pa(2)) < CLng(pb(2)) Then VER_Compare = -1: Exit Function
    If CLng(pa(2)) > CLng(pb(2)) Then VER_Compare = 1: Exit Function
    VER_Compare = 0
End Function

' True iff a >= b (semantic).
Public Function VER_AtLeast(ByVal a As String, ByVal b As String) As Boolean
    VER_AtLeast = (VER_Compare(a, b) >= 0)
End Function

Private Function p_Split(ByVal text As String, ByVal delim As String) As Variant
    p_Split = Split(text, delim)
End Function