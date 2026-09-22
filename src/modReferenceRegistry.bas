Attribute VB_Name = "modReferenceRegistry"
Option Explicit
''
' modReferenceRegistry — reference integrity & binding policy registry (§42, §11).
'
' Every external reference the project actually uses (Tools → References, plus
' any CreateObject ProgID) must be listed here with a documented rationale. This is
' the single source of truth the build step (`build/validate.ps1`) cross-checks so
' we never ship something that silently depends on whatever happens to be checked
' on a developer box.
'
' Binding policy: LATE by default (§42). Early binding is only permitted for a
' reference listed below as "early" AND with a runtime fallback in modBinding.
'
' Depends on: nothing.
' CorelDRAW dependency: none (declarative).

Public Type RefRec
    ProgId As String
    Kind As String          ' "coreldraw" | "shell" | "other"
    Binding As String       ' "late" / "early"
    Guid As String
    Version As String
    Note As String
End Type

Private m_Initialized As Boolean
Private m_Refs() As RefRec

Public Sub REF_Initialize()
    If m_Initialized Then Exit Sub
    ReDim m_Refs(1 To 2)
    With m_Refs(1)
        .ProgId = "CorelDRAW.Application"
        .Kind = "coreldraw"
        .Binding = "late"
        .Guid = ""
        .Version = "host"
        .Note = "CorelDRAW host object model"
    End With
    With m_Refs(2)
        .ProgId = "WScript.Shell"
        .Kind = "shell"
        .Binding = "late"
        .Guid = "{72C24DD5-D70A-438B-8A42-98424B88AFB8}"
        .Version = "1.0"
        .Note = "registry reads (modEnvironment)"
    End With
    m_Initialized = True
End Sub

Public Function REF_Count() As Long
    If Not m_Initialized Then REF_Initialize
    REF_Count = UBound(m_Refs)
End Function

Public Function REF_ProgId(ByVal idx As Long) As String
    If Not m_Initialized Then REF_Initialize
    REF_ProgId = m_Refs(idx).ProgId
End Function

Public Function REF_Binding(ByVal idx As Long) As String
    If Not m_Initialized Then REF_Initialize
    REF_Binding = m_Refs(idx).Binding
End Function

' Predicate: is ProgID declared in this registry?
Public Function REF_IsDocumented(ByVal progId As String) As Boolean
    Dim i As Long
    If Not m_Initialized Then REF_Initialize
    For i = 1 To UBound(m_Refs)
        If StrComp(m_Refs(i).ProgId, progId, vbTextCompare) = 0 Then
            REF_IsDocumented = True
            Exit For
        End If
    Next i
End Function

' Semicolon-joined list for validation output / logs.
Public Function REF_AllNames(Optional ByVal separator As String = ";") As String
    Dim i As Long, acc As String
    If Not m_Initialized Then REF_Initialize
    For i = 1 To UBound(m_Refs)
        acc = acc & m_Refs(i).ProgId & separator
    Next i
    REF_AllNames = acc
End Function