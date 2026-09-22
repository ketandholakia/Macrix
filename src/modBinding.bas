Attribute VB_Name = "modBinding"
Option Explicit
''
' modBinding — reference/binding consistency (§42).
'
' Default policy: LATE binding (CreateObject + generic Object + CallByName) so a
' .bas imports cleanly into machines running different CorelDRAW type-library
' versions without a "missing reference" error. Where early binding is used during
' development for IntelliSense, the exact Reference (Name+GUID+Version) must be
' declared in modReferenceRegistry and a runtime fallback provided here.
''
' Depends on: modReferenceRegistry.
' CorelDRAW dependency: COM only.

' Create a late-bound instance by ProgID; empty if unavailable.
Public Function BND_CreateObject(ByVal progId As String) As Object
    On Error Resume Next
    Set BND_CreateObject = CreateObject(progId)
    On Error GoTo 0
End Function

' Does the object expose member `name` (either method or property)?
' Guarded by CallByName inside a resume so missing members are silent.
Public Function BND_HasMember(ByVal obj As Object, ByVal member As String) As Boolean
    On Error Resume Next
    Dim dummy As Variant
    dummy = CallByName(obj, member, VbMethod)
    BND_HasMember = (Err.Number = 0)
    On Error GoTo 0
End Function

' Return the ProgID for the CorelDRAW host per policy. This is a documented
' assumption — see ReferenceRegistry; adjust to the registered ProgID.
Public Function BND_CorelProgId() As String
    BND_CorelProgId = "CorelDRAW.Application"
End Function