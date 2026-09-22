Attribute VB_Name = "modCorelCompatibility"
Option Explicit
''
' modCorelCompatibility — version-scoped feature switches and guards (§16, §42).
' Where a feature touches an API that changed between CorelDRAW versions, this
' module centralizes the capability check + fallback. All dispatch is guarded via
' modBinding so missing members are silent rather than a hard compile break.
'
' Depends on: modBinding, modVersionDetect. CorelDRAW: yes for live checks.

' Does the host expose `member`? Guarded via late binding.
Public Function CP_HasMember(ByVal app As Object, ByVal member As String) As Boolean
    CP_HasMember = modBinding.BND_HasMember(app, member)
End Function

' Determine whether to use the modern vs legacy export path (example switch).
Public Function CP_UseModernFilter(ByVal hostVersionScore As Long) As Boolean
    CP_UseModernFilter = (hostVersionScore >= 2020)
End Function