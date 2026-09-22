Attribute VB_Name = "modVersionDetect"
Option Explicit
''
' modVersionDetect — CorelDRAW version detection & parsing (§6, §16).
' Parsing logic is pure and unit-testable; reading the live host version is
' characterized manually (Corel-dependent).
'
' Depends on: modBinding (optionally). CorelDRAW: partial.

' Read a numeric version from the host Application.Version; fallback 0.
Public Function VD_HostVersion(ByVal app As Object) As Long
    On Error Resume Next
    Dim v As String
    v = CStr(CallByName(app, "Version", VbGet))
    On Error GoTo 0
    VD_HostVersion = VD_ParseCorelVersion(v)
End Function

' Parse a CorelDRAW version string like "2021.0.0", "18.0.0.448", or "X8" into
' a sortable Long. Splits on "." and weights each segment independently
' (major*1,000,000 + minor*1,000 + patch) rather than concatenating raw
' digits — concatenation broke ordering whenever segments had different digit
' counts (e.g. a trailing build number). Mirrors modVersion.VER_Compare's
' segment-by-segment semantics instead of a separate ad hoc scheme. Pure.
Public Function VD_ParseCorelVersion(ByVal raw As String) As Long
    Dim parts As Variant
    parts = Split(raw, ".")
    Dim major As Long, minor As Long, patch As Long
    major = p_ClampSegment(p_SegmentDigits(parts, 0))
    minor = p_ClampSegment(p_SegmentDigits(parts, 1))
    patch = p_ClampSegment(p_SegmentDigits(parts, 2))
    VD_ParseCorelVersion = (major * 1000000) + (minor * 1000) + patch
End Function

' Is the running CorelDRAW >= minVersionScore per policy?
Public Function VD_MetMinimum(ByVal app As Object, ByVal minScore As Long) As Boolean
    VD_MetMinimum = (VD_HostVersion(app) >= minScore)
End Function

' ---- private helpers ----

' Digits-only value of Split(raw, ".")(idx), or 0 if the segment is absent or
' has no digits (e.g. the "X8" legacy naming scheme has no dot segments at
' all beyond index 0, and that segment is not purely numeric).
Private Function p_SegmentDigits(ByRef parts As Variant, ByVal idx As Long) As Long
    If idx > UBound(parts) Then Exit Function
    Dim raw As String, i As Long, ch As String, digitsOnly As String
    raw = CStr(parts(idx))
    For i = 1 To Len(raw)
        ch = Mid$(raw, i, 1)
        If ch >= "0" And ch <= "9" Then digitsOnly = digitsOnly & ch
    Next i
    If digitsOnly = "" Then Exit Function
    p_SegmentDigits = CLng(digitsOnly)
End Function

' Keep any single segment from overflowing into the next segment's weight.
Private Function p_ClampSegment(ByVal v As Long) As Long
    If v < 0 Then Exit Function
    If v > 999 Then
        p_ClampSegment = 999
    Else
        p_ClampSegment = v
    End If
End Function
