Attribute VB_Name = "modUtils"
Option Explicit
Option Private Module
Rem modUtils.bas - logging and common helpers

Public Sub LogError(procName As String, errNum As Long, errDesc As String)
    On Error Resume Next
    MsgBox "Error in " & procName & ": [" & errNum & "] " & errDesc, vbCritical, "Macrix Error"
End Sub

Public Function SafeTrim(value As Variant) As String
    If IsError(value) Or IsNull(value) Then
        SafeTrim = vbNullString
    Else
        SafeTrim = Trim$(CStr(value))
    End If
End Function

Public Function TryParseDouble(textValue As String, ByRef result As Double) As Boolean
    On Error GoTo Fail
    Dim normalized As String
    normalized = SafeTrim(textValue)
    If Len(normalized) = 0 Then Exit Function
    normalized = Replace(normalized, ",", ".")
    result = CDbl(normalized)
    TryParseDouble = True
    Exit Function
Fail:
    TryParseDouble = False
End Function

Public Function ClampDouble(value As Double, minValue As Double, maxValue As Double) As Double
    If value < minValue Then
        ClampDouble = minValue
    ElseIf value > maxValue Then
        ClampDouble = maxValue
    Else
        ClampDouble = value
    End If
End Function

Public Function HasSelection(doc As Document) As Boolean
    On Error Resume Next
    HasSelection = Not doc Is Nothing And doc.SelectionRange.Count > 0
End Function

Public Function ConvertAreaFromMM2(valueMM2 As Double, unit As VDTUnit) As Double
    Select Case unit
        Case UNIT_MM: ConvertAreaFromMM2 = valueMM2
        Case UNIT_CM: ConvertAreaFromMM2 = valueMM2 / 100#
        Case UNIT_M: ConvertAreaFromMM2 = valueMM2 / 1000000#
        Case UNIT_IN: ConvertAreaFromMM2 = valueMM2 / (25.4 * 25.4)
        Case UNIT_FT: ConvertAreaFromMM2 = valueMM2 / (304.8 * 304.8)
        Case Else: ConvertAreaFromMM2 = valueMM2
    End Select
End Function
Public Function FormatArea(valueMM2 As Double, unit As VDTUnit, Optional decimals As Integer = 2) As String
    FormatArea = FormatValue(ConvertAreaFromMM2(valueMM2, unit), decimals) & " sq " & UnitSuffix(unit)
End Function

'--------------------------------------------------------------
' Colours
'--------------------------------------------------------------
' Parse a colour string into an RGB Long (as produced by RGB()). Accepts
' "#RRGGBB", "RRGGBB", "#RGB" or "R,G,B". Returns -1 when unparseable.
Public Function ColorFromHex(ByVal value As String) As Long
    On Error GoTo Bad
    Dim s As String
    Dim parts() As String
    Dim rr As Long, gg As Long, bb As Long

    ColorFromHex = -1
    s = SafeTrim(value)
    If Len(s) = 0 Then Exit Function

    If InStr(s, ",") > 0 Then
        parts = Split(s, ",")
        If UBound(parts) <> 2 Then Exit Function
        rr = CLng(SafeTrim(parts(0)))
        gg = CLng(SafeTrim(parts(1)))
        bb = CLng(SafeTrim(parts(2)))
        ColorFromHex = RGB(rr, gg, bb)
        Exit Function
    End If

    s = Replace(s, "#", "")
    s = Replace(s, " ", "")
    If Len(s) = 3 Then
        s = Mid$(s, 1, 1) & Mid$(s, 1, 1) & Mid$(s, 2, 1) & Mid$(s, 2, 1) & Mid$(s, 3, 1) & Mid$(s, 3, 1)
    End If
    If Len(s) <> 6 Then Exit Function

    rr = HexPair(s, 1)
    gg = HexPair(s, 3)
    bb = HexPair(s, 5)
    If rr < 0 Or gg < 0 Or bb < 0 Then Exit Function
    ColorFromHex = RGB(rr, gg, bb)
    Exit Function
Bad:
    ColorFromHex = -1
End Function

' Two hex digits -> 0..255, or -1 when either character is not hex. Deliberately avoids
' CLng's hex-string parsing: CLng("&HFF&") raises an error, which On Error turned into
' "-1" and silently disabled the caption colour altogether.
Private Function HexPair(ByVal s As String, ByVal startIndex As Long) As Long
    Dim hi As Long, lo As Long
    hi = HexDigit(Mid$(s, startIndex, 1))
    lo = HexDigit(Mid$(s, startIndex + 1, 1))
    If hi < 0 Or lo < 0 Then
        HexPair = -1
    Else
        HexPair = hi * 16 + lo
    End If
End Function
Private Function HexDigit(ByVal ch As String) As Long
    Dim i As Long
    HexDigit = -1
    For i = 1 To Len("0123456789ABCDEF")
        If UCase$(ch) = Mid$("0123456789ABCDEF", i, 1) Then
            HexDigit = i - 1
            Exit Function
        End If
    Next i
End Function

' RGB Long -> "#RRGGBB" (for showing a stored colour in a text box).
Public Function ColorToHex(ByVal rgbValue As Long) As String
    On Error Resume Next
    If rgbValue < 0 Then
        ColorToHex = "#000000"
        Exit Function
    End If
    ColorToHex = "#" & Right$("0" & Hex$(rgbValue And &HFF), 2) & _
        Right$("0" & Hex$((rgbValue \ &H100) And &HFF), 2) & _
        Right$("0" & Hex$((rgbValue \ &H10000) And &HFF), 2)
End Function
