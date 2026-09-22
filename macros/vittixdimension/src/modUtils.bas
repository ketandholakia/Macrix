Attribute VB_Name = "modUtils"
Option Explicit
Option Private Module
Rem modUtils.bas - logging and common helpers

Public Sub LogError(procName As String, errNum As Long, errDesc As String)
    On Error Resume Next
    MsgBox "Error in " & procName & ": [" & errNum & "] " & errDesc, vbCritical, "Vittix Error"
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
