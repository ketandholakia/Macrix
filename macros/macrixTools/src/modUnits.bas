Attribute VB_Name = "modUnits"
Option Explicit
Option Private Module
Rem modUnits.bas - unit conversion helpers

Public Enum VDTUnit
    UNIT_MM = 0
    UNIT_CM = 1
    UNIT_M = 2
    UNIT_IN = 3
    UNIT_FT = 4
End Enum

Public Function ConvertToMM(value As Double, unit As VDTUnit) As Double
    Select Case unit
        Case UNIT_MM: ConvertToMM = value
        Case UNIT_CM: ConvertToMM = value * 10#
        Case UNIT_M: ConvertToMM = value * 1000#
        Case UNIT_IN: ConvertToMM = value * 25.4
        Case UNIT_FT: ConvertToMM = value * 304.8
        Case Else: ConvertToMM = value
    End Select
End Function

Public Function ConvertFromMM(value As Double, unit As VDTUnit) As Double
    Select Case unit
        Case UNIT_MM: ConvertFromMM = value
        Case UNIT_CM: ConvertFromMM = value / 10#
        Case UNIT_M: ConvertFromMM = value / 1000#
        Case UNIT_IN: ConvertFromMM = value / 25.4
        Case UNIT_FT: ConvertFromMM = value / 304.8
        Case Else: ConvertFromMM = value
    End Select
End Function

Public Function UnitSuffix(unit As VDTUnit) As String
    Select Case unit
        Case UNIT_MM: UnitSuffix = "mm"
        Case UNIT_CM: UnitSuffix = "cm"
        Case UNIT_M: UnitSuffix = "m"
        Case UNIT_IN: UnitSuffix = "in"
        Case UNIT_FT: UnitSuffix = "ft"
        Case Else: UnitSuffix = "mm"
    End Select
End Function

Public Function UnitToText(unit As VDTUnit) As String
    UnitToText = UCase$(UnitSuffix(unit))
End Function

Public Function TextToUnit(unitText As String) As VDTUnit
    Select Case UCase$(Trim$(unitText))
        Case "MM": TextToUnit = UNIT_MM
        Case "CM": TextToUnit = UNIT_CM
        Case "M": TextToUnit = UNIT_M
        Case "IN": TextToUnit = UNIT_IN
        Case "FT": TextToUnit = UNIT_FT
        Case Else: TextToUnit = UNIT_MM
    End Select
End Function

Public Function FormatValue(value As Double, Optional decimals As Integer = 2) As String
    If decimals > 0 Then
        FormatValue = Format(value, "0." & String(decimals, "0"))
    Else
        FormatValue = Format(value, "0")
    End If
End Function

Public Function FormatMeasurement(valueMM As Double, unit As VDTUnit, Optional decimals As Integer = 2) As String
    FormatMeasurement = FormatValue(ConvertFromMM(valueMM, unit), decimals) & " " & UnitSuffix(unit)
End Function
