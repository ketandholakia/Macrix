Attribute VB_Name = "mdlUnits"
Option Explicit

' ===========================================================================
' mdlUnits
'
' Shape.SizeWidth / Shape.SizeHeight / Outline.Width in the CorelDRAW VBA
' object model are reported in the ACTIVE DOCUMENT's current ruler unit.
' If the user's document happens to be set to inches, points, pixels, etc.,
' comparisons would silently work correctly too (since both reference and
' candidate are read in the same unit) - but displaying values to the user
' and applying a fixed numeric tolerance (e.g. 0.001) only makes sense in a
' known unit. This module temporarily forces the document to millimeters
' for the duration of a scan, then restores whatever the user had.
'
' This is a display/measurement setting only - it is not tracked by Undo
' and does not modify any shape.
' ===========================================================================

Public Function NormalizeToMillimeters(doc As Document) As Long
    ' Returns the previous unit so the caller can restore it later.
    On Error Resume Next
    NormalizeToMillimeters = doc.Unit
    doc.Unit = cdrMillimeter
    On Error GoTo 0
End Function

Public Sub RestoreUnit(doc As Document, ByVal originalUnit As Long)
    On Error Resume Next
    doc.Unit = originalUnit
    On Error GoTo 0
End Sub

