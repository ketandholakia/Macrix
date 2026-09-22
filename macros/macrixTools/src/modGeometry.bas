Attribute VB_Name = "modGeometry"
Option Explicit
Option Private Module
Rem modGeometry.bas - geometry helper functions

Public Type VDT_Bounds
    Left As Double
    Top As Double
    Right As Double
    Bottom As Double
    Width As Double
    Height As Double
End Type

Public Function Geometry_GetWidth(s As shape) As Double
    Geometry_GetWidth = s.SizeWidth
End Function

Public Function Geometry_GetHeight(s As shape) As Double
    Geometry_GetHeight = s.SizeHeight
End Function

Public Function Geometry_GetRangeBounds(sel As ShapeRange) As VDT_Bounds
    Dim bounds As VDT_Bounds
    If sel Is Nothing Or sel.Count = 0 Then
        Geometry_GetRangeBounds = bounds
        Exit Function
    End If

    bounds.Left = sel.leftX
    bounds.Top = sel.topY
    bounds.Right = sel.rightX
    bounds.Bottom = sel.bottomY
    bounds.Width = bounds.Right - bounds.Left
    bounds.Height = bounds.Top - bounds.Bottom
    Geometry_GetRangeBounds = bounds
End Function

Public Function Geometry_GetWidthMM(sel As ShapeRange) As Double
    Geometry_GetWidthMM = Geometry_GetRangeBounds(sel).Width
End Function

Public Function Geometry_GetHeightMM(sel As ShapeRange) As Double
    Geometry_GetHeightMM = Geometry_GetRangeBounds(sel).Height
End Function

Public Function Geometry_GetAreaApproxMM2(s As shape) As Double
    Geometry_GetAreaApproxMM2 = s.SizeWidth * s.SizeHeight
End Function

Public Function Geometry_GetPerimeterApproxMM(s As shape) As Double
    Geometry_GetPerimeterApproxMM = 2# * (s.SizeWidth + s.SizeHeight)
End Function

Public Function Geometry_GetCenterXFromBounds(bounds As VDT_Bounds) As Double
    Geometry_GetCenterXFromBounds = bounds.Left + (bounds.Width / 2#)
End Function

Public Function Geometry_GetCenterYFromBounds(bounds As VDT_Bounds) As Double
    Geometry_GetCenterYFromBounds = bounds.Bottom + (bounds.Height / 2#)
End Function

Public Function Geometry_GetCenterXFromShape(shapeObj As shape) As Double
    Geometry_GetCenterXFromShape = (shapeObj.leftX + shapeObj.rightX) / 2#
End Function

Public Function Geometry_GetCenterYFromShape(shapeObj As shape) As Double
    Geometry_GetCenterYFromShape = (shapeObj.topY + shapeObj.bottomY) / 2#
End Function

Public Function Geometry_GetCenterX(sel As ShapeRange) As Double
    Dim bounds As VDT_Bounds
    bounds = Geometry_GetRangeBounds(sel)
    Geometry_GetCenterX = bounds.Left + (bounds.Width / 2#)
End Function

Public Function Geometry_GetCenterY(sel As ShapeRange) As Double
    Dim bounds As VDT_Bounds
    bounds = Geometry_GetRangeBounds(sel)
    Geometry_GetCenterY = bounds.Bottom + (bounds.Height / 2#)
End Function
