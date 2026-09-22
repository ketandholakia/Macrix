Attribute VB_Name = "modLabel"
Option Explicit
Rem modLabel.bas - label creation and styling

Public Function Label_BuildText(styleName As String, info As VDT_DimensionInfo, settings As VDT_Settings) As String
    Dim baseText As String
    
    If Trim(settings.Template) <> "" Then
        baseText = BuildTemplateLabel(info, settings)
    Else
        Select Case LCase$(styleName)
            Case "technical": baseText = BuildTechnicalLabel(info, settings)
            Case "engineering": baseText = BuildEngineeringLabel(info, settings)
            Case "packaging": baseText = BuildPackagingLabel(info, settings)
            Case Else: baseText = BuildMinimalLabel(info, settings)
        End Select
        
        If settings.ShowArea Then
            If baseText <> "" Then baseText = baseText & vbCrLf
            baseText = baseText & "Area: " & FormatArea(info.AreaMM2, settings.unit, settings.decimals)
        End If
        
        If settings.ShowPerimeter Then
            If baseText <> "" Then baseText = baseText & vbCrLf
            baseText = baseText & "Perimeter: " & FormatMeasurement(info.PerimeterMM, settings.unit, settings.decimals)
        End If
        
        If settings.ShowObjectCount Then
            If baseText <> "" Then baseText = baseText & vbCrLf
            baseText = baseText & "Objects: " & CStr(info.ObjectCount)
        End If
    End If
    
    Label_BuildText = baseText
End Function

Private Function BuildTemplateLabel(info As VDT_DimensionInfo, settings As VDT_Settings) As String
    Dim res As String
    res = settings.Template
    
    res = Replace(res, "{W}", FormatValue(ConvertFromMM(info.WidthMM, settings.unit), settings.decimals))
    res = Replace(res, "{H}", FormatValue(ConvertFromMM(info.HeightMM, settings.unit), settings.decimals))
    res = Replace(res, "{U}", UnitSuffix(settings.unit))
    res = Replace(res, "{A}", FormatArea(info.AreaMM2, settings.unit, settings.decimals))
    res = Replace(res, "{P}", FormatMeasurement(info.PerimeterMM, settings.unit, settings.decimals))
    res = Replace(res, "{C}", CStr(info.ObjectCount))
    
    BuildTemplateLabel = res
End Function

Public Function CreateDimensionLabel(doc As Document, x As Double, y As Double, text As String, settings As VDT_Settings, Optional objectWidthMM As Double = 0) As shape
    On Error GoTo ErrHandler
    Dim textShape As shape
    Dim boxShape As shape
    Dim finalShape As shape
    If doc Is Nothing Then Exit Function
    Set textShape = CreateArtisticTextShape(doc, x, y, text, CDbl(settings.fontSize), objectWidthMM, settings)
    If textShape Is Nothing Then Exit Function
    
    Set finalShape = textShape
    If settings.BackgroundBox Then
        Set boxShape = CreateBackgroundShape(doc, textShape, settings)
        EnsureTextAboveBackground textShape, boxShape
        Set finalShape = GroupLabelShapes(textShape, boxShape)
    End If
    Set CreateDimensionLabel = finalShape
    Exit Function
ErrHandler:
    LogError "CreateDimensionLabel", Err.Number, Err.Description
End Function

Public Sub Label_PositionShape(labelShape As shape, bounds As VDT_Bounds, positionName As String, gapMM As Double)
    Dim targetX As Double
    Dim targetY As Double

    targetX = Geometry_GetCenterXFromBounds(bounds)
    targetY = Geometry_GetCenterYFromBounds(bounds)

    Select Case LCase$(positionName)
        Case "above"
            targetY = bounds.Top + gapMM
        Case "below"
            targetY = bounds.Bottom - gapMM
        Case "left"
            targetX = bounds.Left - gapMM
        Case "right"
            targetX = bounds.Right + gapMM
        Case Else
            ApplyAutoPosition bounds, gapMM, targetX, targetY
    End Select

    MoveShapeToPoint labelShape, targetX, targetY
End Sub

Public Function CreateArtisticTextShape(doc As Document, x As Double, y As Double, text As String, ByVal fontSize As Double, Optional objectWidthMM As Double = 0, Optional settings As VDT_Settings) As shape
    On Error Resume Next
    Dim targetLayer As Layer
    Set targetLayer = ActiveTargetLayer(doc)
    If targetLayer Is Nothing Then Set targetLayer = doc.ActiveLayer
    If targetLayer Is Nothing Then Exit Function

    Set CreateArtisticTextShape = targetLayer.CreateArtisticText(x, y, text)
    If CreateArtisticTextShape Is Nothing Then Exit Function
    
    ' Set initial font
    With CreateArtisticTextShape.Text.Story
        .FontSize = fontSize
        .FontName = "Arial"
        If Not .TextRange Is Nothing Then
            .TextRange.FontSize = fontSize
            .TextRange.FontName = "Arial"
        End If
    End With
    
    ' Scale text width to match percentage of object width
    If objectWidthMM > 0 And settings.TextWidthPercent > 0 Then
        ScaleTextToWidth CreateArtisticTextShape, objectWidthMM, settings.TextWidthPercent
    End If
End Function

Private Sub ScaleTextToWidth(textShape As shape, objectWidthMM As Double, widthPercent As Double)
    On Error Resume Next
    Dim targetWidthMM As Double
    Dim currentWidthMM As Double
    Dim scaleFactor As Double
    Dim newFontSize As Double
    
    targetWidthMM = objectWidthMM * (widthPercent / 100#)
    
    ' Get current text width in document units (MM)
    Dim doc As Document
    Set doc = ActiveDocument
    If doc Is Nothing Then Exit Sub
    
    Dim prevUnit As Long
    prevUnit = doc.unit
    doc.unit = cdrMillimeter
    
    currentWidthMM = textShape.SizeWidth
    doc.unit = prevUnit
    
    If currentWidthMM > 0 Then
        scaleFactor = targetWidthMM / currentWidthMM
        newFontSize = textShape.Text.Story.FontSize * scaleFactor
        
        ' Apply scaled font size
        With textShape.Text.Story
            .FontSize = newFontSize
            If Not .TextRange Is Nothing Then
                .TextRange.FontSize = newFontSize
            End If
        End With
    End If
End Sub

Private Function BuildMinimalLabel(info As VDT_DimensionInfo, settings As VDT_Settings) As String
    Dim parts As String
    parts = ""
    
    If settings.ShowWidth Then
        parts = FormatMeasurement(info.WidthMM, settings.unit, settings.decimals)
    End If
    
    If settings.ShowHeight Then
        If parts <> "" Then parts = parts & " x "
        parts = parts & FormatMeasurement(info.HeightMM, settings.unit, settings.decimals)
    End If
    
    BuildMinimalLabel = parts
End Function

Private Function BuildTechnicalLabel(info As VDT_DimensionInfo, settings As VDT_Settings) As String
    BuildTechnicalLabel = "SIZE" & vbCrLf & BuildMinimalLabel(info, settings)
End Function

Private Function BuildEngineeringLabel(info As VDT_DimensionInfo, settings As VDT_Settings) As String
    Dim parts As String
    parts = ""
    
    If settings.ShowWidth Then
        parts = "W: " & FormatMeasurement(info.WidthMM, settings.unit, settings.decimals)
    End If
    
    If settings.ShowHeight Then
        If parts <> "" Then parts = parts & vbCrLf
        parts = parts & "H: " & FormatMeasurement(info.HeightMM, settings.unit, settings.decimals)
    End If
    
    BuildEngineeringLabel = parts
End Function

Private Function BuildPackagingLabel(info As VDT_DimensionInfo, settings As VDT_Settings) As String
    BuildPackagingLabel = "CUT SIZE" & vbCrLf & BuildMinimalLabel(info, settings)
End Function

Private Function CreateBackgroundShape(doc As Document, textShape As shape, settings As VDT_Settings) As shape
    Dim box As shape
    Dim padX As Double
    Dim padY As Double
    padX = settings.Padding
    padY = settings.Padding
    Set box = CreateRectangleShape(doc, textShape.leftX - padX, textShape.topY + padY, textShape.rightX + padX, textShape.bottomY - padY)
    
    If settings.RoundedBackground Then
        ' CorelDRAW handles corner radius per corner
        box.Rectangle.RadiusUpperLeft = settings.CornerRadius
        box.Rectangle.RadiusUpperRight = settings.CornerRadius
        box.Rectangle.RadiusLowerLeft = settings.CornerRadius
        box.Rectangle.RadiusLowerRight = settings.CornerRadius
    End If
    ' Add a white background fill so you can read the text!
    box.Fill.UniformColor.RGBAssign 255, 255, 255
    Set CreateBackgroundShape = box
End Function

Private Function GroupLabelShapes(ByVal textShape As shape, ByVal boxShape As shape) As shape
    On Error Resume Next
    Dim sr As New ShapeRange
    If Not textShape Is Nothing And Not boxShape Is Nothing Then
        sr.Add textShape
        sr.Add boxShape
        Set GroupLabelShapes = sr.Group
    Else
        Set GroupLabelShapes = textShape
    End If
End Function

Private Sub EnsureTextAboveBackground(ByVal textShape As shape, ByVal boxShape As shape)
    On Error Resume Next
    If Not textShape Is Nothing Then textShape.OrderToFront
    If Not boxShape Is Nothing Then boxShape.OrderToBack
End Sub

Private Function ActiveTargetLayer(doc As Document) As Layer
    On Error Resume Next
    Set ActiveTargetLayer = doc.ActiveLayer
End Function

Private Sub ApplyFontSize(shapeObj As shape, fontSize As Double)
    ' No longer used - font set directly in CreateArtisticTextShape
End Sub

Private Function CreateRectangleShape(doc As Document, leftX As Double, topY As Double, rightX As Double, bottomY As Double) As shape
    On Error GoTo ErrHandler
    Dim targetLayer As Layer
    Set targetLayer = ActiveTargetLayer(doc)
    If targetLayer Is Nothing Then Exit Function
    Set CreateRectangleShape = targetLayer.CreateRectangle(leftX, topY, rightX, bottomY)
    Exit Function
ErrHandler:
    LogError "CreateRectangleShape", Err.Number, Err.Description
End Function

Private Sub MoveShapeToPoint(shapeObj As shape, targetX As Double, targetY As Double)
    On Error Resume Next
    If shapeObj Is Nothing Then Exit Sub
    shapeObj.Move targetX - Geometry_GetCenterXFromShape(shapeObj), targetY - Geometry_GetCenterYFromShape(shapeObj)
End Sub

Private Sub ApplyAutoPosition(bounds As VDT_Bounds, gapMM As Double, ByRef targetX As Double, ByRef targetY As Double)
    If bounds.Width >= bounds.Height Then
        targetY = bounds.Top + gapMM
    Else
        targetX = bounds.Right + gapMM
    End If
End Sub
