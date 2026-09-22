Attribute VB_Name = "modLabel"
Option Explicit
Rem modLabel.bas - label creation and styling

' Base point size used when building a label. The label is then scaled to a
' percentage of the selected object's width, so this is only a starting point.
Private Const DEFAULT_TEXT_POINTS As Double = 12#

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
    Set textShape = CreateArtisticTextShape(doc, x, y, text, settings, objectWidthMM)
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
    Dim rotate90 As Boolean

    targetX = Geometry_GetCenterXFromBounds(bounds)
    targetY = Geometry_GetCenterYFromBounds(bounds)

    Select Case LCase$(Trim$(positionName))
        Case "center"
            ' dead centre of the object
            targetX = Geometry_GetCenterXFromBounds(bounds)
            targetY = Geometry_GetCenterYFromBounds(bounds)
        Case "above"
            targetY = bounds.Top + gapMM
        Case "below"
            targetY = bounds.Bottom - gapMM
        Case "left"
            targetX = bounds.Left - gapMM
            rotate90 = True
        Case "right"
            targetX = bounds.Right + gapMM
            rotate90 = True
        Case Else
            ApplyAutoPosition bounds, gapMM, targetX, targetY
    End Select

    ' Left/Right read vertically: rotate the label, then place its centre on the edge.
    If rotate90 Then RotateShape labelShape, 90
    MoveShapeToPoint labelShape, targetX, targetY
End Sub

Private Sub RotateShape(shapeObj As shape, ByVal angleDeg As Double)
    On Error Resume Next
    If shapeObj Is Nothing Then Exit Sub
    shapeObj.Rotate angleDeg
End Sub

' Build the label text shape. There is no font-size setting any more: the label is
' sized by TextWidthPercent, i.e. a percentage of the selected object's width.
Public Function CreateArtisticTextShape(doc As Document, x As Double, y As Double, text As String, settings As VDT_Settings, Optional objectWidthMM As Double = 0) As shape
    On Error GoTo ErrHandler
    Dim targetLayer As Layer
    Dim sh As shape

    Set targetLayer = ActiveTargetLayer(doc)
    If targetLayer Is Nothing Then Set targetLayer = doc.ActiveLayer
    If targetLayer Is Nothing Then Exit Function

    Set sh = targetLayer.CreateArtisticText(x, y, text)
    Set CreateArtisticTextShape = sh
    If sh Is Nothing Then Exit Function

    ApplyTextFormatting sh

    ' Scale the label to a percentage of the selected object's width.
    If objectWidthMM > 0 And settings.TextWidthPercent > 0 Then
        ScaleTextToWidth sh, objectWidthMM, settings.TextWidthPercent
    End If
    Exit Function
ErrHandler:
    LogError "CreateArtisticTextShape", Err.Number, Err.Description
End Function

' Font name and size live on the Story (a TextRange). IMPORTANT: the size member is
' .Size -- there is no .FontSize / .FontName, and those fail silently.
Private Sub ApplyTextFormatting(sh As shape)
    On Error GoTo ErrHandler
    Dim story As Object
    If sh Is Nothing Then Exit Sub
    Set story = sh.Text.Story
    If story Is Nothing Then Exit Sub
    story.Font = "Arial"
    story.Size = DEFAULT_TEXT_POINTS
    Exit Sub
ErrHandler:
    LogError "ApplyTextFormatting", Err.Number, Err.Description
End Sub

' Scale the label so its rendered width equals `widthPercent` % of the object width.
' CreateDimensionsForSelection sets the document unit to millimetres before labels
' are created, so Shape.SizeWidth and objectWidthMM are both in mm here.
Private Sub ScaleTextToWidth(sh As shape, ByVal objectWidthMM As Double, ByVal widthPercent As Double)
    On Error GoTo ErrHandler
    Dim story As Object
    Dim targetWidthMM As Double
    Dim currentWidthMM As Double
    Dim currentSize As Double
    Dim newSize As Double

    If sh Is Nothing Then Exit Sub
    If objectWidthMM <= 0 Or widthPercent <= 0 Then Exit Sub

    Set story = sh.Text.Story
    If story Is Nothing Then Exit Sub

    currentWidthMM = sh.SizeWidth
    If currentWidthMM <= 0 Then Exit Sub

    currentSize = story.Size
    If currentSize <= 0 Then currentSize = DEFAULT_TEXT_POINTS

    targetWidthMM = objectWidthMM * (widthPercent / 100#)
    If targetWidthMM <= 0 Then Exit Sub

    newSize = currentSize * (targetWidthMM / currentWidthMM)
    If newSize < 0.5 Then newSize = 0.5
    If newSize > 2000# Then newSize = 2000#

    story.Size = newSize
    Exit Sub
ErrHandler:
    LogError "ScaleTextToWidth", Err.Number, Err.Description
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
