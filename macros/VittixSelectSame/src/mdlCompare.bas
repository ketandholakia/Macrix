Attribute VB_Name = "mdlCompare"
Option Explicit

' ===========================================================================
' mdlCompare
'
' All shape-comparison logic. Every IsSameXxx function is defensive: it
' wraps CorelDRAW API access in error handling and returns False (rather
' than raising) when a property cannot be read on a given candidate. This
' means unsupported object types (bitmaps, text, groups, PowerClips, mesh
' fills, etc.) simply fail that one criterion instead of crashing the scan
' (see spec section 18, "Robust Error Handling").
'
' IsSimilarObject() is the single entry point used by the scanner. It only
' evaluates criteria the user actually enabled, and it exits as soon as any
' enabled criterion fails (cheapest checks - ObjectType - first), which
' keeps large documents fast (spec section 17, "Performance").
' ===========================================================================

' ---------------------------------------------------------------------------
' Friendly type name, also used to sub-classify Artistic vs Paragraph text.
' ---------------------------------------------------------------------------
Public Function GetObjectType(sh As Shape) As String
    On Error GoTo EH
    Select Case sh.Type
        Case cdrRectangleShape:    GetObjectType = "Rectangle"
        Case cdrEllipseShape:      GetObjectType = "Ellipse"
        Case cdrPolygonShape:      GetObjectType = "Polygon"
        Case cdrCurveShape:        GetObjectType = "Curve"
        Case cdrTextShape
            If sh.Text.Type = cdrArtisticText Then
                GetObjectType = "Artistic Text"
            Else
                GetObjectType = "Paragraph Text"
            End If
        Case cdrBitmapShape:       GetObjectType = "Bitmap"
        Case cdrGroupShape:        GetObjectType = "Group"
        Case cdrOLEObjectShape:    GetObjectType = "OLE Object"
        Case cdrConnectorShape:    GetObjectType = "Connector"
        Case cdrCustomShape:       GetObjectType = "Custom Shape"
        Case Else:                 GetObjectType = "Shape (type " & CStr(sh.Type) & ")"
    End Select
    Exit Function
EH:
    GetObjectType = "Unknown"
End Function

Public Function IsSameObjectType(ref As clsRefProperties, cand As Shape) As Boolean
    On Error GoTo EH
    If ref.ObjType <> cand.Type Then
        IsSameObjectType = False
        Exit Function
    End If
    If ref.ObjType = cdrTextShape Then
        IsSameObjectType = (ref.TextIsArtistic = (cand.Text.Type = cdrArtisticText))
    Else
        IsSameObjectType = True
    End If
    Exit Function
EH:
    IsSameObjectType = False
End Function

' ---------------------------------------------------------------------------
' Full Color: compares the actual fill, not a displayed color name.
' ---------------------------------------------------------------------------
Public Function IsSameFillColor(ref As clsRefProperties, cand As Shape, crit As TCriteria) As Boolean
    On Error GoTo EH

    Dim candFillType As Long
    candFillType = cdrNoFill
    candFillType = cand.Fill.Type

    If ref.FillType <> candFillType Then
        IsSameFillColor = False
        Exit Function
    End If

    Select Case ref.FillType
        Case cdrNoFill
            IsSameFillColor = True

        Case cdrUniformFill
            IsSameFillColor = ref.FillColor.IsSame(cand.Fill.UniformColor)

        Case cdrFountainFill
            Dim candF As FountainFill
            Set candF = cand.Fill.Fountain
            IsSameFillColor = (ref.FountainType = candF.Type) _
                And WithinTolerance(NormalizeAngle(ref.FountainAngle), NormalizeAngle(candF.Angle), crit.RotationTolerance) _
                And (ref.FountainSteps = candF.Steps) _
                And ref.FountainBeginColor.IsSame(candF.BeginColor) _
                And ref.FountainEndColor.IsSame(candF.EndColor)

        Case cdrPatternFill
            ' CorelDRAW's Pattern fill API differs between 2-color, full-color,
            ' and bitmap patterns; a reliable cross-version comparison is
            ' limited to the pattern type. See README for how to extend this.
            On Error Resume Next
            IsSameFillColor = (ref.PatternType = cand.Fill.Pattern.Type)
            If Err.Number <> 0 Then IsSameFillColor = False
            On Error GoTo 0

        Case Else
            ' Mesh / Texture / PostScript fill: types already matched above;
            ' deeper comparison is not attempted. Treated as a match on type
            ' only, per spec 4 ("handle safely without crashing").
            IsSameFillColor = True
    End Select
    Exit Function
EH:
    IsSameFillColor = False
End Function

Public Function IsSameOutlineColor(ref As clsRefProperties, cand As Shape) As Boolean
    On Error GoTo EH

    Dim candHasOutline As Boolean
    candHasOutline = (Not cand.Outline Is Nothing)
    If candHasOutline Then candHasOutline = (cand.Outline.Type <> cdrNoOutline)

    If ref.HasOutline <> candHasOutline Then
        IsSameOutlineColor = False
        Exit Function
    End If
    If Not ref.HasOutline Then
        IsSameOutlineColor = True ' both have no outline
        Exit Function
    End If
    IsSameOutlineColor = ref.OutlineColor.IsSame(cand.Outline.Color)
    Exit Function
EH:
    IsSameOutlineColor = False
End Function

Public Function IsSameOutlineWidth(ref As clsRefProperties, cand As Shape, tolerance As Double) As Boolean
    On Error GoTo EH

    Dim candHasOutline As Boolean
    candHasOutline = (Not cand.Outline Is Nothing)
    If candHasOutline Then candHasOutline = (cand.Outline.Type <> cdrNoOutline)

    If ref.HasOutline <> candHasOutline Then
        IsSameOutlineWidth = False
        Exit Function
    End If
    If Not ref.HasOutline Then
        IsSameOutlineWidth = True
        Exit Function
    End If
    IsSameOutlineWidth = WithinTolerance(ref.OutlineWidth, cand.Outline.Width, tolerance)
    Exit Function
EH:
    IsSameOutlineWidth = False
End Function

' ---------------------------------------------------------------------------
' Size. Per spec section 7:
'   - "Ignore Rotation" checked  -> compare the shape's own SizeWidth/Height
'                                   (rotation-independent "geometric" size)
'   - "Ignore Rotation" unchecked -> compare the on-page bounding box, which
'                                    DOES change as an object is rotated
' ---------------------------------------------------------------------------
Public Sub GetComparisonSize(sh As Shape, ByVal ignoreRotation As Boolean, ByRef w As Double, ByRef h As Double)
    On Error GoTo EH
    If ignoreRotation Then
        w = sh.SizeWidth
        h = sh.SizeHeight
    Else
        Dim x1 As Double, y1 As Double, x2 As Double, y2 As Double
        sh.GetBoundingBox x1, y1, x2, y2
        w = Abs(x2 - x1)
        h = Abs(y2 - y1)
    End If
    Exit Sub
EH:
    w = -1: h = -1
End Sub

Public Function IsSameSize(ref As clsRefProperties, cand As Shape, tolerance As Double, ignoreRotation As Boolean) As Boolean
    Dim w As Double, h As Double
    Dim refW As Double, refH As Double
    On Error GoTo EH

    If ignoreRotation Then
        refW = ref.Width: refH = ref.Height
    Else
        Dim x1 As Double, y1 As Double, x2 As Double, y2 As Double
        ref.ShapeRef.GetBoundingBox x1, y1, x2, y2
        refW = Abs(x2 - x1): refH = Abs(y2 - y1)
    End If

    GetComparisonSize cand, ignoreRotation, w, h
    If w < 0 Then
        IsSameSize = False
        Exit Function
    End If

    IsSameSize = WithinTolerance(refW, w, tolerance) And WithinTolerance(refH, h, tolerance)
    Exit Function
EH:
    IsSameSize = False
End Function

Public Function IsSameWidth(ref As clsRefProperties, cand As Shape, tolerance As Double, ignoreRotation As Boolean) As Boolean
    Dim w As Double, h As Double
    On Error GoTo EH
    GetComparisonSize cand, ignoreRotation, w, h
    If w < 0 Then
        IsSameWidth = False
    Else
        Dim refW As Double
        If ignoreRotation Then
            refW = ref.Width
        Else
            Dim x1 As Double, y1 As Double, x2 As Double, y2 As Double
            ref.ShapeRef.GetBoundingBox x1, y1, x2, y2
            refW = Abs(x2 - x1)
        End If
        IsSameWidth = WithinTolerance(refW, w, tolerance)
    End If
    Exit Function
EH:
    IsSameWidth = False
End Function

Public Function IsSameHeight(ref As clsRefProperties, cand As Shape, tolerance As Double, ignoreRotation As Boolean) As Boolean
    Dim w As Double, h As Double
    On Error GoTo EH
    GetComparisonSize cand, ignoreRotation, w, h
    If h < 0 Then
        IsSameHeight = False
    Else
        Dim refH As Double
        If ignoreRotation Then
            refH = ref.Height
        Else
            Dim x1 As Double, y1 As Double, x2 As Double, y2 As Double
            ref.ShapeRef.GetBoundingBox x1, y1, x2, y2
            refH = Abs(y2 - y1)
        End If
        IsSameHeight = WithinTolerance(refH, h, tolerance)
    End If
    Exit Function
EH:
    IsSameHeight = False
End Function

Public Function IsSameNodeCount(ref As clsRefProperties, cand As Shape) As Boolean
    On Error GoTo EH
    If Not ref.NodeCountValid Then
        IsSameNodeCount = False
        Exit Function
    End If
    Dim candCount As Long
    Err.Clear
    candCount = cand.Curve.Nodes.Count
    If Err.Number <> 0 Then
        IsSameNodeCount = False
    Else
        IsSameNodeCount = (candCount = ref.NodeCount)
    End If
    Exit Function
EH:
    IsSameNodeCount = False
End Function

Public Function IsSameRotation(ref As clsRefProperties, cand As Shape, tolerance As Double) As Boolean
    On Error GoTo EH
    IsSameRotation = WithinTolerance(NormalizeAngle(ref.Rotation), NormalizeAngle(cand.RotationAngle), tolerance)
    Exit Function
EH:
    IsSameRotation = False
End Function

Public Function IsSameAspectRatio(ref As clsRefProperties, cand As Shape, tolerance As Double) As Boolean
    On Error GoTo EH
    If ref.Height = 0 Then
        IsSameAspectRatio = False
        Exit Function
    End If
    If cand.SizeHeight = 0 Then
        IsSameAspectRatio = False
        Exit Function
    End If
    Dim refRatio As Double, candRatio As Double
    refRatio = ref.Width / ref.Height
    candRatio = cand.SizeWidth / cand.SizeHeight
    IsSameAspectRatio = (Abs(refRatio - candRatio) <= tolerance)
    Exit Function
EH:
    IsSameAspectRatio = False
End Function

Public Function IsSameTransparency(ref As clsRefProperties, cand As Shape, tolerance As Double) As Boolean
    On Error GoTo EH
    If Not ref.TransparencyValid Then
        IsSameTransparency = False
        Exit Function
    End If
    Dim candAmount As Double
    Err.Clear
    candAmount = cand.Fill.Transparency.Amount
    If Err.Number <> 0 Then
        IsSameTransparency = False
    Else
        IsSameTransparency = WithinTolerance(ref.TransparencyAmount, candAmount, tolerance)
    End If
    Exit Function
EH:
    IsSameTransparency = False
End Function

Public Function IsSameFillPresence(ref As clsRefProperties, cand As Shape) As Boolean
    On Error GoTo EH
    Dim candHasFill As Boolean
    candHasFill = (cand.Fill.Type <> cdrNoFill)
    IsSameFillPresence = (ref.HasFill = candHasFill)
    Exit Function
EH:
    IsSameFillPresence = False
End Function

Public Function IsSameOutlinePresence(ref As clsRefProperties, cand As Shape) As Boolean
    On Error GoTo EH
    Dim candHasOutline As Boolean
    candHasOutline = (Not cand.Outline Is Nothing)
    If candHasOutline Then candHasOutline = (cand.Outline.Type <> cdrNoOutline)
    IsSameOutlinePresence = (ref.HasOutline = candHasOutline)
    Exit Function
EH:
    IsSameOutlinePresence = False
End Function

' ---------------------------------------------------------------------------
' Text properties. Only meaningful when the reference is a text shape; the
' caller (IsSimilarObject) is responsible for only asking for this when
' ref.IsText is True. Font/size read here come from the story's dominant
' (first-run) formatting - a run with mixed formatting is a known,
' documented limitation (spec section 20 treats these as separate opt-in
' criteria for exactly this reason).
' ---------------------------------------------------------------------------
Public Function IsSameTextProperties(ref As clsRefProperties, cand As Shape, crit As TCriteria) As Boolean
    On Error GoTo EH

    If cand.Type <> cdrTextShape Then
        IsSameTextProperties = False
        Exit Function
    End If
    If Not ref.TextValid Then
        IsSameTextProperties = False
        Exit Function
    End If

    Dim ok As Boolean
    ok = True

    If crit.TextFont Then
        ok = ok And (StrComp(ref.TextFontName, cand.Text.Story.Font, vbTextCompare) = 0)
    End If
    If ok And crit.TextSize Then
        ok = ok And WithinTolerance(ref.TextFontSize, cand.Text.Story.Size, 0.01)
    End If
    If ok And crit.TextContent Then
        ok = ok And (ref.TextContent = cand.Text.Story.Text)
    End If
    If ok And crit.TextStyle Then
        ok = ok And (ref.TextBold = cand.Text.Story.Bold) And (ref.TextItalic = cand.Text.Story.Italic)
    End If

    IsSameTextProperties = ok
    Exit Function
EH:
    IsSameTextProperties = False
End Function

' ---------------------------------------------------------------------------
' Bitmap properties. Only meaningful when the reference is a bitmap.
' Note: exact Bitmap object member names (SizeWidth/SizeHeight/Mode/
' GetResolution) have varied slightly across CorelDRAW releases - verify
' against Object Browser (F2) for the target version before relying on
' this in a version you have not tested against (see README).
' ---------------------------------------------------------------------------
Public Function IsSameBitmapProperties(ref As clsRefProperties, cand As Shape, crit As TCriteria) As Boolean
    On Error GoTo EH

    If cand.Type <> cdrBitmapShape Then
        IsSameBitmapProperties = False
        Exit Function
    End If
    If Not ref.BitmapValid Then
        IsSameBitmapProperties = False
        Exit Function
    End If

    Dim ok As Boolean
    ok = True

    If crit.BitmapResolution Then
        Dim rx As Double, ry As Double
        cand.Bitmap.GetResolution rx, ry
        ok = ok And WithinTolerance(ref.BitmapResX, rx, 0.5) And WithinTolerance(ref.BitmapResY, ry, 0.5)
    End If
    If ok And crit.BitmapWidth Then
        ok = ok And (ref.BitmapPxWidth = cand.Bitmap.SizeWidth)
    End If
    If ok And crit.BitmapHeight Then
        ok = ok And (ref.BitmapPxHeight = cand.Bitmap.SizeHeight)
    End If
    If ok And crit.BitmapColorMode Then
        ok = ok And (ref.BitmapColorMode = cand.Bitmap.Mode)
    End If

    IsSameBitmapProperties = ok
    Exit Function
EH:
    IsSameBitmapProperties = False
End Function

' ---------------------------------------------------------------------------
' "Fountain Fill" / "Pattern Fill" (section 3, additional options) are
' lightweight filters independent of Full Color: they only require the
' candidate to use the same fill *category* as the reference, useful when
' the user wants "anything with a gradient" rather than an exact gradient
' match.
' ---------------------------------------------------------------------------
Public Function IsFountainFillMatch(ref As clsRefProperties, cand As Shape) As Boolean
    On Error GoTo EH
    IsFountainFillMatch = (ref.FillType = cdrFountainFill) And (cand.Fill.Type = cdrFountainFill)
    Exit Function
EH:
    IsFountainFillMatch = False
End Function

Public Function IsPatternFillMatch(ref As clsRefProperties, cand As Shape) As Boolean
    On Error GoTo EH
    IsPatternFillMatch = (ref.FillType = cdrPatternFill) And (cand.Fill.Type = cdrPatternFill)
    Exit Function
EH:
    IsPatternFillMatch = False
End Function

' ---------------------------------------------------------------------------
' Central dispatcher. Cheapest / most-discriminating checks run first so
' that a mismatch short-circuits before any expensive property is read
' (fill/outline/node reads are skipped entirely when their checkbox is off).
' ---------------------------------------------------------------------------
Public Function IsSimilarObject(ref As clsRefProperties, candidateShape As Shape, crit As TCriteria) As Boolean
    On Error GoTo EH

    If crit.ObjectType Then
        If Not IsSameObjectType(ref, candidateShape) Then GoTo NoMatch
    End If

    If crit.SameSize Then
        If Not IsSameSize(ref, candidateShape, crit.SizeTolerance, crit.IgnoreRotationSize) Then GoTo NoMatch
    End If
    If crit.WidthOnly Then
        If Not IsSameWidth(ref, candidateShape, crit.SizeTolerance, crit.IgnoreRotationSize) Then GoTo NoMatch
    End If
    If crit.HeightOnly Then
        If Not IsSameHeight(ref, candidateShape, crit.SizeTolerance, crit.IgnoreRotationSize) Then GoTo NoMatch
    End If
    If crit.AspectRatio Then
        If Not IsSameAspectRatio(ref, candidateShape, 0.01) Then GoTo NoMatch
    End If
    If crit.Rotation Then
        If Not IsSameRotation(ref, candidateShape, crit.RotationTolerance) Then GoTo NoMatch
    End If

    If crit.NoFillPresent Then
        If Not IsSameFillPresence(ref, candidateShape) Then GoTo NoMatch
    End If
    If crit.NoOutlinePresent Then
        If Not IsSameOutlinePresence(ref, candidateShape) Then GoTo NoMatch
    End If

    If crit.FullColor Then
        If Not IsSameFillColor(ref, candidateShape, crit) Then GoTo NoMatch
    End If
    If crit.OutlineColor Then
        If Not IsSameOutlineColor(ref, candidateShape) Then GoTo NoMatch
    End If
    If crit.OutlineWidth Then
        If Not IsSameOutlineWidth(ref, candidateShape, crit.OutlineTolerance) Then GoTo NoMatch
    End If
    If crit.Transparency Then
        If Not IsSameTransparency(ref, candidateShape, 0.5) Then GoTo NoMatch
    End If

    If crit.NodeCount Then
        If Not IsSameNodeCount(ref, candidateShape) Then GoTo NoMatch
    End If

    If crit.FountainFill Then
        If Not IsFountainFillMatch(ref, candidateShape) Then GoTo NoMatch
    End If
    If crit.PatternFill Then
        If Not IsPatternFillMatch(ref, candidateShape) Then GoTo NoMatch
    End If

    If ref.IsText And (crit.TextFont Or crit.TextSize Or crit.TextContent Or crit.TextStyle) Then
        If Not IsSameTextProperties(ref, candidateShape, crit) Then GoTo NoMatch
    End If

    If ref.IsBitmap And (crit.BitmapResolution Or crit.BitmapWidth Or crit.BitmapHeight Or crit.BitmapColorMode) Then
        If Not IsSameBitmapProperties(ref, candidateShape, crit) Then GoTo NoMatch
    End If

    IsSimilarObject = True
    Exit Function

NoMatch:
    IsSimilarObject = False
    Exit Function
EH:
    ' Any unexpected error on this candidate simply excludes it - the scan
    ' continues (spec section 18).
    IsSimilarObject = False
End Function

' ---------------------------------------------------------------------------
' Small numeric helpers
' ---------------------------------------------------------------------------
Public Function WithinTolerance(ByVal a As Double, ByVal b As Double, ByVal tolerance As Double) As Boolean
    WithinTolerance = (Abs(a - b) <= tolerance)
End Function

Public Function NormalizeAngle(ByVal degrees As Double) As Double
    Dim d As Double
    d = degrees Mod 360
    If d < 0 Then d = d + 360
    NormalizeAngle = d
End Function

Public Function AtLeastOneCriterionSelected(crit As TCriteria) As Boolean
    AtLeastOneCriterionSelected = crit.ObjectType Or crit.FullColor Or crit.OutlineColor Or _
        crit.OutlineWidth Or crit.SameSize Or crit.NodeCount Or crit.WidthOnly Or crit.HeightOnly Or _
        crit.AspectRatio Or crit.Rotation Or crit.Transparency Or crit.FountainFill Or crit.PatternFill Or _
        crit.NoFillPresent Or crit.NoOutlinePresent Or crit.TextFont Or crit.TextSize Or crit.TextContent Or _
        crit.TextStyle Or crit.BitmapResolution Or crit.BitmapWidth Or crit.BitmapHeight Or crit.BitmapColorMode
End Function


