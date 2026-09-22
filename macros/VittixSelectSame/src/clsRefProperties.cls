VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
END
Attribute VB_Name = "clsRefProperties"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = False
Option Explicit

' ===========================================================================
' clsRefProperties
'
' Caches every property of the reference shape that the comparison engine
' might need, so the scan loop never has to re-read them from the reference
' object for each of (possibly) thousands of candidates. See section 26 of
' the spec ("Reference Properties").
'
' Populate with PopulateFromShape, called once when the dialog opens (and
' again if the user re-runs the macro on a new selection).
' ===========================================================================

Public ShapeRef            As Shape
Public ObjType             As Long      ' cdrShapeType
Public TypeLabel           As String

Public Width               As Double    ' document units (mm, see mdlUnits)
Public Height              As Double
Public Rotation            As Double    ' degrees

Public HasFill             As Boolean
Public FillType            As Long      ' cdrFillType
Public FillColor           As New Color
Public FountainType        As Long      ' cdrFountainFillType
Public FountainAngle       As Double
Public FountainSteps       As Long
Public FountainBeginColor  As New Color
Public FountainEndColor    As New Color
Public PatternType         As Long

Public HasOutline          As Boolean
Public OutlineColor        As New Color
Public OutlineWidth        As Double

Public NodeCount           As Long
Public NodeCountValid      As Boolean

Public TransparencyAmount  As Double
Public TransparencyValid   As Boolean

Public IsText              As Boolean
Public TextIsArtistic      As Boolean
Public TextFontName        As String
Public TextFontSize        As Double
Public TextBold            As Boolean
Public TextItalic          As Boolean
Public TextContent         As String
Public TextValid           As Boolean

Public IsBitmap            As Boolean
Public BitmapResX          As Double
Public BitmapResY          As Double
Public BitmapPxWidth       As Long
Public BitmapPxHeight      As Long
Public BitmapColorMode     As Long
Public BitmapValid         As Boolean

' Call this once the document's Unit has already been normalized
' (see mdlUnits.NormalizeToMillimeters) so Width/Height/OutlineWidth
' come back in a known, consistent unit.
Public Sub PopulateFromShape(sh As Shape)

    Set ShapeRef = sh

    On Error Resume Next
    ObjType = sh.Type
    TypeLabel = mdlCompare.GetObjectType(sh)
    Width = sh.SizeWidth
    Height = sh.SizeHeight
    Rotation = sh.RotationAngle
    On Error GoTo 0

    ' --- Fill -------------------------------------------------------------
    HasFill = False
    FillType = cdrNoFill
    On Error Resume Next
    FillType = sh.Fill.Type
    On Error GoTo 0

    If FillType <> cdrNoFill Then
        HasFill = True
        Select Case FillType
            Case cdrUniformFill
                On Error Resume Next
                FillColor.CopyAssign sh.Fill.UniformColor
                On Error GoTo 0

            Case cdrFountainFill
                On Error Resume Next
                FountainType = sh.Fill.Fountain.Type
                FountainAngle = sh.Fill.Fountain.Angle
                FountainSteps = sh.Fill.Fountain.Steps
                FountainBeginColor.CopyAssign sh.Fill.Fountain.BeginColor
                FountainEndColor.CopyAssign sh.Fill.Fountain.EndColor
                On Error GoTo 0

            Case cdrPatternFill
                On Error Resume Next
                PatternType = sh.Fill.Pattern.Type
                On Error GoTo 0

            Case Else
                ' Mesh / Texture / PostScript fills: no further comparison
                ' is attempted (see IsSameFillColor in mdlCompare). This
                ' does not crash the macro; it is a documented limitation.
        End Select
    End If

    ' --- Outline ------------------------------------------------------------
    HasOutline = False
    On Error Resume Next
    If Not sh.Outline Is Nothing Then
        If sh.Outline.Type <> cdrNoOutline Then
            HasOutline = True
            OutlineColor.CopyAssign sh.Outline.Color
            OutlineWidth = sh.Outline.Width
        End If
    End If
    On Error GoTo 0

    ' --- Nodes ----------------------------------------------------------------
    NodeCountValid = False
    On Error Resume Next
    Err.Clear
    NodeCount = sh.Curve.Nodes.Count
    If Err.Number = 0 Then NodeCountValid = True
    On Error GoTo 0

    ' --- Transparency ------------------------------------------------------------
    TransparencyValid = False
    On Error Resume Next
    Err.Clear
    TransparencyAmount = sh.Fill.Transparency.Amount
    If Err.Number = 0 Then TransparencyValid = True
    On Error GoTo 0

    ' --- Text -------------------------------------------------------------------
    IsText = (ObjType = cdrTextShape)
    TextValid = False
    If IsText Then
        On Error Resume Next
        Err.Clear
        TextIsArtistic = (sh.Text.Type = cdrArtisticText)
        TextFontName = sh.Text.Story.Font
        TextFontSize = sh.Text.Story.Size
        TextBold = sh.Text.Story.Bold
        TextItalic = sh.Text.Story.Italic
        TextContent = sh.Text.Story.Text
        If Err.Number = 0 Then TextValid = True
        On Error GoTo 0
    End If

    ' --- Bitmap -------------------------------------------------------------------
    IsBitmap = (ObjType = cdrBitmapShape)
    BitmapValid = False
    If IsBitmap Then
        On Error Resume Next
        Err.Clear
        sh.Bitmap.GetResolution BitmapResX, BitmapResY
        BitmapPxWidth = sh.Bitmap.SizeWidth
        BitmapPxHeight = sh.Bitmap.SizeHeight
        BitmapColorMode = sh.Bitmap.Mode
        If Err.Number = 0 Then BitmapValid = True
        On Error GoTo 0
    End If

End Sub



