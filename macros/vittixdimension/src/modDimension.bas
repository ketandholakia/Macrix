Attribute VB_Name = "modDimension"
Option Explicit
Option Private Module
Rem modDimension.bas - orchestration for measuring and placing dimension labels

Public Type VDT_DimensionInfo
    WidthMM As Double
    HeightMM As Double
    AreaMM2 As Double
    PerimeterMM As Double
    ObjectCount As Long
End Type

Public Sub ShowDimensionForm()
    ' frmDimension is a REAL design-time form in the project (its controls and
    ' code-behind are built by tooling and exported back to forms\frmDimension.frm),
    ' so it is referenced directly. UserForm_Initialize -> LoadFormState loads the
    ' current settings. No VBE / runtime form generation is involved.
    frmDimension.Show
End Sub

Public Function BuildSelectionDimensionInfo(sel As ShapeRange) As VDT_DimensionInfo
    Dim info As VDT_DimensionInfo
    Dim shape As shape

    If sel Is Nothing Then
        BuildSelectionDimensionInfo = info
        Exit Function
    End If

    info.ObjectCount = sel.Count
    info.WidthMM = Geometry_GetWidthMM(sel)
    info.HeightMM = Geometry_GetHeightMM(sel)
    info.AreaMM2 = 0#
    info.PerimeterMM = 0#

    ' Area/Perimeter are the sum of each shape's own bounding box, not the
    ' selection's overall bounding box - summing both double-counted these.
    For Each shape In sel
        info.AreaMM2 = info.AreaMM2 + Geometry_GetAreaApproxMM2(shape)
        info.PerimeterMM = info.PerimeterMM + Geometry_GetPerimeterApproxMM(shape)
    Next shape

    BuildSelectionDimensionInfo = info
End Function

Public Sub CreateDimensionsForSelection(ByRef settings As VDT_Settings)
    On Error GoTo ErrHandler
    Dim doc As Document
    Dim sel As ShapeRange
    Dim info As VDT_DimensionInfo
    Dim previousUnit As Long

    Set doc = ActiveDocument
    If doc Is Nothing Then Err.Raise vbObjectError + 100, "CreateDimensionsForSelection", "No active document."
    If Not HasSelection(doc) Then Err.Raise vbObjectError + 101, "CreateDimensionsForSelection", "No objects selected."

    previousUnit = Settings_PreserveDocumentUnit(doc)
    doc.unit = cdrMillimeter
    Set sel = doc.SelectionRange
    info = BuildSelectionDimensionInfo(sel)

    If settings.CreateLayer Then
        Dim dimLayer As Layer
        Set dimLayer = EnsureDimensionLayer(doc)
        dimLayer.Activate
    End If
    CreateSelectionLabels doc, sel, info, settings

CleanExit:
    Settings_RestoreDocumentUnit doc, previousUnit
    Exit Sub
ErrHandler:
    LogError "CreateDimensionsForSelection", Err.Number, Err.Description
    Resume CleanExit
End Sub

Private Sub CreateSelectionLabels(doc As Document, sel As ShapeRange, info As VDT_DimensionInfo, settings As VDT_Settings)
    Dim bounds As VDT_Bounds
    Dim labelText As String
    Dim labelShape As shape

    bounds = Geometry_GetRangeBounds(sel)
    labelText = Label_BuildText(settings.styleName, info, settings)
    Set labelShape = CreateDimensionLabel(doc, Geometry_GetCenterXFromBounds(bounds), Geometry_GetCenterYFromBounds(bounds), labelText, settings, bounds.Width)

    If Not labelShape Is Nothing Then
        Label_PositionShape labelShape, bounds, settings.Position, settings.Gap
    End If
End Sub
