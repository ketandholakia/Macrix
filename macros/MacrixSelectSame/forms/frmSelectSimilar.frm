VERSION 5.00
Begin {C62A69F0-16DC-11CE-9E98-00AA00574A4F} frmSelectSimilar 
   Caption         =   "Select Similar Objects"
   ClientHeight    =   10770
   ClientLeft      =   45
   ClientTop       =   390
   ClientWidth     =   8310.001
   OleObjectBlob   =   "frmSelectSimilar.frx":0000
   StartUpPosition =   1  'CenterOwner
End
Attribute VB_Name = "frmSelectSimilar"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit
Public ReferenceShape As Shape
Private mRefProps As clsRefProperties
Private mMatches As Collection
Private mOriginalUnit As Long
Private mUnitChanged As Boolean
Private Sub UserForm_Initialize()
    mUnitChanged = False
    If ReferenceShape Is Nothing Then Set ReferenceShape = mdlMain.GetReferenceShapeForMacro()
    If ReferenceShape Is Nothing Then
        MsgBox "Please select a single object in CorelDRAW first.", vbExclamation, "Select Similar Objects"
        Unload Me: Exit Sub
    End If
    mOriginalUnit = mdlUnits.NormalizeToMillimeters(ActiveDocument)
    mUnitChanged = True
    Set mRefProps = New clsRefProperties
    mRefProps.PopulateFromShape ReferenceShape
    lblRefInfo.Caption = BuildReferenceInfoText(mRefProps)
    lblMatchCount.Caption = "Matches found: (not yet searched)"
    cmdSelectMatches.Enabled = False
End Sub
Private Sub UserForm_QueryClose(Cancel As Integer, CloseMode As Integer)
    If mUnitChanged Then mdlUnits.RestoreUnit ActiveDocument, mOriginalUnit
End Sub
Private Function BuildReferenceInfoText(ref As clsRefProperties) As String
    Dim sText As String
    sText = "Type:          " & ref.TypeLabel & vbCrLf
    sText = sText & "Size:          " & Format(ref.Width, "0.00") & " x " & Format(ref.Height, "0.00") & " mm" & vbCrLf
    If ref.HasFill Then
        If ref.FillType = cdrUniformFill Then
            sText = sText & "Fill:          Uniform, " & ref.FillColor.Name(True) & vbCrLf
        Else
            sText = sText & "Fill:          Special Fill" & vbCrLf
        End If
    Else
        sText = sText & "Fill:          None" & vbCrLf
    End If
    If ref.HasOutline Then
        sText = sText & "Outline:       " & ref.OutlineColor.Name(True) & " / " & Format(ref.OutlineWidth, "0.00") & " mm" & vbCrLf
    Else
        sText = sText & "Outline:       None" & vbCrLf
    End If
    If ref.NodeCountValid Then
        sText = sText & "Nodes:         " & ref.NodeCount
    Else
        sText = sText & "Nodes:         n/a"
    End If
    BuildReferenceInfoText = sText
End Function
Private Sub SetAllCriteria(ByVal State As Boolean)
    chkObjectType.Value = State: chkFullColor.Value = False
    chkOutlineColor.Value = False: chkOutlineWidth.Value = False
    chkSameSize.Value = False: chkNodeCount.Value = False
    chkWidthOnly.Value = False: chkHeightOnly.Value = False
    chkAspectRatio.Value = False: chkRotation.Value = False
    chkTransparency.Value = False: chkFountainFill.Value = False
    chkPatternFill.Value = False: chkNoFillPresent.Value = False
    chkTextFont.Value = False: chkTextSize.Value = False
    chkTextContent.Value = False: chkTextStyle.Value = False
    chkBitmapResolution.Value = False: chkBitmapWidth.Value = False
    chkBitmapHeight.Value = False: chkBitmapColorMode.Value = False
End Sub
Private Sub cmdExactMatch_Click(): SetAllCriteria False: chkObjectType.Value = True: chkFullColor.Value = True: chkOutlineColor.Value = True: chkOutlineWidth.Value = True: chkSameSize.Value = True: chkNodeCount.Value = True: End Sub
Private Sub cmdAppearance_Click(): SetAllCriteria False: chkFullColor.Value = True: chkOutlineColor.Value = True: chkOutlineWidth.Value = True: End Sub
Private Sub cmdClearPresets_Click(): SetAllCriteria False: End Sub
Private Function BuildCriteria() As TCriteria
    Dim c As TCriteria
    c.ObjectType = chkObjectType.Value: c.FullColor = chkFullColor.Value
    c.OutlineColor = chkOutlineColor.Value: c.OutlineWidth = chkOutlineWidth.Value
    c.SameSize = chkSameSize.Value: c.NodeCount = chkNodeCount.Value
    c.WidthOnly = chkWidthOnly.Value: c.HeightOnly = chkHeightOnly.Value
    c.AspectRatio = chkAspectRatio.Value: c.Rotation = chkRotation.Value
    c.Transparency = chkTransparency.Value: c.FountainFill = chkFountainFill.Value
    c.PatternFill = chkPatternFill.Value: c.NoFillPresent = chkNoFillPresent.Value
    c.TextFont = chkTextFont.Value: c.TextSize = chkTextSize.Value
    c.TextContent = chkTextContent.Value: c.TextStyle = chkTextStyle.Value
    c.BitmapResolution = chkBitmapResolution.Value: c.BitmapWidth = chkBitmapWidth.Value
    c.BitmapHeight = chkBitmapHeight.Value: c.BitmapColorMode = chkBitmapColorMode.Value
    c.SizeTolerance = SafeCDbl(txtSizeTolerance.Text, 0.001)
    c.OutlineTolerance = SafeCDbl(txtOutlineTolerance.Text, 0.001)
    c.RotationTolerance = SafeCDbl(txtRotationTolerance.Text, 0.1)
    BuildCriteria = c
End Function
Private Function SafeCDbl(ByVal ValText As String, ByVal DefaultVal As Double) As Double
    If IsNumeric(ValText) Then SafeCDbl = CDbl(ValText) Else SafeCDbl = DefaultVal
End Function
Private Function BuildScanOptions() As TScanOptions
    Dim o As TScanOptions
    If optCurrentPage.Value Then
        o.Scope = SCOPE_CURRENT_PAGE
    ElseIf optAllPages.Value Then
        o.Scope = SCOPE_ALL_PAGES
    Else
        o.Scope = SCOPE_CURRENT_SELECTION
    End If
    o.SearchInsideGroups = chkSearchInsideGroups.Value: o.IncludeLocked = chkIncludeLocked.Value
    o.IncludeHidden = chkIncludeHidden.Value: o.IncludeReference = chkIncludeReference.Value: o.AddToExisting = chkAddToExisting.Value
    BuildScanOptions = o
End Function
Private Sub cmdFindMatches_Click()
    Dim crit As TCriteria: crit = BuildCriteria()
    If Not mdlCompare.AtLeastOneCriterionSelected(crit) Then MsgBox "Please select at least one criterion.", vbExclamation, "Select Similar Objects": Exit Sub
    Dim scanOpts As TScanOptions: scanOpts = BuildScanOptions()
    On Error Resume Next: ActiveDocument.BeginCommandGroup "Select Similar Objects": Optimization = True: On Error GoTo 0
    Dim candidates As Collection: Set candidates = mdlScan.CollectCandidates(scanOpts)
    Set mMatches = New Collection
    Dim scannedCount As Long, unsupportedCount As Long, sh As Shape, isMatch As Boolean
    For Each sh In candidates
        If Not (sh Is ReferenceShape) Then
            scannedCount = scannedCount + 1
            On Error Resume Next: Err.Clear
            isMatch = mdlCompare.IsSimilarObject(mRefProps, sh, crit)
            If Err.Number <> 0 Then unsupportedCount = unsupportedCount + 1: isMatch = False
            On Error GoTo 0
            If isMatch Then mMatches.Add sh
        End If
    Next sh
    On Error Resume Next: Optimization = False: ActiveWindow.Refresh: ActiveDocument.EndCommandGroup: On Error GoTo 0
    lblMatchCount.Caption = "Matches found: " & mMatches.Count & "   (scanned " & scannedCount & ", unsupported " & unsupportedCount & ")"
    cmdSelectMatches.Enabled = (mMatches.Count > 0)
End Sub
Private Sub cmdSelectMatches_Click()
    If mMatches Is Nothing Then Exit Sub
    Dim scanOpts As TScanOptions: scanOpts = BuildScanOptions()
    Dim unlockedCount As Long: unlockedCount = mdlScan.SelectMatches(mMatches, ReferenceShape, scanOpts)
    If unlockedCount > 0 Then MsgBox unlockedCount & " unlocked object(s) were selected.", vbInformation, "Select Similar Objects"
End Sub
Private Sub cmdClose_Click(): Unload Me: End Sub
