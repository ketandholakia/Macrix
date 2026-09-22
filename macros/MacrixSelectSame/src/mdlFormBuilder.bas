Attribute VB_Name = "mdlFormBuilder"
Option Explicit

Public Sub BuildSelectSimilarFormWithCode()
    Dim vbProj As Object
    Set vbProj = Application.VBE.ActiveVBProject
    
    On Error Resume Next
    Dim existingComp As Object
    Set existingComp = vbProj.VBComponents("frmSelectSimilar")
    If Not existingComp Is Nothing Then
        vbProj.VBComponents.Remove existingComp
    End If
    On Error GoTo 0
    
    Dim newFormComp As Object
    Set newFormComp = vbProj.VBComponents.Add(3)
    newFormComp.Name = "frmSelectSimilar"
    
    With newFormComp.Properties
        .Item("Caption").Value = "Select Similar Objects"
        .Item("Width").Value = 420
        .Item("Height").Value = 560
    End With
    
    Dim formDesigner As Object
    Set formDesigner = newFormComp.Designer
    
    Dim fraRef As Object, fraCrit As Object, fraExt As Object, fraScope As Object
    Set fraRef = formDesigner.Controls.Add("Forms.Frame.1", "fraReference")
    With fraRef: .Caption = "Reference Object": .Left = 8: .Top = 8: .Width = 390: .Height = 65: End With
    
    Set fraCrit = formDesigner.Controls.Add("Forms.Frame.1", "fraCriteria")
    With fraCrit: .Caption = "Match Criteria": .Left = 8: .Top = 78: .Width = 390: .Height = 150: End With
    
    Set fraExt = formDesigner.Controls.Add("Forms.Frame.1", "fraExtended")
    With fraExt: .Caption = "Extended Criteria (Text / Bitmap)": .Left = 8: .Top = 232: .Width = 390: .Height = 100: End With
    
    Set fraScope = formDesigner.Controls.Add("Forms.Frame.1", "fraScope")
    With fraScope: .Caption = "Search Scope": .Left = 8: .Top = 336: .Width = 390: .Height = 60: End With
    
    Dim lblRef As Object
    Set lblRef = fraRef.Controls.Add("Forms.Label.1", "lblRefInfo")
    With lblRef
        .Caption = "(no object selected)": .Left = 10: .Top = 12: .Width = 370: .Height = 48
        .Font.Name = "Consolas": .Font.Size = 9
    End With

    AddCheckBox fraCrit, "chkObjectType", "Object Type", 10, 15, 170, True
    AddCheckBox fraCrit, "chkFullColor", "Full Color", 190, 15, 170, False
    AddCheckBox fraCrit, "chkOutlineColor", "Outline Color", 10, 30, 170, False
    AddCheckBox fraCrit, "chkOutlineWidth", "Outline Width", 190, 30, 170, False
    AddCheckBox fraCrit, "chkSameSize", "Same Size", 10, 45, 170, False
    AddCheckBox fraCrit, "chkNodeCount", "Node Count", 190, 45, 170, False

    Dim lblAdd As Object
    Set lblAdd = fraCrit.Controls.Add("Forms.Label.1", "lblAdditionalHeader")
    With lblAdd: .Caption = "Additional": .Left = 10: .Top = 62: .Width = 150: .Height = 12: .Font.Bold = True: End With

    AddCheckBox fraCrit, "chkWidthOnly", "Width Only", 10, 75, 170, False
    AddCheckBox fraCrit, "chkHeightOnly", "Height Only", 190, 75, 170, False
    AddCheckBox fraCrit, "chkAspectRatio", "Aspect Ratio", 10, 90, 170, False
    AddCheckBox fraCrit, "chkRotation", "Rotation Angle", 190, 90, 170, False
    AddCheckBox fraCrit, "chkTransparency", "Transparency", 10, 105, 170, False
    AddCheckBox fraCrit, "chkFountainFill", "Fountain Fill", 190, 105, 170, False
    AddCheckBox fraCrit, "chkPatternFill", "Pattern Fill", 10, 120, 170, False
    AddCheckBox fraCrit, "chkNoFillPresent", "No Fill / Fill Present", 190, 120, 170, False

    AddCheckBox fraExt, "chkTextFont", "Font", 10, 15, 170, False
    AddCheckBox fraExt, "chkTextSize", "Font Size", 190, 15, 170, False
    AddCheckBox fraExt, "chkTextContent", "Text Content", 10, 30, 170, False
    AddCheckBox fraExt, "chkTextStyle", "Text Style", 190, 30, 170, False

    AddCheckBox fraExt, "chkBitmapResolution", "Resolution", 10, 55, 170, False
    AddCheckBox fraExt, "chkBitmapWidth", "Bitmap Width (px)", 190, 55, 170, False
    AddCheckBox fraExt, "chkBitmapHeight", "Bitmap Height (px)", 10, 70, 170, False
    AddCheckBox fraExt, "chkBitmapColorMode", "Color Mode", 190, 70, 170, False

    AddOptionButton fraScope, "optCurrentPage", "Current Page", 10, 15, 170, True
    AddOptionButton fraScope, "optAllPages", "All Pages", 190, 15, 170, False
    AddOptionButton fraScope, "optCurrentLayer", "Current Layer", 10, 30, 170, False
    AddOptionButton fraScope, "optAllLayers", "All Layers", 190, 30, 170, False
    AddOptionButton fraScope, "optCurrentSelection", "Within Current Selection", 10, 43, 220, False

    AddCheckBox formDesigner, "chkSearchInsideGroups", "Search Inside Groups", 10, 402, 180, True
    AddCheckBox formDesigner, "chkIncludeReference", "Include Reference Object", 200, 402, 180, True
    AddCheckBox formDesigner, "chkIncludeLocked", "Include Locked Objects", 10, 417, 180, False
    AddCheckBox formDesigner, "chkIncludeHidden", "Include Hidden Objects", 200, 417, 180, False
    AddCheckBox formDesigner, "chkAddToExisting", "Add Matches To Existing Selection", 10, 432, 220, False

    AddButton formDesigner, "cmdExactMatch", "Exact Match", 8, 452, 72, 20
    AddButton formDesigner, "cmdAppearance", "Appearance", 84, 452, 72, 20
    AddButton formDesigner, "cmdGeometry", "Geometry", 160, 452, 72, 20
    AddButton formDesigner, "cmdShape", "Shape", 236, 452, 72, 20
    AddButton formDesigner, "cmdClearPresets", "Clear", 312, 452, 80, 20

    AddLabel formDesigner, "lblSizeTolLabel", "Size tol. (mm):", 8, 478, 70, 15
    AddTextBox formDesigner, "txtSizeTolerance", "0.001", 80, 476, 35, 18

    AddLabel formDesigner, "lblOutlineTolLabel", "Outline tol. (mm):", 125, 478, 80, 15
    AddTextBox formDesigner, "txtOutlineTolerance", "0.001", 210, 476, 35, 18

    AddLabel formDesigner, "lblRotTolLabel", "Rotation tol. (deg):", 255, 478, 90, 15
    AddTextBox formDesigner, "txtRotationTolerance", "0.1", 350, 476, 35, 18

    AddLabel formDesigner, "lblMatchCount", "Matches found: (not yet searched)", 8, 500, 380, 15, True

    AddButton formDesigner, "cmdFindMatches", "Find Matches", 8, 518, 100, 22
    Dim btnSelect As Object
    Set btnSelect = AddButton(formDesigner, "cmdSelectMatches", "Select Matches", 114, 518, 100, 22)
    btnSelect.Enabled = False

    AddButton formDesigner, "cmdClose", "Close", 312, 518, 80, 22

    InjectCodeBehind newFormComp.CodeModule
    MsgBox "Form built successfully!", vbInformation, "CorelDRAW Auto-Builder"
End Sub

Private Sub InjectCodeBehind(ByVal codeMod As Object)
    Dim s As String
    s = "Option Explicit" & vbCrLf
    s = s & "Public ReferenceShape As Shape" & vbCrLf
    s = s & "Private mRefProps As clsRefProperties" & vbCrLf
    s = s & "Private mMatches As Collection" & vbCrLf
    s = s & "Private mOriginalUnit As Long" & vbCrLf
    s = s & "Private mUnitChanged As Boolean" & vbCrLf
    s = s & "Private Sub UserForm_Initialize()" & vbCrLf
    s = s & "    mUnitChanged = False" & vbCrLf
    s = s & "    If ReferenceShape Is Nothing Then Set ReferenceShape = mdlMain.GetReferenceShapeForMacro()" & vbCrLf
    s = s & "    If ReferenceShape Is Nothing Then" & vbCrLf
    s = s & "        MsgBox ""Please select a single object in CorelDRAW first."", vbExclamation, ""Select Similar Objects""" & vbCrLf
    s = s & "        Unload Me: Exit Sub" & vbCrLf
    s = s & "    End If" & vbCrLf
    s = s & "    mOriginalUnit = mdlUnits.NormalizeToMillimeters(ActiveDocument)" & vbCrLf
    s = s & "    mUnitChanged = True" & vbCrLf
    s = s & "    Set mRefProps = New clsRefProperties" & vbCrLf
    s = s & "    mRefProps.PopulateFromShape ReferenceShape" & vbCrLf
    s = s & "    lblRefInfo.Caption = BuildReferenceInfoText(mRefProps)" & vbCrLf
    s = s & "    lblMatchCount.Caption = ""Matches found: (not yet searched)""" & vbCrLf
    s = s & "    cmdSelectMatches.Enabled = False" & vbCrLf
    s = s & "End Sub" & vbCrLf
    s = s & "Private Sub UserForm_QueryClose(Cancel As Integer, CloseMode As Integer)" & vbCrLf
    s = s & "    If mUnitChanged Then mdlUnits.RestoreUnit ActiveDocument, mOriginalUnit" & vbCrLf
    s = s & "End Sub" & vbCrLf
    s = s & "Private Function BuildReferenceInfoText(ref As clsRefProperties) As String" & vbCrLf
    s = s & "    Dim sText As String" & vbCrLf
    s = s & "    sText = ""Type:          "" & ref.TypeLabel & vbCrLf" & vbCrLf
    s = s & "    sText = sText & ""Size:          "" & Format(ref.Width, ""0.00"") & "" x "" & Format(ref.Height, ""0.00"") & "" mm"" & vbCrLf" & vbCrLf
    s = s & "    If ref.HasFill Then" & vbCrLf
    s = s & "        If ref.FillType = cdrUniformFill Then" & vbCrLf
    s = s & "            sText = sText & ""Fill:          Uniform, "" & ref.FillColor.Name(True) & vbCrLf" & vbCrLf
    s = s & "        Else" & vbCrLf
    s = s & "            sText = sText & ""Fill:          Special Fill"" & vbCrLf" & vbCrLf
    s = s & "        End If" & vbCrLf
    s = s & "    Else" & vbCrLf
    s = s & "        sText = sText & ""Fill:          None"" & vbCrLf" & vbCrLf
    s = s & "    End If" & vbCrLf
    s = s & "    If ref.HasOutline Then" & vbCrLf
    s = s & "        sText = sText & ""Outline:       "" & ref.OutlineColor.Name(True) & "" / "" & Format(ref.OutlineWidth, ""0.00"") & "" mm"" & vbCrLf" & vbCrLf
    s = s & "    Else" & vbCrLf
    s = s & "        sText = sText & ""Outline:       None"" & vbCrLf" & vbCrLf
    s = s & "    End If" & vbCrLf
    s = s & "    If ref.NodeCountValid Then" & vbCrLf
    s = s & "        sText = sText & ""Nodes:         "" & ref.NodeCount" & vbCrLf
    s = s & "    Else" & vbCrLf
    s = s & "        sText = sText & ""Nodes:         n/a""" & vbCrLf
    s = s & "    End If" & vbCrLf
    s = s & "    BuildReferenceInfoText = sText" & vbCrLf
    s = s & "End Function" & vbCrLf
    s = s & "Private Sub SetAllCriteria(ByVal State As Boolean)" & vbCrLf
    s = s & "    chkObjectType.Value = State: chkFullColor.Value = False" & vbCrLf
    s = s & "    chkOutlineColor.Value = False: chkOutlineWidth.Value = False" & vbCrLf
    s = s & "    chkSameSize.Value = False: chkNodeCount.Value = False" & vbCrLf
    s = s & "    chkWidthOnly.Value = False: chkHeightOnly.Value = False" & vbCrLf
    s = s & "    chkAspectRatio.Value = False: chkRotation.Value = False" & vbCrLf
    s = s & "    chkTransparency.Value = False: chkFountainFill.Value = False" & vbCrLf
    s = s & "    chkPatternFill.Value = False: chkNoFillPresent.Value = False" & vbCrLf
    s = s & "    chkTextFont.Value = False: chkTextSize.Value = False" & vbCrLf
    s = s & "    chkTextContent.Value = False: chkTextStyle.Value = False" & vbCrLf
    s = s & "    chkBitmapResolution.Value = False: chkBitmapWidth.Value = False" & vbCrLf
    s = s & "    chkBitmapHeight.Value = False: chkBitmapColorMode.Value = False" & vbCrLf
    s = s & "End Sub" & vbCrLf
    s = s & "Private Sub cmdExactMatch_Click(): SetAllCriteria False: chkObjectType.Value = True: chkFullColor.Value = True: chkOutlineColor.Value = True: chkOutlineWidth.Value = True: chkSameSize.Value = True: chkNodeCount.Value = True: End Sub" & vbCrLf
    s = s & "Private Sub cmdAppearance_Click(): SetAllCriteria False: chkFullColor.Value = True: chkOutlineColor.Value = True: chkOutlineWidth.Value = True: End Sub" & vbCrLf
    s = s & "Private Sub cmdClearPresets_Click(): SetAllCriteria False: End Sub" & vbCrLf
    s = s & "Private Function BuildCriteria() As TCriteria" & vbCrLf
    s = s & "    Dim c As TCriteria" & vbCrLf
    s = s & "    c.ObjectType = chkObjectType.Value: c.FullColor = chkFullColor.Value" & vbCrLf
    s = s & "    c.OutlineColor = chkOutlineColor.Value: c.OutlineWidth = chkOutlineWidth.Value" & vbCrLf
    s = s & "    c.SameSize = chkSameSize.Value: c.NodeCount = chkNodeCount.Value" & vbCrLf
    s = s & "    c.WidthOnly = chkWidthOnly.Value: c.HeightOnly = chkHeightOnly.Value" & vbCrLf
    s = s & "    c.AspectRatio = chkAspectRatio.Value: c.Rotation = chkRotation.Value" & vbCrLf
    s = s & "    c.Transparency = chkTransparency.Value: c.FountainFill = chkFountainFill.Value" & vbCrLf
    s = s & "    c.PatternFill = chkPatternFill.Value: c.NoFillPresent = chkNoFillPresent.Value" & vbCrLf
    s = s & "    c.TextFont = chkTextFont.Value: c.TextSize = chkTextSize.Value" & vbCrLf
    s = s & "    c.TextContent = chkTextContent.Value: c.TextStyle = chkTextStyle.Value" & vbCrLf
    s = s & "    c.BitmapResolution = chkBitmapResolution.Value: c.BitmapWidth = chkBitmapWidth.Value" & vbCrLf
    s = s & "    c.BitmapHeight = chkBitmapHeight.Value: c.BitmapColorMode = chkBitmapColorMode.Value" & vbCrLf
    s = s & "    c.SizeTolerance = SafeCDbl(txtSizeTolerance.Text, 0.001)" & vbCrLf
    s = s & "    c.OutlineTolerance = SafeCDbl(txtOutlineTolerance.Text, 0.001)" & vbCrLf
    s = s & "    c.RotationTolerance = SafeCDbl(txtRotationTolerance.Text, 0.1)" & vbCrLf
    s = s & "    BuildCriteria = c" & vbCrLf
    s = s & "End Function" & vbCrLf
    s = s & "Private Function SafeCDbl(ByVal ValText As String, ByVal DefaultVal As Double) As Double" & vbCrLf
    s = s & "    If IsNumeric(ValText) Then SafeCDbl = CDbl(ValText) Else SafeCDbl = DefaultVal" & vbCrLf
    s = s & "End Function" & vbCrLf
    s = s & "Private Function BuildScanOptions() As TScanOptions" & vbCrLf
    s = s & "    Dim o As TScanOptions" & vbCrLf
    s = s & "    If optCurrentPage.Value Then" & vbCrLf
    s = s & "        o.Scope = SCOPE_CURRENT_PAGE" & vbCrLf
    s = s & "    ElseIf optAllPages.Value Then" & vbCrLf
    s = s & "        o.Scope = SCOPE_ALL_PAGES" & vbCrLf
    s = s & "    Else" & vbCrLf
    s = s & "        o.Scope = SCOPE_CURRENT_SELECTION" & vbCrLf
    s = s & "    End If" & vbCrLf
    s = s & "    o.SearchInsideGroups = chkSearchInsideGroups.Value: o.IncludeLocked = chkIncludeLocked.Value" & vbCrLf
    s = s & "    o.IncludeHidden = chkIncludeHidden.Value: o.IncludeReference = chkIncludeReference.Value: o.AddToExisting = chkAddToExisting.Value" & vbCrLf
    s = s & "    BuildScanOptions = o" & vbCrLf
    s = s & "End Function" & vbCrLf
    s = s & "Private Sub cmdFindMatches_Click()" & vbCrLf
    s = s & "    Dim crit As TCriteria: crit = BuildCriteria()" & vbCrLf
    s = s & "    If Not mdlCompare.AtLeastOneCriterionSelected(crit) Then MsgBox ""Please select at least one criterion."", vbExclamation, ""Select Similar Objects"": Exit Sub" & vbCrLf
    s = s & "    Dim scanOpts As TScanOptions: scanOpts = BuildScanOptions()" & vbCrLf
    s = s & "    On Error Resume Next: ActiveDocument.BeginCommandGroup ""Select Similar Objects"": Optimization = True: On Error GoTo 0" & vbCrLf
    s = s & "    Dim candidates As Collection: Set candidates = mdlScan.CollectCandidates(scanOpts)" & vbCrLf
    s = s & "    Set mMatches = New Collection" & vbCrLf
    s = s & "    Dim scannedCount As Long, unsupportedCount As Long, sh As Shape, isMatch As Boolean" & vbCrLf
    s = s & "    For Each sh In candidates" & vbCrLf
    s = s & "        If Not (sh Is ReferenceShape) Then" & vbCrLf
    s = s & "            scannedCount = scannedCount + 1" & vbCrLf
    s = s & "            On Error Resume Next: Err.Clear" & vbCrLf
    s = s & "            isMatch = mdlCompare.IsSimilarObject(mRefProps, sh, crit)" & vbCrLf
    s = s & "            If Err.Number <> 0 Then unsupportedCount = unsupportedCount + 1: isMatch = False" & vbCrLf
    s = s & "            On Error GoTo 0" & vbCrLf
    s = s & "            If isMatch Then mMatches.Add sh" & vbCrLf
    s = s & "        End If" & vbCrLf
    s = s & "    Next sh" & vbCrLf
    s = s & "    On Error Resume Next: Optimization = False: ActiveWindow.Refresh: ActiveDocument.EndCommandGroup: On Error GoTo 0" & vbCrLf
    s = s & "    lblMatchCount.Caption = ""Matches found: "" & mMatches.Count & ""   (scanned "" & scannedCount & "", unsupported "" & unsupportedCount & "")""" & vbCrLf
    s = s & "    cmdSelectMatches.Enabled = (mMatches.Count > 0)" & vbCrLf
    s = s & "End Sub" & vbCrLf
    s = s & "Private Sub cmdSelectMatches_Click()" & vbCrLf
    s = s & "    If mMatches Is Nothing Then Exit Sub" & vbCrLf
    s = s & "    Dim scanOpts As TScanOptions: scanOpts = BuildScanOptions()" & vbCrLf
    s = s & "    Dim unlockedCount As Long: unlockedCount = mdlScan.SelectMatches(mMatches, ReferenceShape, scanOpts)" & vbCrLf
    s = s & "    If unlockedCount > 0 Then MsgBox unlockedCount & "" unlocked object(s) were selected."", vbInformation, ""Select Similar Objects""" & vbCrLf
    s = s & "End Sub" & vbCrLf
    s = s & "Private Sub cmdClose_Click(): Unload Me: End Sub"

    codeMod.DeleteLines 1, codeMod.CountOfLines
    codeMod.AddFromString s
End Sub

Private Function AddCheckBox(ByVal container As Object, ByVal cName As String, ByVal cCaption As String, _
                             ByVal cLeft As Single, ByVal cTop As Single, ByVal cWidth As Single, ByVal cVal As Boolean) As Object
    Set AddCheckBox = container.Controls.Add("Forms.CheckBox.1", cName)
    With AddCheckBox: .Caption = cCaption: .Left = cLeft: .Top = cTop: .Width = cWidth: .Height = 16: .Value = cVal: End With
End Function
Private Function AddOptionButton(ByVal container As Object, ByVal cName As String, ByVal cCaption As String, _
                                 ByVal cLeft As Single, ByVal cTop As Single, ByVal cWidth As Single, ByVal cVal As Boolean) As Object
    Set AddOptionButton = container.Controls.Add("Forms.OptionButton.1", cName)
    With AddOptionButton: .Caption = cCaption: .Left = cLeft: .Top = cTop: .Width = cWidth: .Height = 16: .Value = cVal: End With
End Function
Private Function AddButton(ByVal container As Object, ByVal cName As String, ByVal cCaption As String, _
                           ByVal cLeft As Single, ByVal cTop As Single, ByVal cWidth As Single, ByVal cHeight As Single) As Object
    Set AddButton = container.Controls.Add("Forms.CommandButton.1", cName)
    With AddButton: .Caption = cCaption: .Left = cLeft: .Top = cTop: .Width = cWidth: .Height = cHeight: End With
End Function
Private Function AddTextBox(ByVal container As Object, ByVal cName As String, ByVal cText As String, _
                            ByVal cLeft As Single, ByVal cTop As Single, ByVal cWidth As Single, ByVal cHeight As Single) As Object
    Set AddTextBox = container.Controls.Add("Forms.TextBox.1", cName)
    With AddTextBox: .Text = cText: .Left = cLeft: .Top = cTop: .Width = cWidth: .Height = cHeight: End With
End Function
Private Function AddLabel(ByVal container As Object, ByVal cName As String, ByVal cCaption As String, _
                           ByVal cLeft As Single, ByVal cTop As Single, ByVal cWidth As Single, ByVal cHeight As Single, _
                           Optional ByVal isBold As Boolean = False) As Object
    Set AddLabel = container.Controls.Add("Forms.Label.1", cName)
    With AddLabel: .Caption = cCaption: .Left = cLeft: .Top = cTop: .Width = cWidth: .Height = cHeight: .Font.Bold = isBold: End With
End Function
