Attribute VB_Name = "mdlFormBuilder"
'==============================================================
' mdlFormBuilder (VittixDimensionTools)
'
' Builds frmDimension as a real MSForms VBA UserForm at
' runtime via the VBE Extensibility model. This replaces importing
' the legacy VB6-format .frm file, which CorelDRAW's VBE cannot
' load reliably.
'
' The injected code-behind mirrors the form logic in modSettings.
'==============================================================
Option Explicit

Private Const FORM_NAME As String = "frmDimension"

Public Sub BuildDimensionFormWithCode()

    Dim vbProj As Object
    Set vbProj = Application.VBE.ActiveVBProject

    ' 1. Remove any existing/broken form to prevent duplicates
    On Error Resume Next
    Dim existingComp As Object
    Set existingComp = vbProj.VBComponents(FORM_NAME)
    If Not existingComp Is Nothing Then
        vbProj.VBComponents.Remove existingComp
    End If
    Err.Clear
    On Error GoTo 0

    ' 2. Add the UserForm component (3 = MSForm)
    Dim newFormComp As Object
    Set newFormComp = vbProj.VBComponents.Add(3)
    newFormComp.Name = FORM_NAME

    ' 3. Configure the form itself
    With newFormComp.Properties
        .Item("Caption").Value = "Vittix Dimension"
        .Item("Width").Value = 320
        .Item("Height").Value = 520
        .Item("StartUpPosition").Value = 1 ' CenterOwner
    End With

    Dim d As Object
    Set d = newFormComp.Designer

    ' 4. Add UI controls (left, top, width, height in points)
    Dim yPos As Single
    yPos = 16
    
    ' Unit
    AddLabel d, "lblUnit", "Unit", 16, yPos, 100, 17
    AddComboList d, "cmbUnit", 120, yPos - 3, 160, 20
    yPos = yPos + 30

    ' Decimals
    AddLabel d, "lblDecimals", "Decimals", 16, yPos, 100, 17
    AddComboList d, "cmbDecimals", 120, yPos - 3, 160, 20
    yPos = yPos + 30

    ' Position
    AddLabel d, "lblPosition", "Position", 16, yPos, 100, 17
    AddComboList d, "cmbPosition", 120, yPos - 3, 160, 20
    yPos = yPos + 30

    ' Font Size
    AddLabel d, "lblFontSize", "Font Size (pt)", 16, yPos, 100, 17
    AddTextBox d, "txtFontSize", "", 120, yPos, 160, 20
    yPos = yPos + 30

    ' Text Width %
    AddLabel d, "lblTextWidthPercent", "Text Width %", 16, yPos, 100, 17
    AddTextBox d, "txtTextWidthPercent", "", 120, yPos, 160, 20
    yPos = yPos + 30

    ' Gap
    AddLabel d, "lblGap", "Gap (mm)", 16, yPos, 100, 17
    AddTextBox d, "txtGap", "", 120, yPos, 160, 20
    yPos = yPos + 30

    ' Padding
    AddLabel d, "lblPadding", "Padding (mm)", 16, yPos, 100, 17
    AddTextBox d, "txtPadding", "", 120, yPos, 160, 20
    yPos = yPos + 30

    ' Corner Radius
    AddLabel d, "lblCornerRadius", "Corner Radius (mm)", 16, yPos, 130, 17
    AddTextBox d, "txtCornerRadius", "", 150, yPos, 130, 20
    yPos = yPos + 35

    ' Checkboxes column 1
    AddCheckBox d, "chkShowWidth", "Show Width", 16, yPos, 130, True
    AddCheckBox d, "chkShowHeight", "Show Height", 16, yPos + 22, 130, True
    AddCheckBox d, "chkShowArea", "Show Area", 16, yPos + 44, 130, False
    AddCheckBox d, "chkShowPerimeter", "Show Perimeter", 16, yPos + 66, 130, False

    ' Checkboxes column 2
    AddCheckBox d, "chkShowObjectCount", "Show Object Count", 150, yPos, 150, False
    AddCheckBox d, "chkBackgroundBox", "Background Box", 150, yPos + 22, 150, True
    AddCheckBox d, "chkRoundedBackground", "Rounded Background", 150, yPos + 44, 150, False
    AddCheckBox d, "chkCreateLayer", "Create Layer", 150, yPos + 66, 150, True
    
    yPos = yPos + 90
    
    AddCheckBox d, "chkRememberSettings", "Remember Settings", 16, yPos, 200, True
    yPos = yPos + 30

    ' Template
    AddLabel d, "lblTemplate", "Template", 16, yPos, 100, 17
    AddTextBox d, "txtTemplate", "", 120, yPos, 160, 20
    yPos = yPos + 35

    ' Buttons
    AddButton d, "cmdOK", "OK", 60, yPos, 90, 25
    AddButton d, "cmdCancel", "Cancel", 160, yPos, 90, 25

    ' 5. Inject code-behind
    InjectCodeBehind newFormComp.CodeModule

End Sub

' Returns True if a form component with the expected name exists.
Public Function DimensionFormExists() As Boolean
    On Error Resume Next
    DimensionFormExists = Not Application.VBE.ActiveVBProject.VBComponents(FORM_NAME) Is Nothing
    Err.Clear
    On Error GoTo 0
End Function

'--------------------------------------------------------------
' Control helpers
'--------------------------------------------------------------
Private Function AddComboList(ByVal container As Object, ByVal cName As String, _
                              ByVal cLeft As Single, ByVal cTop As Single, _
                              ByVal cWidth As Single, ByVal cHeight As Single) As Object
    Set AddComboList = container.Controls.Add("Forms.ComboBox.1", cName)
    With AddComboList
        .Left = cLeft: .Top = cTop: .Width = cWidth: .Height = cHeight
        .Style = 2 ' fmStyleDropDownList
    End With
End Function

Private Function AddCheckBox(ByVal container As Object, ByVal cName As String, ByVal cCaption As String, _
                             ByVal cLeft As Single, ByVal cTop As Single, ByVal cWidth As Single, _
                             ByVal cVal As Boolean) As Object
    Set AddCheckBox = container.Controls.Add("Forms.CheckBox.1", cName)
    With AddCheckBox: .Caption = cCaption: .Left = cLeft: .Top = cTop: .Width = cWidth: .Height = 17: .Value = cVal: End With
End Function

Private Function AddButton(ByVal container As Object, ByVal cName As String, ByVal cCaption As String, _
                           ByVal cLeft As Single, ByVal cTop As Single, ByVal cWidth As Single, _
                           ByVal cHeight As Single) As Object
    Set AddButton = container.Controls.Add("Forms.CommandButton.1", cName)
    With AddButton: .Caption = cCaption: .Left = cLeft: .Top = cTop: .Width = cWidth: .Height = cHeight: End With
End Function

Private Function AddTextBox(ByVal container As Object, ByVal cName As String, ByVal cText As String, _
                            ByVal cLeft As Single, ByVal cTop As Single, ByVal cWidth As Single, _
                            ByVal cHeight As Single) As Object
    Set AddTextBox = container.Controls.Add("Forms.TextBox.1", cName)
    With AddTextBox: .Text = cText: .Left = cLeft: .Top = cTop: .Width = cWidth: .Height = cHeight: End With
End Function

Private Function AddLabel(ByVal container As Object, ByVal cName As String, ByVal cCaption As String, _
                          ByVal cLeft As Single, ByVal cTop As Single, ByVal cWidth As Single, _
                          ByVal cHeight As Single) As Object
    Set AddLabel = container.Controls.Add("Forms.Label.1", cName)
    With AddLabel: .Caption = cCaption: .Left = cLeft: .Top = cTop: .Width = cWidth: .Height = cHeight: End With
End Function

'--------------------------------------------------------------
' Code-behind injection
'--------------------------------------------------------------
Private Sub InjectCodeBehind(ByVal codeMod As Object)
    Dim s As String
    s = "Option Explicit" & vbCrLf
    s = s & "" & vbCrLf
    s = s & "Private Sub UserForm_Initialize()" & vbCrLf
    s = s & "" & vbCrLf
    s = s & "    If Not g_SettingsInitialized Then LoadSettings" & vbCrLf
    s = s & "" & vbCrLf
    s = s & "    cmbUnit.AddItem ""MM""" & vbCrLf
    s = s & "    cmbUnit.AddItem ""CM""" & vbCrLf
    s = s & "    cmbUnit.AddItem ""M""" & vbCrLf
    s = s & "    cmbUnit.AddItem ""IN""" & vbCrLf
    s = s & "    cmbUnit.AddItem ""FT""" & vbCrLf
    s = s & "    cmbUnit.ListIndex = UnitToComboIndex(gSettings.unit)" & vbCrLf
    s = s & "" & vbCrLf
    s = s & "    cmbDecimals.AddItem ""0""" & vbCrLf
    s = s & "    cmbDecimals.AddItem ""1""" & vbCrLf
    s = s & "    cmbDecimals.AddItem ""2""" & vbCrLf
    s = s & "    cmbDecimals.AddItem ""3""" & vbCrLf
    s = s & "    cmbDecimals.AddItem ""4""" & vbCrLf
    s = s & "    cmbDecimals.ListIndex = gSettings.decimals" & vbCrLf
    s = s & "" & vbCrLf
    s = s & "    cmbPosition.AddItem ""Auto""" & vbCrLf
    s = s & "    cmbPosition.AddItem ""Above""" & vbCrLf
    s = s & "    cmbPosition.AddItem ""Below""" & vbCrLf
    s = s & "    cmbPosition.AddItem ""Left""" & vbCrLf
    s = s & "    cmbPosition.AddItem ""Right""" & vbCrLf
    s = s & "    cmbPosition.ListIndex = PositionToIndex(gSettings.Position)" & vbCrLf
    s = s & "" & vbCrLf
    s = s & "    txtFontSize.Text = gSettings.fontSize" & vbCrLf
    s = s & "    txtTextWidthPercent.Text = gSettings.TextWidthPercent" & vbCrLf
    s = s & "    txtGap.Text = gSettings.Gap" & vbCrLf
    s = s & "    txtPadding.Text = gSettings.Padding" & vbCrLf
    s = s & "    txtCornerRadius.Text = gSettings.CornerRadius" & vbCrLf
    s = s & "" & vbCrLf
    s = s & "    chkShowWidth.Value = IIf(gSettings.ShowWidth, 1, 0)" & vbCrLf
    s = s & "    chkShowHeight.Value = IIf(gSettings.ShowHeight, 1, 0)" & vbCrLf
    s = s & "    chkShowArea.Value = IIf(gSettings.ShowArea, 1, 0)" & vbCrLf
    s = s & "    chkShowPerimeter.Value = IIf(gSettings.ShowPerimeter, 1, 0)" & vbCrLf
    s = s & "    chkShowObjectCount.Value = IIf(gSettings.ShowObjectCount, 1, 0)" & vbCrLf
    s = s & "    chkBackgroundBox.Value = IIf(gSettings.BackgroundBox, 1, 0)" & vbCrLf
    s = s & "    chkRoundedBackground.Value = IIf(gSettings.RoundedBackground, 1, 0)" & vbCrLf
    s = s & "    chkCreateLayer.Value = IIf(gSettings.CreateLayer, 1, 0)" & vbCrLf
    s = s & "    chkRememberSettings.Value = IIf(gSettings.RememberSettings, 1, 0)" & vbCrLf
    s = s & "" & vbCrLf
    s = s & "    txtTemplate.Text = gSettings.Template" & vbCrLf
    s = s & "" & vbCrLf
    s = s & "End Sub" & vbCrLf
    s = s & "" & vbCrLf
    s = s & "Private Sub cmdCancel_Click()" & vbCrLf
    s = s & "    Unload Me" & vbCrLf
    s = s & "End Sub" & vbCrLf
    s = s & "" & vbCrLf
    s = s & "Private Sub cmdOK_Click()" & vbCrLf
    s = s & "    On Error GoTo ErrHandler" & vbCrLf
    s = s & "    ReadSettingsFromForm Me, gSettings" & vbCrLf
    s = s & "    SaveSettings" & vbCrLf
    s = s & "    VittixDimensionTools_Create" & vbCrLf
    s = s & "    Unload Me" & vbCrLf
    s = s & "    Exit Sub" & vbCrLf
    s = s & "ErrHandler:" & vbCrLf
    s = s & "    LogError ""frmDimension.cmdOK_Click"", Err.Number, Err.Description" & vbCrLf
    s = s & "End Sub" & vbCrLf
    s = s & "" & vbCrLf
    s = s & "Private Sub UserForm_QueryClose(Cancel As Integer, CloseMode As Integer)" & vbCrLf
    s = s & "    If CloseMode = vbFormControlMenu Then" & vbCrLf
    s = s & "        Cancel = True" & vbCrLf
    s = s & "        Unload Me" & vbCrLf
    s = s & "    End If" & vbCrLf
    s = s & "End Sub" & vbCrLf
    s = s & "" & vbCrLf
    s = s & "'--------------------------------------------------------------" & vbCrLf
    s = s & "' Helper functions" & vbCrLf
    s = s & "'--------------------------------------------------------------" & vbCrLf
    s = s & "Private Function UnitToComboIndex(u As VDTUnit) As Integer" & vbCrLf
    s = s & "    Select Case u" & vbCrLf
    s = s & "        Case UNIT_MM: UnitToComboIndex = 0" & vbCrLf
    s = s & "        Case UNIT_CM: UnitToComboIndex = 1" & vbCrLf
    s = s & "        Case UNIT_M:  UnitToComboIndex = 2" & vbCrLf
    s = s & "        Case UNIT_IN: UnitToComboIndex = 3" & vbCrLf
    s = s & "        Case UNIT_FT: UnitToComboIndex = 4" & vbCrLf
    s = s & "        Case Else:    UnitToComboIndex = 0" & vbCrLf
    s = s & "    End Select" & vbCrLf
    s = s & "End Function" & vbCrLf
    s = s & "" & vbCrLf
    s = s & "Private Function ComboIndexToUnit(idx As Integer) As VDTUnit" & vbCrLf
    s = s & "    Select Case idx" & vbCrLf
    s = s & "        Case 0: ComboIndexToUnit = UNIT_MM" & vbCrLf
    s = s & "        Case 1: ComboIndexToUnit = UNIT_CM" & vbCrLf
    s = s & "        Case 2: ComboIndexToUnit = UNIT_M" & vbCrLf
    s = s & "        Case 3: ComboIndexToUnit = UNIT_IN" & vbCrLf
    s = s & "        Case 4: ComboIndexToUnit = UNIT_FT" & vbCrLf
    s = s & "        Case Else: ComboIndexToUnit = UNIT_MM" & vbCrLf
    s = s & "    End Select" & vbCrLf
    s = s & "End Function" & vbCrLf
    s = s & "" & vbCrLf
    s = s & "Private Function PositionToIndex(pos As String) As Integer" & vbCrLf
    s = s & "    Select Case LCase$(pos)" & vbCrLf
    s = s & "        Case ""auto"":    PositionToIndex = 0" & vbCrLf
    s = s & "        Case ""above"":   PositionToIndex = 1" & vbCrLf
    s = s & "        Case ""below"":   PositionToIndex = 2" & vbCrLf
    s = s & "        Case ""left"":    PositionToIndex = 3" & vbCrLf
    s = s & "        Case ""right"":   PositionToIndex = 4" & vbCrLf
    s = s & "        Case Else:       PositionToIndex = 0" & vbCrLf
    s = s & "    End Select" & vbCrLf
    s = s & "End Function" & vbCrLf
    s = s & "" & vbCrLf
    s = s & "Private Function IndexToPosition(idx As Integer) As String" & vbCrLf
    s = s & "    Select Case idx" & vbCrLf
    s = s & "        Case 0: IndexToPosition = ""Auto""" & vbCrLf
    s = s & "        Case 1: IndexToPosition = ""Above""" & vbCrLf
    s = s & "        Case 2: IndexToPosition = ""Below""" & vbCrLf
    s = s & "        Case 3: IndexToPosition = ""Left""" & vbCrLf
    s = s & "        Case 4: IndexToPosition = ""Right""" & vbCrLf
    s = s & "        Case Else: IndexToPosition = ""Auto""" & vbCrLf
    s = s & "    End Select" & vbCrLf
    s = s & "End Function" & vbCrLf

    codeMod.DeleteLines 1, codeMod.CountOfLines
    codeMod.AddFromString s
End Sub